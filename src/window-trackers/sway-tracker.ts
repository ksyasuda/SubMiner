/*
  SubMiner - All-in-one sentence mining overlay
  Copyright (C) 2024 sudacode

  This program is free software: you can redistribute it and/or modify
  it under the terms of the GNU General Public License as published by
  the Free Software Foundation, either version 3 of the License, or
  (at your option) any later version.

  This program is distributed in the hope that it will be useful,
  but WITHOUT ANY WARRANTY; without even the implied warranty of
  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
  GNU General Public License for more details.

  You should have received a copy of the GNU General Public License
  along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

import { BaseWindowTracker } from './base-tracker';
import {
  createCoalescedTask,
  execFileWithTimeout,
  ProcessCommandLineCache,
  type CommandRunner,
} from './linux-poll';

interface SwayRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface SwayNode {
  pid?: number;
  app_id?: string;
  window_properties?: { class?: string };
  rect?: SwayRect;
  nodes?: SwayNode[];
  floating_nodes?: SwayNode[];
}

export class SwayWindowTracker extends BaseWindowTracker {
  private pollInterval: ReturnType<typeof setInterval> | null = null;
  private readonly targetMpvSocketPath: string | null;
  private readonly runCommand: CommandRunner;
  private readonly commandLines: ProcessCommandLineCache;
  private stopped = false;
  private readonly pollGeometry = createCoalescedTask(() => this.pollGeometryOnce());

  constructor(targetMpvSocketPath?: string, runCommand: CommandRunner = execFileWithTimeout) {
    super();
    this.targetMpvSocketPath = targetMpvSocketPath?.trim() || null;
    this.runCommand = runCommand;
    this.commandLines = new ProcessCommandLineCache(runCommand);
  }

  start(): void {
    this.stopped = false;
    this.pollInterval = setInterval(() => void this.pollGeometry(), 250);
    void this.pollGeometry();
  }

  stop(): void {
    this.stopped = true;
    if (this.pollInterval) {
      clearInterval(this.pollInterval);
      this.pollInterval = null;
    }
  }

  override refreshNow(): Promise<void> {
    return this.pollGeometry();
  }

  private collectMpvWindows(node: SwayNode): SwayNode[] {
    const windows: SwayNode[] = [];
    if (node.app_id === 'mpv' || node.window_properties?.class === 'mpv') {
      windows.push(node);
    }

    if (node.nodes) {
      for (const child of node.nodes) {
        windows.push(...this.collectMpvWindows(child));
      }
    }

    if (node.floating_nodes) {
      for (const child of node.floating_nodes) {
        windows.push(...this.collectMpvWindows(child));
      }
    }

    return windows;
  }

  private async findTargetSocketWindow(node: SwayNode): Promise<SwayNode | null> {
    const windows = this.collectMpvWindows(node);
    if (!this.targetMpvSocketPath) {
      return windows[0] || null;
    }

    const commandLines = await this.commandLines.lookup(
      windows.flatMap((window) => (window.pid ? [window.pid] : [])),
    );
    return (
      windows.find((candidate) => {
        const commandLine = candidate.pid ? commandLines.get(candidate.pid) : undefined;
        return commandLine !== undefined && this.matchesTargetSocket(commandLine);
      }) || null
    );
  }

  private matchesTargetSocket(commandLine: string): boolean {
    return (
      commandLine.includes(`--input-ipc-server=${this.targetMpvSocketPath}`) ||
      commandLine.includes(`--input-ipc-server ${this.targetMpvSocketPath}`)
    );
  }

  private async pollGeometryOnce(): Promise<void> {
    if (this.stopped) {
      return;
    }

    let mpvWindow: SwayNode | null;
    try {
      const tree = JSON.parse(await this.runCommand('swaymsg', ['-t', 'get_tree'])) as SwayNode;
      mpvWindow = await this.findTargetSocketWindow(tree);
    } catch {
      // swaymsg/ps not available, timed out, or failed - keep the last geometry
      return;
    }
    if (this.stopped) {
      return;
    }

    if (mpvWindow && mpvWindow.rect) {
      this.updateGeometry({
        x: mpvWindow.rect.x,
        y: mpvWindow.rect.y,
        width: mpvWindow.rect.width,
        height: mpvWindow.rect.height,
      });
    } else {
      this.updateGeometry(null);
    }
  }
}
