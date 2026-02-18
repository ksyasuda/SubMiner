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

import { execSync } from 'child_process';
import { BaseWindowTracker } from './base-tracker';

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

  constructor(targetMpvSocketPath?: string) {
    super();
    this.targetMpvSocketPath = targetMpvSocketPath?.trim() || null;
  }

  start(): void {
    this.pollInterval = setInterval(() => this.pollGeometry(), 250);
    this.pollGeometry();
  }

  stop(): void {
    if (this.pollInterval) {
      clearInterval(this.pollInterval);
      this.pollInterval = null;
    }
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

  private findTargetSocketWindow(node: SwayNode): SwayNode | null {
    const windows = this.collectMpvWindows(node);
    if (!this.targetMpvSocketPath) {
      return windows[0] || null;
    }

    return windows.find((candidate) => this.isWindowForTargetSocket(candidate)) || null;
  }

  private isWindowForTargetSocket(node: SwayNode): boolean {
    if (!node.pid) {
      return false;
    }

    const commandLine = this.getWindowCommandLine(node.pid);
    if (!commandLine) {
      return false;
    }

    return (
      commandLine.includes(`--input-ipc-server=${this.targetMpvSocketPath}`) ||
      commandLine.includes(`--input-ipc-server ${this.targetMpvSocketPath}`)
    );
  }

  private getWindowCommandLine(pid: number): string | null {
    const commandLine = execSync(`ps -p ${pid} -o args=`, {
      encoding: 'utf-8',
    }).trim();
    return commandLine || null;
  }

  private pollGeometry(): void {
    try {
      const output = execSync('swaymsg -t get_tree', { encoding: 'utf-8' });
      const tree: SwayNode = JSON.parse(output);
      const mpvWindow = this.findTargetSocketWindow(tree);

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
    } catch (err) {
      // swaymsg not available or failed - silent fail
    }
  }
}
