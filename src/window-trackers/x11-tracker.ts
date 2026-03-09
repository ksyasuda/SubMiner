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

import { execFile } from 'child_process';
import { BaseWindowTracker } from './base-tracker';

type CommandRunner = (command: string, args: string[]) => Promise<string>;

function execFileUtf8(command: string, args: string[]): Promise<string> {
  return new Promise((resolve, reject) => {
    execFile(command, args, { encoding: 'utf-8' }, (error, stdout) => {
      if (error) {
        reject(error);
        return;
      }
      resolve(stdout);
    });
  });
}

export function parseX11WindowGeometry(winInfo: string): {
  x: number;
  y: number;
  width: number;
  height: number;
} | null {
  const xMatch = winInfo.match(/Absolute upper-left X:\s*(-?\d+)/);
  const yMatch = winInfo.match(/Absolute upper-left Y:\s*(-?\d+)/);
  const widthMatch = winInfo.match(/Width:\s*(\d+)/);
  const heightMatch = winInfo.match(/Height:\s*(\d+)/);
  if (!xMatch || !yMatch || !widthMatch || !heightMatch) {
    return null;
  }
  return {
    x: parseInt(xMatch[1]!, 10),
    y: parseInt(yMatch[1]!, 10),
    width: parseInt(widthMatch[1]!, 10),
    height: parseInt(heightMatch[1]!, 10),
  };
}

export function parseX11WindowPid(raw: string): number | null {
  const pidMatch = raw.match(/= (\d+)/);
  if (!pidMatch) {
    return null;
  }
  const pid = Number.parseInt(pidMatch[1]!, 10);
  return Number.isInteger(pid) ? pid : null;
}

export class X11WindowTracker extends BaseWindowTracker {
  private pollInterval: ReturnType<typeof setInterval> | null = null;
  private readonly targetMpvSocketPath: string | null;
  private readonly runCommand: CommandRunner;
  private pollInFlight = false;
  private currentPollIntervalMs = 750;
  private readonly stablePollIntervalMs = 250;

  constructor(targetMpvSocketPath?: string, runCommand: CommandRunner = execFileUtf8) {
    super();
    this.targetMpvSocketPath = targetMpvSocketPath?.trim() || null;
    this.runCommand = runCommand;
  }

  start(): void {
    this.resetPollInterval(this.currentPollIntervalMs);
    this.pollGeometry();
  }

  stop(): void {
    if (this.pollInterval) {
      clearInterval(this.pollInterval);
      this.pollInterval = null;
    }
  }

  private resetPollInterval(intervalMs: number): void {
    if (this.pollInterval) {
      clearInterval(this.pollInterval);
      this.pollInterval = null;
    }
    this.pollInterval = setInterval(() => this.pollGeometry(), intervalMs);
  }

  private pollGeometry(): void {
    if (this.pollInFlight) {
      return;
    }
    this.pollInFlight = true;
    void this.pollGeometryAsync()
      .catch(() => {
        this.updateGeometry(null);
      })
      .finally(() => {
        this.pollInFlight = false;
      });
  }

  private async pollGeometryAsync(): Promise<void> {
    const windowIdsOutput = await this.runCommand('xdotool', [
      'search',
      '--onlyvisible',
      '--class',
      'mpv',
    ]);
    const windowIds = windowIdsOutput.trim();
    if (!windowIds) {
      this.updateGeometry(null);
      return;
    }

    const windowIdList = windowIds.split(/\s+/).filter(Boolean);
    if (windowIdList.length === 0) {
      this.updateGeometry(null);
      return;
    }

    const windowId = await this.findTargetWindowId(windowIdList);
    if (!windowId) {
      this.updateGeometry(null);
      return;
    }

    const winInfo = await this.runCommand('xwininfo', ['-id', windowId]);
    const geometry = parseX11WindowGeometry(winInfo);
    if (!geometry) {
      this.updateGeometry(null);
      return;
    }

    this.updateGeometry(geometry);
    if (this.pollInterval && this.currentPollIntervalMs !== this.stablePollIntervalMs) {
      this.currentPollIntervalMs = this.stablePollIntervalMs;
      this.resetPollInterval(this.currentPollIntervalMs);
    }
  }

  private async findTargetWindowId(windowIds: string[]): Promise<string | null> {
    if (!this.targetMpvSocketPath) {
      return windowIds[0] ?? null;
    }

    for (const windowId of windowIds) {
      if (await this.isWindowForTargetSocket(windowId)) {
        return windowId;
      }
    }

    return null;
  }

  private async isWindowForTargetSocket(windowId: string): Promise<boolean> {
    const pid = await this.getWindowPid(windowId);
    if (pid === null) {
      return false;
    }

    const commandLine = await this.getWindowCommandLine(pid);
    if (!commandLine) {
      return false;
    }

    return (
      commandLine.includes(`--input-ipc-server=${this.targetMpvSocketPath}`) ||
      commandLine.includes(`--input-ipc-server ${this.targetMpvSocketPath}`)
    );
  }

  private async getWindowPid(windowId: string): Promise<number | null> {
    let windowPid: string;
    try {
      windowPid = await this.runCommand('xprop', ['-id', windowId, '_NET_WM_PID']);
    } catch {
      return null;
    }
    return parseX11WindowPid(windowPid);
  }

  private async getWindowCommandLine(pid: number): Promise<string | null> {
    let raw: string;
    try {
      raw = await this.runCommand('ps', ['-p', String(pid), '-o', 'args=']);
    } catch {
      return null;
    }
    const commandLine = raw.trim();
    return commandLine || null;
  }
}
