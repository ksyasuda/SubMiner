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
export type ScreenToDipPoint = (point: { x: number; y: number }) => { x: number; y: number };

const preservePoint: ScreenToDipPoint = (point) => point;

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

export function normalizeX11WindowId(raw: string): string | null {
  const trimmed = raw.trim();
  if (!trimmed) {
    return null;
  }
  try {
    return BigInt(trimmed).toString();
  } catch {
    return null;
  }
}

export function parseX11RootActiveWindowId(raw: string): string | null {
  const match = raw.match(/window id #\s*(\S+)/i);
  if (!match) {
    return null;
  }
  return normalizeX11WindowId(match[1]!);
}

export class X11WindowTracker extends BaseWindowTracker {
  private pollInterval: ReturnType<typeof setInterval> | null = null;
  private readonly targetMpvSocketPath: string | null;
  private readonly runCommand: CommandRunner;
  private readonly screenToDipPoint: ScreenToDipPoint;
  private targetWindowId: string | null = null;
  private targetWindowPid: number | null = null;
  private pollInFlight = false;
  private currentPollIntervalMs = 750;
  private readonly stablePollIntervalMs = 250;

  constructor(
    targetMpvSocketPath?: string,
    runCommand: CommandRunner = execFileUtf8,
    screenToDipPoint: ScreenToDipPoint = preservePoint,
  ) {
    super();
    this.targetMpvSocketPath = targetMpvSocketPath?.trim() || null;
    this.runCommand = runCommand;
    this.screenToDipPoint = screenToDipPoint;
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

  override getTargetWindowMediaSourceId(): string | null {
    const normalizedWindowId = this.targetWindowId
      ? normalizeX11WindowId(this.targetWindowId)
      : null;
    return normalizedWindowId ? `window:${normalizedWindowId}:0` : null;
  }

  override getTargetWindowNativeId(): string | null {
    return this.targetWindowId ? normalizeX11WindowId(this.targetWindowId) : null;
  }

  override async raiseTargetWindow(): Promise<boolean> {
    const targetWindowId = this.targetWindowId;
    if (!targetWindowId) {
      return false;
    }
    let raised = false;
    try {
      await this.runCommand('xdotool', ['windowactivate', targetWindowId]);
      raised = true;
    } catch {
      // Some WMs reject activation but accept a plain restack below.
    }
    try {
      await this.runCommand('xdotool', ['windowraise', targetWindowId]);
      raised = true;
    } catch {
      // Keep any successful activation result.
    }
    return raised;
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
      this.targetWindowId = null;
      this.targetWindowPid = null;
      this.updateGeometry(null);
      return;
    }
    this.targetWindowId = windowId;
    const targetPid = this.targetWindowPid ?? (await this.getWindowPid(windowId));
    this.targetWindowPid = targetPid;

    const winInfo = await this.runCommand('xwininfo', ['-id', windowId]);
    const physicalGeometry = parseX11WindowGeometry(winInfo);
    if (!physicalGeometry) {
      this.updateGeometry(null);
      return;
    }
    const topLeft = this.screenToDipPoint({
      x: physicalGeometry.x,
      y: physicalGeometry.y,
    });
    const bottomRight = this.screenToDipPoint({
      x: physicalGeometry.x + physicalGeometry.width,
      y: physicalGeometry.y + physicalGeometry.height,
    });
    const geometry = {
      x: topLeft.x,
      y: topLeft.y,
      width: bottomRight.x - topLeft.x,
      height: bottomRight.y - topLeft.y,
    };

    const focused = await this.isWindowActive(windowId, targetPid);
    this.updateGeometry(geometry, focused);
    this.updateTargetWindowFocused(focused);
    if (this.pollInterval && this.currentPollIntervalMs !== this.stablePollIntervalMs) {
      this.currentPollIntervalMs = this.stablePollIntervalMs;
      this.resetPollInterval(this.currentPollIntervalMs);
    }
  }

  private async findTargetWindowId(windowIds: string[]): Promise<string | null> {
    this.targetWindowId = null;
    this.targetWindowPid = null;
    if (!this.targetMpvSocketPath) {
      const windowId = windowIds[0] ?? null;
      if (windowId) {
        this.targetWindowPid = await this.getWindowPid(windowId);
      }
      return windowId;
    }

    for (const windowId of windowIds) {
      const pid = await this.getTargetSocketWindowPid(windowId);
      if (pid !== null) {
        this.targetWindowPid = pid;
        return windowId;
      }
    }

    return null;
  }

  private async getTargetSocketWindowPid(windowId: string): Promise<number | null> {
    const pid = await this.getWindowPid(windowId);
    if (pid === null) {
      return null;
    }

    const commandLine = await this.getWindowCommandLine(pid);
    if (!commandLine) {
      return null;
    }

    const matchesTargetSocket =
      commandLine.includes(`--input-ipc-server=${this.targetMpvSocketPath}`) ||
      commandLine.includes(`--input-ipc-server ${this.targetMpvSocketPath}`);
    return matchesTargetSocket ? pid : null;
  }

  private async isWindowActive(windowId: string, targetPid: number | null): Promise<boolean> {
    const activeWindowId = await this.getX11ActiveWindowId();
    if (!activeWindowId) {
      return true;
    }
    const normalizedTarget = normalizeX11WindowId(windowId);
    const normalizedActive = normalizeX11WindowId(activeWindowId);
    if (!normalizedTarget || !normalizedActive) {
      return true;
    }
    if (targetPid !== null) {
      const activePid = await this.getWindowPid(normalizedActive);
      if (activePid !== null) {
        return activePid === targetPid;
      }
    }
    return normalizedTarget === normalizedActive;
  }

  private async getX11ActiveWindowId(): Promise<string | null> {
    try {
      const rootActiveWindow = parseX11RootActiveWindowId(
        await this.runCommand('xprop', ['-root', '_NET_ACTIVE_WINDOW']),
      );
      if (rootActiveWindow) {
        return rootActiveWindow;
      }
    } catch {
      // Fall back below. Some minimal WMs do not expose _NET_ACTIVE_WINDOW.
    }
    try {
      return normalizeX11WindowId(await this.runCommand('xdotool', ['getactivewindow']));
    } catch {
      return null;
    }
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
