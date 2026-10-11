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

import * as net from 'net';
import { BaseWindowTracker } from './base-tracker';
import {
  createCoalescedTask,
  execFileWithTimeout,
  ProcessCommandLineCache,
  type CommandRunner,
} from './linux-poll';
import { createLogger } from '../logger';
import type { WindowGeometry } from '../types';

const log = createLogger('tracker').child('hyprland');

export interface HyprlandClient {
  address?: string;
  class: string;
  initialClass?: string;
  at: [number, number];
  size: [number, number];
  monitor?: number;
  fullscreen?: number;
  fullscreenClient?: number;
  pid?: number;
  mapped?: boolean;
  hidden?: boolean;
}

export interface HyprlandMonitor {
  id: number;
  x: number;
  y: number;
  width: number;
  height: number;
}

interface SelectHyprlandMpvWindowOptions {
  targetMpvSocketPath: string | null;
  activeWindowAddress: string | null;
  getWindowCommandLine: (pid: number) => string | null;
}

function extractHyprctlJsonPayload(output: string): string | null {
  const trimmed = output.trim();
  if (!trimmed) {
    return null;
  }

  const arrayStart = trimmed.indexOf('[');
  const objectStart = trimmed.indexOf('{');
  const startCandidates = [arrayStart, objectStart].filter((index) => index >= 0);
  if (startCandidates.length === 0) {
    return null;
  }

  const startIndex = Math.min(...startCandidates);
  return trimmed.slice(startIndex);
}

function matchesTargetSocket(commandLine: string, targetMpvSocketPath: string): boolean {
  return (
    commandLine.includes(`--input-ipc-server=${targetMpvSocketPath}`) ||
    commandLine.includes(`--input-ipc-server ${targetMpvSocketPath}`)
  );
}

function preferActiveHyprlandWindow(
  clients: HyprlandClient[],
  activeWindowAddress: string | null,
): HyprlandClient | null {
  if (activeWindowAddress) {
    const activeClient = clients.find((client) => client.address === activeWindowAddress);
    if (activeClient) {
      return activeClient;
    }
  }

  return clients[0] ?? null;
}

function isMpvClassName(value: string | undefined): boolean {
  if (!value) {
    return false;
  }

  return value.trim().toLowerCase().includes('mpv');
}

function filterVisibleMpvWindows(clients: HyprlandClient[]): HyprlandClient[] {
  return clients.filter(
    (client) =>
      (isMpvClassName(client.class) || isMpvClassName(client.initialClass)) &&
      client.mapped !== false &&
      client.hidden !== true,
  );
}

export function selectHyprlandMpvWindow(
  clients: HyprlandClient[],
  options: SelectHyprlandMpvWindowOptions,
): HyprlandClient | null {
  const visibleMpvWindows = filterVisibleMpvWindows(clients);

  if (!options.targetMpvSocketPath) {
    return preferActiveHyprlandWindow(visibleMpvWindows, options.activeWindowAddress);
  }
  const targetMpvSocketPath = options.targetMpvSocketPath;

  const matchingWindows = visibleMpvWindows.filter((client) => {
    if (!client.pid) {
      return false;
    }

    const commandLine = options.getWindowCommandLine(client.pid);
    if (!commandLine) {
      return false;
    }

    return matchesTargetSocket(commandLine, targetMpvSocketPath);
  });

  return preferActiveHyprlandWindow(matchingWindows, options.activeWindowAddress);
}

export function parseHyprctlClients(output: string): HyprlandClient[] | null {
  const jsonPayload = extractHyprctlJsonPayload(output);
  if (!jsonPayload) {
    return null;
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(jsonPayload) as unknown;
  } catch {
    return null;
  }
  if (!Array.isArray(parsed)) {
    return null;
  }

  return parsed as HyprlandClient[];
}

export function parseHyprctlMonitors(output: string): HyprlandMonitor[] | null {
  const jsonPayload = extractHyprctlJsonPayload(output);
  if (!jsonPayload) {
    return null;
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(jsonPayload) as unknown;
  } catch {
    return null;
  }
  if (!Array.isArray(parsed)) {
    return null;
  }

  return parsed as HyprlandMonitor[];
}

function isHyprlandFullscreenClient(client: HyprlandClient): boolean {
  return (client.fullscreen ?? 0) > 0 || (client.fullscreenClient ?? 0) > 0;
}

export function resolveHyprlandWindowGeometry(
  client: HyprlandClient,
  monitors: HyprlandMonitor[] | null,
): WindowGeometry {
  if (isHyprlandFullscreenClient(client) && typeof client.monitor === 'number') {
    const monitor = monitors?.find((candidate) => candidate.id === client.monitor);
    if (monitor) {
      return {
        x: monitor.x,
        y: monitor.y,
        width: monitor.width,
        height: monitor.height,
      };
    }
  }

  return {
    x: client.at[0],
    y: client.at[1],
    width: client.size[0],
    height: client.size[1],
  };
}

export function isHyprlandGeometryEvent(name: string): boolean {
  return (
    name === 'movewindow' ||
    name === 'movewindowv2' ||
    name === 'resizewindow' ||
    name === 'resizewindowv2' ||
    name === 'openwindow' ||
    name === 'closewindow' ||
    name === 'fullscreen' ||
    name === 'fullscreenv2' ||
    name === 'changefloatingmode' ||
    name === 'workspace' ||
    name === 'workspacev2' ||
    name === 'focusedmon' ||
    name === 'monitoradded' ||
    name === 'monitoraddedv2' ||
    name === 'monitorremoved'
  );
}

export class HyprlandWindowTracker extends BaseWindowTracker {
  private pollInterval: ReturnType<typeof setInterval> | null = null;
  private pollTimeouts: Array<ReturnType<typeof setTimeout>> = [];
  private eventSocket: net.Socket | null = null;
  private readonly targetMpvSocketPath: string | null;
  private readonly runCommand: CommandRunner;
  private readonly commandLines: ProcessCommandLineCache;
  private activeWindowAddress: string | null = null;
  private stopped = false;
  // Interval ticks, focus events, and geometry bursts all funnel through here,
  // so a slow hyprctl never has more than one poll in flight.
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
    this.connectEventSocket();
  }

  stop(): void {
    this.stopped = true;
    if (this.pollInterval) {
      clearInterval(this.pollInterval);
      this.pollInterval = null;
    }
    for (const timeout of this.pollTimeouts) {
      clearTimeout(timeout);
    }
    this.pollTimeouts = [];
    if (this.eventSocket) {
      this.eventSocket.destroy();
      this.eventSocket = null;
    }
  }

  override refreshNow(): Promise<void> {
    return this.pollGeometry();
  }

  private connectEventSocket(): void {
    const hyprlandSig = process.env.HYPRLAND_INSTANCE_SIGNATURE;
    if (!hyprlandSig) {
      log.info('HYPRLAND_INSTANCE_SIGNATURE not set, skipping event socket');
      return;
    }

    const xdgRuntime = process.env.XDG_RUNTIME_DIR || '/tmp';
    const socketPath = `${xdgRuntime}/hypr/${hyprlandSig}/.socket2.sock`;
    this.eventSocket = new net.Socket();

    this.eventSocket.on('connect', () => {
      log.info('Connected to Hyprland event socket');
    });

    this.eventSocket.on('data', (data: Buffer) => {
      const events = data.toString().split('\n');
      for (const event of events) {
        this.handleSocketEvent(event);
      }
    });

    this.eventSocket.on('error', (err: Error) => {
      log.error('Hyprland event socket error:', err.message);
    });

    this.eventSocket.on('close', () => {
      log.info('Hyprland event socket closed');
    });

    this.eventSocket.connect(socketPath);
  }

  private handleSocketEvent(event: string): void {
    const trimmedEvent = event.trim();
    if (!trimmedEvent) {
      return;
    }

    const [name, rawData = ''] = trimmedEvent.split('>>', 2);
    if (!name) {
      return;
    }
    const data = rawData.trim();

    if (name === 'activewindowv2') {
      const wasFocused = this.isTargetWindowFocused();
      this.activeWindowAddress = data || null;
      void this.pollGeometry().then(() => {
        // Re-emit unchanged focus so the overlay still refreshes its z-order.
        if (!this.stopped && this.isTargetWindowFocused() === wasFocused) {
          this.onWindowFocusChange?.(this.isTargetWindowFocused());
        }
      });
      return;
    }

    if (name === 'closewindow' && data === this.activeWindowAddress) {
      this.activeWindowAddress = null;
    }

    if (isHyprlandGeometryEvent(name)) {
      this.scheduleGeometryPollBurst();
    }
  }

  private scheduleGeometryPollBurst(): void {
    for (const timeout of this.pollTimeouts) {
      clearTimeout(timeout);
    }
    this.pollTimeouts = [0, 50, 150, 300].map((delayMs) => {
      const pollTimeout = setTimeout(() => {
        this.pollTimeouts = this.pollTimeouts.filter((timeout) => timeout !== pollTimeout);
        void this.pollGeometry();
      }, delayMs);
      return pollTimeout;
    });
    for (const pollTimeout of this.pollTimeouts) {
      pollTimeout.unref?.();
    }
  }

  private async pollGeometryOnce(): Promise<void> {
    if (this.stopped) {
      return;
    }

    let mpvWindow: HyprlandClient | null;
    let monitors: HyprlandMonitor[] | null = null;
    try {
      const clients = parseHyprctlClients(await this.runCommand('hyprctl', ['-j', 'clients']));
      mpvWindow = clients ? await this.findTargetWindow(clients) : null;
      if (mpvWindow) {
        monitors = await this.getHyprlandMonitors(mpvWindow);
      }
    } catch {
      // hyprctl/ps not available, timed out, or failed - keep the last geometry
      return;
    }
    if (this.stopped) {
      return;
    }

    if (mpvWindow) {
      const focused = !this.activeWindowAddress || mpvWindow.address === this.activeWindowAddress;
      this.updateGeometry(resolveHyprlandWindowGeometry(mpvWindow, monitors), focused);
      this.updateTargetWindowFocused(focused);
    } else {
      this.updateGeometry(null);
    }
  }

  private async findTargetWindow(clients: HyprlandClient[]): Promise<HyprlandClient | null> {
    // Resolve command lines up front so window selection itself stays synchronous.
    const commandLines = this.targetMpvSocketPath
      ? await this.commandLines.lookup(
          filterVisibleMpvWindows(clients).flatMap((client) => (client.pid ? [client.pid] : [])),
        )
      : null;
    return selectHyprlandMpvWindow(clients, {
      targetMpvSocketPath: this.targetMpvSocketPath,
      activeWindowAddress: this.activeWindowAddress,
      getWindowCommandLine: (pid) => commandLines?.get(pid) ?? null,
    });
  }

  private async getHyprlandMonitors(client: HyprlandClient): Promise<HyprlandMonitor[] | null> {
    if (!isHyprlandFullscreenClient(client)) {
      return null;
    }

    try {
      return parseHyprctlMonitors(await this.runCommand('hyprctl', ['-j', 'monitors']));
    } catch {
      return null;
    }
  }
}
