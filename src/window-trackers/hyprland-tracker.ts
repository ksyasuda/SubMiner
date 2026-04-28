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
import { execSync } from 'child_process';
import { BaseWindowTracker } from './base-tracker';
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

export function selectHyprlandMpvWindow(
  clients: HyprlandClient[],
  options: SelectHyprlandMpvWindowOptions,
): HyprlandClient | null {
  const visibleMpvWindows = clients.filter(
    (client) =>
      (isMpvClassName(client.class) || isMpvClassName(client.initialClass)) &&
      client.mapped !== false &&
      client.hidden !== true,
  );

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

  const parsed = JSON.parse(jsonPayload) as unknown;
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

  const parsed = JSON.parse(jsonPayload) as unknown;
  if (!Array.isArray(parsed)) {
    return null;
  }

  return parsed as HyprlandMonitor[];
}

function isHyprlandFullscreenClient(client: HyprlandClient): boolean {
  return (client.fullscreen ?? 0) > 0;
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
    name === 'windowtitle' ||
    name === 'windowtitlev2' ||
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
  private activeWindowAddress: string | null = null;

  constructor(targetMpvSocketPath?: string) {
    super();
    this.targetMpvSocketPath = targetMpvSocketPath?.trim() || null;
  }

  start(): void {
    this.pollInterval = setInterval(() => this.pollGeometry(), 250);
    this.pollGeometry();
    this.connectEventSocket();
  }

  stop(): void {
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
      this.activeWindowAddress = data || null;
      this.pollGeometry();
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
        this.pollGeometry();
      }, delayMs);
      return pollTimeout;
    });
    for (const pollTimeout of this.pollTimeouts) {
      pollTimeout.unref?.();
    }
  }

  private pollGeometry(): void {
    try {
      const output = execSync('hyprctl -j clients', { encoding: 'utf-8' });
      const clients = parseHyprctlClients(output);
      if (!clients) {
        this.updateGeometry(null);
        return;
      }
      const mpvWindow = this.findTargetWindow(clients);

      if (mpvWindow) {
        this.updateGeometry(
          resolveHyprlandWindowGeometry(mpvWindow, this.getHyprlandMonitors(mpvWindow)),
        );
      } else {
        this.updateGeometry(null);
      }
    } catch (err) {
      // hyprctl not available or failed - silent fail
    }
  }

  private findTargetWindow(clients: HyprlandClient[]): HyprlandClient | null {
    return selectHyprlandMpvWindow(clients, {
      targetMpvSocketPath: this.targetMpvSocketPath,
      activeWindowAddress: this.activeWindowAddress,
      getWindowCommandLine: (pid) => this.getWindowCommandLine(pid),
    });
  }

  private getHyprlandMonitors(client: HyprlandClient): HyprlandMonitor[] | null {
    if (!isHyprlandFullscreenClient(client)) {
      return null;
    }

    try {
      const output = execSync('hyprctl -j monitors', { encoding: 'utf-8' });
      return parseHyprctlMonitors(output);
    } catch {
      return null;
    }
  }

  private getWindowCommandLine(pid: number): string | null {
    const commandLine = execSync(`ps -p ${pid} -o args=`, {
      encoding: 'utf-8',
    }).trim();
    return commandLine || null;
  }
}
