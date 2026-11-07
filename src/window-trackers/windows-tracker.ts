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
import type { WindowGeometry } from '../types';
import type { MpvPollResult } from './win32';
import { queryMpvProcessId } from './mpv-process-id';
import { createLogger } from '../logger';

const log = createLogger('tracker').child('windows');

const MPV_PID_LOOKUP_INTERVAL_MS = 1_000;
const MPV_PID_REVALIDATE_INTERVAL_MS = 10_000;
const MPV_PID_MAX_AGE_MS = 30_000;
const NO_MPV_WINDOW: MpvPollResult = { matches: [], focusState: false, windowState: 'not-found' };

type WindowsTrackerDeps = {
  pollMpvWindows?: (targetPid: number | null) => MpvPollResult;
  queryMpvProcessId?: (socketPath: string) => Promise<number>;
  maxConsecutiveMisses?: number;
  trackingLossGraceMs?: number;
  minimizedTrackingLossGraceMs?: number;
  now?: () => number;
};

function defaultPollMpvWindows(targetPid: number | null): MpvPollResult {
  const win32 = require('./win32') as typeof import('./win32');
  return win32.findMpvWindows(targetPid);
}

export class WindowsWindowTracker extends BaseWindowTracker {
  private pollInterval: ReturnType<typeof setInterval> | null = null;
  private pollInFlight = false;
  private readonly pollMpvWindows: (targetPid: number | null) => MpvPollResult;
  private readonly queryMpvProcessId: (socketPath: string) => Promise<number>;
  private readonly maxConsecutiveMisses: number;
  private readonly trackingLossGraceMs: number;
  private readonly minimizedTrackingLossGraceMs: number;
  private readonly now: () => number;
  private lastPollErrorFingerprint: string | null = null;
  private lastPollErrorLoggedAtMs = 0;
  private consecutiveMisses = 0;
  private trackingLossStartedAtMs: number | null = null;
  private targetWindowMinimized = false;
  private readonly targetMpvSocketPath: string | null;
  private currentTargetWindowHwnd: number | null = null;
  private targetMpvPid: number | null = null;
  private pidLookupInFlight = false;
  private lastPidLookupAtMs: number | null = null;
  private lastPidConfirmedAtMs: number | null = null;
  private pidLookupFailureLogged = false;

  constructor(_targetMpvSocketPath?: string, deps: WindowsTrackerDeps = {}) {
    super();
    this.targetMpvSocketPath = _targetMpvSocketPath?.trim() || null;
    this.pollMpvWindows = deps.pollMpvWindows ?? defaultPollMpvWindows;
    this.queryMpvProcessId = deps.queryMpvProcessId ?? queryMpvProcessId;
    this.maxConsecutiveMisses = Math.max(1, Math.floor(deps.maxConsecutiveMisses ?? 2));
    this.trackingLossGraceMs = Math.max(0, Math.floor(deps.trackingLossGraceMs ?? 1_500));
    this.minimizedTrackingLossGraceMs = Math.max(
      0,
      Math.floor(deps.minimizedTrackingLossGraceMs ?? 500),
    );
    this.now = deps.now ?? (() => Date.now());
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

  override isTargetWindowMinimized(): boolean {
    return this.targetWindowMinimized;
  }

  getTargetWindowHandle(): number | null {
    return this.currentTargetWindowHwnd;
  }

  getDiagnosticState() {
    return {
      socketPath: this.targetMpvSocketPath,
      targetPid: this.targetMpvPid,
      targetHwnd: this.currentTargetWindowHwnd,
      pidLookupPending: this.pidLookupInFlight,
      pidAgeMs: this.lastPidConfirmedAtMs === null ? null : this.now() - this.lastPidConfirmedAtMs,
    };
  }

  private maybeLogPollError(error: Error): void {
    const now = Date.now();
    const fingerprint = error.message;
    const shouldLog =
      this.lastPollErrorFingerprint !== fingerprint || now - this.lastPollErrorLoggedAtMs >= 5000;
    if (!shouldLog) return;

    this.lastPollErrorFingerprint = fingerprint;
    this.lastPollErrorLoggedAtMs = now;
    log.warn('Windows native poll failed', { error: error.message });
  }

  private resetTrackingLossState(): void {
    this.consecutiveMisses = 0;
    this.trackingLossStartedAtMs = null;
  }

  private shouldDropTracking(graceMs = this.trackingLossGraceMs): boolean {
    if (!this.isTracking()) {
      return true;
    }
    if (graceMs === 0) {
      return this.consecutiveMisses >= this.maxConsecutiveMisses;
    }
    if (this.trackingLossStartedAtMs === null) {
      this.trackingLossStartedAtMs = this.now();
      return false;
    }
    return this.now() - this.trackingLossStartedAtMs > graceMs;
  }

  private registerTrackingMiss(graceMs = this.trackingLossGraceMs): void {
    this.consecutiveMisses += 1;
    if (this.shouldDropTracking(graceMs)) {
      this.updateGeometry(null);
      this.resetTrackingLossState();
    }
  }

  private selectBestMatch(
    result: MpvPollResult,
  ): { geometry: WindowGeometry; focused: boolean; hwnd: number } | null {
    if (result.matches.length === 0) return null;

    const focusedMatch = result.matches.find((m) => m.isForeground);
    const best =
      focusedMatch ??
      [...result.matches].sort((a, b) => b.area - a.area || b.bounds.width - a.bounds.width)[0]!;

    return {
      geometry: best.bounds,
      focused: best.isForeground,
      hwnd: best.hwnd,
    };
  }

  // Asks the mpv on the target socket which process it is. Runs in the background so a slow
  // answer never stalls the poll, which shares the main thread with the mouse hook.
  private refreshTargetMpvPid(socketPath: string, intervalMs = MPV_PID_LOOKUP_INTERVAL_MS): void {
    const nowMs = this.now();
    if (
      this.pidLookupInFlight ||
      (this.lastPidLookupAtMs !== null && nowMs - this.lastPidLookupAtMs < intervalMs)
    ) {
      return;
    }

    this.pidLookupInFlight = true;
    this.lastPidLookupAtMs = nowMs;
    this.queryMpvProcessId(socketPath)
      .then(
        (pid) => {
          this.targetMpvPid = pid;
          this.lastPidConfirmedAtMs = this.now();
          if (this.pidLookupFailureLogged) {
            log.warn('Windows mpv PID lookup recovered', this.getDiagnosticState());
          }
          this.pidLookupFailureLogged = false;
        },
        (error: unknown) => {
          // The overlay stays hidden until mpv is identified, so leave a trace of why it isn't.
          if (this.pidLookupFailureLogged) return;
          this.pidLookupFailureLogged = true;
          log.warn('Windows mpv PID lookup failed; overlay attachment may be unavailable', {
            ...this.getDiagnosticState(),
            error: error instanceof Error ? error.message : String(error),
          });
        },
      )
      .finally(() => {
        this.pidLookupInFlight = false;
      });
  }

  private pollTargetWindows(): MpvPollResult {
    const socketPath = this.targetMpvSocketPath;
    if (!socketPath) return this.pollMpvWindows(null);

    // Windows can reuse a PID for an unrelated visible window. A cached match must
    // remain tied to the process answering on our socket, even while it is visible.
    if (
      this.lastPidConfirmedAtMs !== null &&
      this.now() - this.lastPidConfirmedAtMs >= MPV_PID_MAX_AGE_MS
    ) {
      this.targetMpvPid = null;
    }

    if (this.targetMpvPid === null) {
      this.refreshTargetMpvPid(socketPath);
      return NO_MPV_WINDOW;
    }

    const result = this.pollMpvWindows(this.targetMpvPid);
    // A new mpv may have taken over the socket after the known process exited.
    this.refreshTargetMpvPid(
      socketPath,
      result.windowState === 'not-found'
        ? MPV_PID_LOOKUP_INTERVAL_MS
        : MPV_PID_REVALIDATE_INTERVAL_MS,
    );
    return result;
  }

  private pollGeometry(): void {
    if (this.pollInFlight) return;
    this.pollInFlight = true;

    try {
      const result = this.pollTargetWindows();
      const best = this.selectBestMatch(result);

      if (best) {
        this.resetTrackingLossState();
        this.targetWindowMinimized = false;
        this.currentTargetWindowHwnd = best.hwnd;
        this.updateGeometry(best.geometry, best.focused);
        this.updateTargetWindowFocused(best.focused);
        return;
      }

      if (result.windowState === 'minimized') {
        this.targetWindowMinimized = true;
        this.currentTargetWindowHwnd = null;
        this.updateTargetWindowFocused(false);
        this.registerTrackingMiss(this.minimizedTrackingLossGraceMs);
        return;
      }

      this.targetWindowMinimized = false;
      this.currentTargetWindowHwnd = null;
      this.updateTargetWindowFocused(false);
      this.registerTrackingMiss();
    } catch (error: unknown) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.maybeLogPollError(err);
      this.targetWindowMinimized = false;
      this.currentTargetWindowHwnd = null;
      this.updateTargetWindowFocused(false);
      this.registerTrackingMiss();
    } finally {
      this.pollInFlight = false;
    }
  }
}
