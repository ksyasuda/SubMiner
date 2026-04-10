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
import { queryWindowsTrackerMpvWindows } from './windows-helper';
import { createLogger } from '../logger';

const log = createLogger('tracker').child('windows');

type WindowsTrackerDeps = {
  pollMpvWindows?: () => MpvPollResult;
  maxConsecutiveMisses?: number;
  trackingLossGraceMs?: number;
  minimizedTrackingLossGraceMs?: number;
  now?: () => number;
};

function shouldUsePowershellTrackerFallback(): boolean {
  const helperMode = process.env.SUBMINER_WINDOWS_TRACKER_HELPER?.trim().toLowerCase();
  if (helperMode === 'powershell') {
    return true;
  }

  const helperPath = process.env.SUBMINER_WINDOWS_TRACKER_HELPER_PATH?.trim().toLowerCase();
  return helperPath?.endsWith('.ps1') ?? false;
}

function defaultPollMpvWindows(targetMpvSocketPath?: string | null): MpvPollResult {
  if (targetMpvSocketPath && shouldUsePowershellTrackerFallback()) {
    const helperResult = queryWindowsTrackerMpvWindows({
      targetMpvSocketPath,
    });
    if (helperResult) {
      return helperResult;
    }
  }

  const win32 = require('./win32') as typeof import('./win32');
  return win32.findMpvWindows();
}

export class WindowsWindowTracker extends BaseWindowTracker {
  private pollInterval: ReturnType<typeof setInterval> | null = null;
  private pollInFlight = false;
  private readonly pollMpvWindows: () => MpvPollResult;
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

  constructor(_targetMpvSocketPath?: string, deps: WindowsTrackerDeps = {}) {
    super();
    this.targetMpvSocketPath = _targetMpvSocketPath?.trim() || null;
    this.pollMpvWindows = deps.pollMpvWindows ?? (() => defaultPollMpvWindows(this.targetMpvSocketPath));
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
  ): { geometry: WindowGeometry; focused: boolean } | null {
    if (result.matches.length === 0) return null;

    const focusedMatch = result.matches.find((m) => m.isForeground);
    const best =
      focusedMatch ??
      result.matches.sort((a, b) => b.area - a.area || b.bounds.width - a.bounds.width)[0]!;

    return {
      geometry: best.bounds,
      focused: best.isForeground,
    };
  }

  private pollGeometry(): void {
    if (this.pollInFlight) return;
    this.pollInFlight = true;

    try {
      const result = this.pollMpvWindows();
      const best = this.selectBestMatch(result);

      if (best) {
        this.resetTrackingLossState();
        this.targetWindowMinimized = false;
        this.updateTargetWindowFocused(best.focused);
        this.updateGeometry(best.geometry);
        return;
      }

      if (result.windowState === 'minimized') {
        this.targetWindowMinimized = true;
        this.updateTargetWindowFocused(false);
        this.registerTrackingMiss(this.minimizedTrackingLossGraceMs);
        return;
      }

      this.targetWindowMinimized = false;
      this.updateTargetWindowFocused(false);
      this.registerTrackingMiss();
    } catch (error: unknown) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.maybeLogPollError(err);
      this.targetWindowMinimized = false;
      this.updateTargetWindowFocused(false);
      this.registerTrackingMiss();
    } finally {
      this.pollInFlight = false;
    }
  }
}
