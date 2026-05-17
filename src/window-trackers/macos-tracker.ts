/*
  subminer - Yomitan integration for mpv
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
import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';
import { BaseWindowTracker } from './base-tracker';
import { createLogger } from '../logger';
import type { WindowGeometry } from '../types';

const log = createLogger('tracker').child('macos');
const MACOS_FAST_POLL_INTERVAL_MS = 250;
const MACOS_STABLE_FOCUSED_POLL_INTERVAL_MS = 1_000;

type MacOSTrackerRunnerResult = {
  stdout: string;
  stderr: string;
};

type MacOSTrackerDeps = {
  resolveHelper?: () => { helperPath: string; helperType: 'binary' | 'swift' } | null;
  runHelper?: (
    helperPath: string,
    helperType: 'binary' | 'swift',
    targetMpvSocketPath: string | null,
  ) => Promise<MacOSTrackerRunnerResult>;
  maxConsecutiveMisses?: number;
  trackingLossGraceMs?: number;
  minimizedTrackingLossGraceMs?: number;
  now?: () => number;
  fastPollIntervalMs?: number;
  stablePollIntervalMs?: number;
  setPollTimeout?: typeof setTimeout;
  clearPollTimeout?: typeof clearTimeout;
};

export type MacOSHelperWindowState =
  | {
      geometry: WindowGeometry;
      focused: boolean;
      minimized?: false;
      active?: false;
      inactive?: false;
    }
  | {
      geometry: null;
      focused: true;
      active: true;
      minimized?: false;
      inactive?: false;
    }
  | {
      geometry: null;
      focused: false;
      inactive: true;
      active?: false;
      minimized?: false;
    }
  | {
      geometry: null;
      focused: false;
      minimized: true;
      active?: false;
      inactive?: false;
    };

function runHelperWithExecFile(
  helperPath: string,
  helperType: 'binary' | 'swift',
  targetMpvSocketPath: string | null,
): Promise<MacOSTrackerRunnerResult> {
  return new Promise((resolve, reject) => {
    const command = helperType === 'binary' ? helperPath : 'swift';
    const args = helperType === 'binary' ? [] : [helperPath];
    if (targetMpvSocketPath) {
      args.push(targetMpvSocketPath);
    }

    execFile(
      command,
      args,
      {
        encoding: 'utf-8',
        timeout: 1000,
        maxBuffer: 1024 * 1024,
      },
      (error, stdout, stderr) => {
        if (error) {
          reject(Object.assign(error, { stderr }));
          return;
        }
        resolve({
          stdout: stdout || '',
          stderr: stderr || '',
        });
      },
    );
  });
}

export function isCompiledMacOSHelperCurrent(
  binaryPath: string,
  sourcePath: string,
  helperFs: Pick<typeof fs, 'existsSync' | 'statSync'> = fs,
): boolean {
  if (!helperFs.existsSync(binaryPath)) {
    return false;
  }
  if (!helperFs.existsSync(sourcePath)) {
    return true;
  }

  try {
    return helperFs.statSync(binaryPath).mtimeMs >= helperFs.statSync(sourcePath).mtimeMs;
  } catch {
    return false;
  }
}

export function parseMacOSHelperOutput(result: string): MacOSHelperWindowState | null {
  const trimmed = result.trim();
  if (trimmed === 'minimized') {
    return {
      geometry: null,
      focused: false,
      minimized: true,
    };
  }
  if (trimmed === 'active') {
    return {
      geometry: null,
      focused: true,
      active: true,
    };
  }
  if (trimmed === 'inactive') {
    return {
      geometry: null,
      focused: false,
      inactive: true,
    };
  }
  if (!trimmed || trimmed === 'not-found') {
    return null;
  }

  const parts = trimmed.split(',');
  if (parts.length !== 4 && parts.length !== 5) {
    return null;
  }

  const x = parseInt(parts[0]!, 10);
  const y = parseInt(parts[1]!, 10);
  const width = parseInt(parts[2]!, 10);
  const height = parseInt(parts[3]!, 10);
  if (
    !Number.isFinite(x) ||
    !Number.isFinite(y) ||
    !Number.isFinite(width) ||
    !Number.isFinite(height) ||
    width <= 0 ||
    height <= 0
  ) {
    return null;
  }

  const focusedRaw = parts[4]?.trim().toLowerCase();
  const focused = focusedRaw === undefined ? true : focusedRaw === '1' || focusedRaw === 'true';

  return {
    geometry: {
      x,
      y,
      width,
      height,
    },
    focused,
  };
}

export class MacOSWindowTracker extends BaseWindowTracker {
  private pollTimeout: ReturnType<typeof setTimeout> | null = null;
  private pollInFlight = false;
  private started = false;
  private helperPath: string | null = null;
  private helperType: 'binary' | 'swift' | null = null;
  private lastExecErrorFingerprint: string | null = null;
  private lastExecErrorLoggedAtMs = 0;
  private readonly targetMpvSocketPath: string | null;
  private readonly runHelper: (
    helperPath: string,
    helperType: 'binary' | 'swift',
    targetMpvSocketPath: string | null,
  ) => Promise<MacOSTrackerRunnerResult>;
  private readonly maxConsecutiveMisses: number;
  private readonly trackingLossGraceMs: number;
  private readonly minimizedTrackingLossGraceMs: number;
  private readonly now: () => number;
  private readonly fastPollIntervalMs: number;
  private readonly stablePollIntervalMs: number;
  private readonly setPollTimeout: typeof setTimeout;
  private readonly clearPollTimeout: typeof clearTimeout;
  private consecutiveMisses = 0;
  private trackingLossStartedAtMs: number | null = null;
  private targetWindowMinimized = false;

  constructor(targetMpvSocketPath?: string, deps: MacOSTrackerDeps = {}) {
    super();
    this.targetMpvSocketPath = targetMpvSocketPath?.trim() || null;
    this.runHelper = deps.runHelper ?? runHelperWithExecFile;
    this.maxConsecutiveMisses = Math.max(1, Math.floor(deps.maxConsecutiveMisses ?? 2));
    this.trackingLossGraceMs = Math.max(0, Math.floor(deps.trackingLossGraceMs ?? 1_500));
    this.minimizedTrackingLossGraceMs = Math.max(
      0,
      Math.floor(deps.minimizedTrackingLossGraceMs ?? 500),
    );
    this.now = deps.now ?? (() => Date.now());
    this.fastPollIntervalMs = Math.max(
      50,
      Math.floor(deps.fastPollIntervalMs ?? MACOS_FAST_POLL_INTERVAL_MS),
    );
    this.stablePollIntervalMs = Math.max(
      this.fastPollIntervalMs,
      Math.floor(deps.stablePollIntervalMs ?? MACOS_STABLE_FOCUSED_POLL_INTERVAL_MS),
    );
    this.setPollTimeout = deps.setPollTimeout ?? setTimeout;
    this.clearPollTimeout = deps.clearPollTimeout ?? clearTimeout;
    const resolvedHelper = deps.resolveHelper?.() ?? null;
    if (resolvedHelper) {
      this.helperPath = resolvedHelper.helperPath;
      this.helperType = resolvedHelper.helperType;
    } else {
      this.detectHelper();
    }
  }

  private materializeAsarHelper(sourcePath: string, helperType: 'binary' | 'swift'): string | null {
    if (!sourcePath.includes('.asar')) {
      return sourcePath;
    }

    const fileName =
      helperType === 'binary' ? 'get-mpv-window-macos' : 'get-mpv-window-macos.swift';
    const targetDir = path.join(os.tmpdir(), 'subminer', 'helpers');
    const targetPath = path.join(targetDir, fileName);

    try {
      fs.mkdirSync(targetDir, { recursive: true });
      fs.copyFileSync(sourcePath, targetPath);
      fs.chmodSync(targetPath, 0o755);
      log.info(`Materialized macOS helper from asar: ${targetPath}`);
      return targetPath;
    } catch (error) {
      log.warn(`Failed to materialize helper from asar: ${sourcePath}`, error);
      return null;
    }
  }

  private tryUseHelper(candidatePath: string, helperType: 'binary' | 'swift'): boolean {
    if (!fs.existsSync(candidatePath)) {
      return false;
    }

    const resolvedPath = this.materializeAsarHelper(candidatePath, helperType);
    if (!resolvedPath) {
      return false;
    }

    this.helperPath = resolvedPath;
    this.helperType = helperType;
    log.info(`Using macOS helper (${helperType}): ${resolvedPath}`);
    return true;
  }

  private tryUseCompiledHelper(candidatePath: string, sourcePath: string): boolean {
    if (!isCompiledMacOSHelperCurrent(candidatePath, sourcePath)) {
      return false;
    }
    return this.tryUseHelper(candidatePath, 'binary');
  }

  private detectHelper(): void {
    const swiftPath = path.join(__dirname, '..', '..', 'scripts', 'get-mpv-window-macos.swift');

    // Prefer resources path (outside asar) in packaged apps.
    const resourcesPath = process.resourcesPath;
    if (resourcesPath) {
      const resourcesBinaryPath = path.join(resourcesPath, 'scripts', 'get-mpv-window-macos');
      if (this.tryUseHelper(resourcesBinaryPath, 'binary')) {
        return;
      }
    }

    // Built source runs from dist/window-trackers, so the compiled helper is a sibling of dist.
    const bundledBinaryPath = path.join(__dirname, '..', 'scripts', 'get-mpv-window-macos');
    if (this.tryUseCompiledHelper(bundledBinaryPath, swiftPath)) {
      return;
    }

    // Source-tree/manual helper build path.
    const sourceTreeBinaryPath = path.join(
      __dirname,
      '..',
      '..',
      'scripts',
      'get-mpv-window-macos',
    );
    if (this.tryUseCompiledHelper(sourceTreeBinaryPath, swiftPath)) {
      return;
    }

    // Fall back to Swift script for development or if binary filtering is not
    // supported in the current environment.
    if (this.tryUseHelper(swiftPath, 'swift')) {
      return;
    }

    log.warn('macOS window tracking helper not found');
  }

  private maybeLogExecError(err: Error, stderr: string): void {
    const now = Date.now();
    const fingerprint = `${err.message}|${stderr.trim()}`;
    const shouldLog =
      this.lastExecErrorFingerprint !== fingerprint || now - this.lastExecErrorLoggedAtMs >= 5000;
    if (!shouldLog) {
      return;
    }
    this.lastExecErrorFingerprint = fingerprint;
    this.lastExecErrorLoggedAtMs = now;
    log.warn('macOS helper execution failed', {
      helperPath: this.helperPath,
      helperType: this.helperType,
      error: err.message,
      stderr: stderr.trim(),
    });
  }

  start(): void {
    if (this.started) {
      return;
    }
    this.started = true;
    this.pollGeometry();
  }

  stop(): void {
    this.started = false;
    this.clearScheduledPoll();
  }

  override isTargetWindowMinimized(): boolean {
    return this.targetWindowMinimized;
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

  private shouldPreserveFocusedTargetOnMiss(): boolean {
    return this.isTracking() && this.isTargetWindowFocused() && this.getGeometry() !== null;
  }

  private registerTrackingMiss(graceMs = this.trackingLossGraceMs): void {
    if (this.shouldPreserveFocusedTargetOnMiss()) {
      this.resetTrackingLossState();
      return;
    }

    this.consecutiveMisses += 1;
    if (this.shouldDropTracking(graceMs)) {
      this.updateGeometry(null);
      this.resetTrackingLossState();
    }
  }

  private resolveNextPollIntervalMs(): number {
    if (
      this.isTracking() &&
      this.isTargetWindowFocused() &&
      !this.targetWindowMinimized &&
      this.getGeometry() !== null
    ) {
      return this.stablePollIntervalMs;
    }

    return this.fastPollIntervalMs;
  }

  private clearScheduledPoll(): void {
    if (!this.pollTimeout) {
      return;
    }

    this.clearPollTimeout(this.pollTimeout);
    this.pollTimeout = null;
  }

  private scheduleNextPoll(): void {
    if (!this.started || this.pollTimeout) {
      return;
    }

    this.pollTimeout = this.setPollTimeout(() => {
      this.pollTimeout = null;
      this.pollGeometry();
    }, this.resolveNextPollIntervalMs());
  }

  private pollGeometry(): void {
    if (this.pollInFlight || !this.helperPath || !this.helperType) {
      return;
    }

    this.pollInFlight = true;
    void this.runHelper(this.helperPath, this.helperType, this.targetMpvSocketPath)
      .then(({ stdout }) => {
        const parsed = parseMacOSHelperOutput(stdout || '');
        if (parsed) {
          if (parsed.minimized) {
            this.targetWindowMinimized = true;
            this.updateTargetWindowFocused(false);
            this.registerTrackingMiss(this.minimizedTrackingLossGraceMs);
            return;
          }
          if (parsed.active) {
            this.resetTrackingLossState();
            this.targetWindowMinimized = false;
            this.updateTargetWindowFocused(true);
            return;
          }
          if (parsed.inactive) {
            this.targetWindowMinimized = false;
            this.updateTargetWindowFocused(false);
            this.registerTrackingMiss();
            return;
          }
          this.resetTrackingLossState();
          this.targetWindowMinimized = false;
          this.updateGeometry(parsed.geometry, parsed.focused);
          this.updateTargetWindowFocused(parsed.focused);
          return;
        }

        this.targetWindowMinimized = false;
        this.registerTrackingMiss();
      })
      .catch((error: unknown) => {
        const err = error instanceof Error ? error : new Error(String(error));
        const stderr =
          typeof error === 'object' &&
          error !== null &&
          'stderr' in error &&
          typeof (error as { stderr?: unknown }).stderr === 'string'
            ? (error as { stderr: string }).stderr
            : '';
        this.maybeLogExecError(err, stderr);
        this.targetWindowMinimized = false;
        this.registerTrackingMiss();
      })
      .finally(() => {
        this.pollInFlight = false;
        this.scheduleNextPoll();
      });
  }
}
