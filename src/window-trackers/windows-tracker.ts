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

import { execFile, type ExecFileException } from 'child_process';
import { BaseWindowTracker } from './base-tracker';
import {
  parseWindowTrackerHelperFocusState,
  parseWindowTrackerHelperOutput,
  resolveWindowsTrackerHelper,
  type WindowsTrackerHelperLaunchSpec,
} from './windows-helper';
import { createLogger } from '../logger';

const log = createLogger('tracker').child('windows');

type WindowsTrackerRunnerResult = {
  stdout: string;
  stderr: string;
};

type WindowsTrackerDeps = {
  resolveHelper?: () => WindowsTrackerHelperLaunchSpec | null;
  runHelper?: (
    spec: WindowsTrackerHelperLaunchSpec,
    mode: 'geometry',
    targetMpvSocketPath: string | null,
  ) => Promise<WindowsTrackerRunnerResult>;
};

function runHelperWithExecFile(
  spec: WindowsTrackerHelperLaunchSpec,
  mode: 'geometry',
  targetMpvSocketPath: string | null,
): Promise<WindowsTrackerRunnerResult> {
  return new Promise((resolve, reject) => {
    const modeArgs = spec.kind === 'native' ? ['--mode', mode] : ['-Mode', mode];
    const args = targetMpvSocketPath
      ? [...spec.args, ...modeArgs, targetMpvSocketPath]
      : [...spec.args, ...modeArgs];
    execFile(
      spec.command,
      args,
      {
        encoding: 'utf-8',
        timeout: 1000,
        maxBuffer: 1024 * 1024,
        windowsHide: true,
      },
      (error: ExecFileException | null, stdout: string, stderr: string) => {
        if (error) {
          reject(Object.assign(error, { stderr }));
          return;
        }
        resolve({ stdout, stderr });
      },
    );
  });
}

export class WindowsWindowTracker extends BaseWindowTracker {
  private pollInterval: ReturnType<typeof setInterval> | null = null;
  private pollInFlight = false;
  private helperSpec: WindowsTrackerHelperLaunchSpec | null;
  private readonly targetMpvSocketPath: string | null;
  private readonly runHelper: (
    spec: WindowsTrackerHelperLaunchSpec,
    mode: 'geometry',
    targetMpvSocketPath: string | null,
  ) => Promise<WindowsTrackerRunnerResult>;
  private lastExecErrorFingerprint: string | null = null;
  private lastExecErrorLoggedAtMs = 0;

  constructor(targetMpvSocketPath?: string, deps: WindowsTrackerDeps = {}) {
    super();
    this.targetMpvSocketPath = targetMpvSocketPath?.trim() || null;
    this.helperSpec = deps.resolveHelper ? deps.resolveHelper() : resolveWindowsTrackerHelper();
    this.runHelper = deps.runHelper ?? runHelperWithExecFile;
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

  private maybeLogExecError(error: Error, stderr: string): void {
    const now = Date.now();
    const fingerprint = `${error.message}|${stderr.trim()}`;
    const shouldLog =
      this.lastExecErrorFingerprint !== fingerprint || now - this.lastExecErrorLoggedAtMs >= 5000;
    if (!shouldLog) {
      return;
    }

    this.lastExecErrorFingerprint = fingerprint;
    this.lastExecErrorLoggedAtMs = now;
    log.warn('Windows helper execution failed', {
      helperPath: this.helperSpec?.helperPath ?? null,
      helperKind: this.helperSpec?.kind ?? null,
      error: error.message,
      stderr: stderr.trim(),
    });
  }

  private async runHelperWithSocketFallback(): Promise<WindowsTrackerRunnerResult> {
    if (!this.helperSpec) {
      return { stdout: 'not-found', stderr: '' };
    }

    try {
      const primary = await this.runHelper(this.helperSpec, 'geometry', this.targetMpvSocketPath);
      const primaryGeometry = parseWindowTrackerHelperOutput(primary.stdout);
      if (primaryGeometry || !this.targetMpvSocketPath) {
        return primary;
      }
    } catch (error) {
      if (!this.targetMpvSocketPath) {
        throw error;
      }
    }

    return await this.runHelper(this.helperSpec, 'geometry', null);
  }

  private pollGeometry(): void {
    if (this.pollInFlight || !this.helperSpec) {
      return;
    }

    this.pollInFlight = true;
    void this.runHelperWithSocketFallback()
      .then(({ stdout, stderr }) => {
        const geometry = parseWindowTrackerHelperOutput(stdout);
        const focusState = parseWindowTrackerHelperFocusState(stderr);
        this.updateTargetWindowFocused(focusState ?? Boolean(geometry));
        this.updateGeometry(geometry);
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
        this.updateGeometry(null);
      })
      .finally(() => {
        this.pollInFlight = false;
      });
  }
}
