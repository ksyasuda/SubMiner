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

const log = createLogger('tracker').child('macos');

export class MacOSWindowTracker extends BaseWindowTracker {
  private pollInterval: ReturnType<typeof setInterval> | null = null;
  private pollInFlight = false;
  private helperPath: string | null = null;
  private helperType: 'binary' | 'swift' | null = null;
  private lastExecErrorFingerprint: string | null = null;
  private lastExecErrorLoggedAtMs = 0;
  private readonly targetMpvSocketPath: string | null;

  constructor(targetMpvSocketPath?: string) {
    super();
    this.targetMpvSocketPath = targetMpvSocketPath?.trim() || null;
    this.detectHelper();
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

  private detectHelper(): void {
    const shouldFilterBySocket = this.targetMpvSocketPath !== null;

    // Fall back to Swift helper first when filtering by socket path to avoid
    // stale prebuilt binaries that don't support the new socket filter argument.
    const swiftPath = path.join(__dirname, '..', '..', 'scripts', 'get-mpv-window-macos.swift');
    if (shouldFilterBySocket && this.tryUseHelper(swiftPath, 'swift')) {
      return;
    }

    // Prefer resources path (outside asar) in packaged apps.
    const resourcesPath = process.resourcesPath;
    if (resourcesPath) {
      const resourcesBinaryPath = path.join(resourcesPath, 'scripts', 'get-mpv-window-macos');
      if (this.tryUseHelper(resourcesBinaryPath, 'binary')) {
        return;
      }
    }

    // Dist binary path (development / unpacked installs).
    const distBinaryPath = path.join(__dirname, '..', '..', 'scripts', 'get-mpv-window-macos');
    if (this.tryUseHelper(distBinaryPath, 'binary')) {
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
    this.pollInterval = setInterval(() => this.pollGeometry(), 250);
    this.pollGeometry();
  }

  stop(): void {
    if (this.pollInterval) {
      clearInterval(this.pollInterval);
      this.pollInterval = null;
    }
  }

  private pollGeometry(): void {
    if (this.pollInFlight || !this.helperPath || !this.helperType) {
      return;
    }

    this.pollInFlight = true;

    // Use Core Graphics API via Swift helper for reliable window detection
    // This works with both bundled and unbundled mpv installations
    const command = this.helperType === 'binary' ? this.helperPath : 'swift';
    const args = this.helperType === 'binary' ? [] : [this.helperPath];
    if (this.targetMpvSocketPath) {
      args.push(this.targetMpvSocketPath);
    }

    execFile(
      command,
      args,
      {
        encoding: 'utf-8',
        timeout: 1000,
        maxBuffer: 1024 * 1024,
      },
      (err, stdout, stderr) => {
        if (err) {
          this.maybeLogExecError(err, stderr || '');
          this.updateGeometry(null);
          this.pollInFlight = false;
          return;
        }

        const result = (stdout || '').trim();
        if (result && result !== 'not-found') {
          const parts = result.split(',');
          if (parts.length === 4) {
            const x = parseInt(parts[0]!, 10);
            const y = parseInt(parts[1]!, 10);
            const width = parseInt(parts[2]!, 10);
            const height = parseInt(parts[3]!, 10);

            if (
              Number.isFinite(x) &&
              Number.isFinite(y) &&
              Number.isFinite(width) &&
              Number.isFinite(height) &&
              width > 0 &&
              height > 0
            ) {
              this.updateGeometry({
                x,
                y,
                width,
                height,
              });
              this.pollInFlight = false;
              return;
            }
          }
        }

        this.updateGeometry(null);
        this.pollInFlight = false;
      },
    );
  }
}
