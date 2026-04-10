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

import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { execFile, type ExecFileException } from 'child_process';
import type { WindowGeometry } from '../types';
import { createLogger } from '../logger';

const log = createLogger('tracker').child('windows-helper');

export type WindowsTrackerHelperKind = 'powershell' | 'native';
export type WindowsTrackerHelperMode = 'auto' | 'powershell' | 'native';
export type WindowsTrackerHelperRunMode =
  | 'geometry'
  | 'foreground-process'
  | 'bind-overlay'
  | 'lower-overlay'
  | 'set-owner'
  | 'clear-owner';

export type WindowsTrackerHelperLaunchSpec = {
  kind: WindowsTrackerHelperKind;
  command: string;
  args: string[];
  helperPath: string;
};

type ResolveWindowsTrackerHelperOptions = {
  dirname?: string;
  resourcesPath?: string;
  helperModeEnv?: string | undefined;
  helperPathEnv?: string | undefined;
  existsSync?: (candidate: string) => boolean;
  mkdirSync?: (candidate: string, options: { recursive: true }) => void;
  copyFileSync?: (source: string, destination: string) => void;
};

const windowsPath = path.win32;

function normalizeHelperMode(value: string | undefined): WindowsTrackerHelperMode {
  const normalized = value?.trim().toLowerCase();
  if (normalized === 'powershell' || normalized === 'native') {
    return normalized;
  }
  return 'auto';
}

function inferHelperKindFromPath(helperPath: string): WindowsTrackerHelperKind | null {
  const normalized = helperPath.trim().toLowerCase();
  if (normalized.endsWith('.exe')) return 'native';
  if (normalized.endsWith('.ps1')) return 'powershell';
  return null;
}

function materializeAsarHelper(
  sourcePath: string,
  kind: WindowsTrackerHelperKind,
  deps: Required<Pick<ResolveWindowsTrackerHelperOptions, 'mkdirSync' | 'copyFileSync'>>,
): string | null {
  if (!sourcePath.includes('.asar')) {
    return sourcePath;
  }

  const fileName = kind === 'native' ? 'get-mpv-window-windows.exe' : 'get-mpv-window-windows.ps1';
  const targetDir = windowsPath.join(os.tmpdir(), 'subminer', 'helpers');
  const targetPath = windowsPath.join(targetDir, fileName);

  try {
    deps.mkdirSync(targetDir, { recursive: true });
    deps.copyFileSync(sourcePath, targetPath);
    log.info(`Materialized Windows helper from asar: ${targetPath}`);
    return targetPath;
  } catch (error) {
    log.warn(`Failed to materialize Windows helper from asar: ${sourcePath}`, error);
    return null;
  }
}

function createLaunchSpec(
  helperPath: string,
  kind: WindowsTrackerHelperKind,
): WindowsTrackerHelperLaunchSpec {
  if (kind === 'native') {
    return {
      kind,
      command: helperPath,
      args: [],
      helperPath,
    };
  }

  return {
    kind,
    command: 'powershell.exe',
    args: ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', helperPath],
    helperPath,
  };
}

function normalizeHelperPathOverride(
  helperPathEnv: string | undefined,
  mode: WindowsTrackerHelperMode,
): { path: string; kind: WindowsTrackerHelperKind } | null {
  const helperPath = helperPathEnv?.trim();
  if (!helperPath) {
    return null;
  }

  const inferredKind = inferHelperKindFromPath(helperPath);
  const kind = mode === 'auto' ? inferredKind : mode;
  if (!kind) {
    log.warn(
      `Ignoring SUBMINER_WINDOWS_TRACKER_HELPER_PATH with unsupported extension: ${helperPath}`,
    );
    return null;
  }

  return { path: helperPath, kind };
}

function getHelperCandidates(
  dirname: string,
  resourcesPath: string | undefined,
): Array<{
  path: string;
  kind: WindowsTrackerHelperKind;
}> {
  const scriptFileBase = 'get-mpv-window-windows';
  const candidates: Array<{ path: string; kind: WindowsTrackerHelperKind }> = [];

  if (resourcesPath) {
    candidates.push({
      path: windowsPath.join(resourcesPath, 'scripts', `${scriptFileBase}.exe`),
      kind: 'native',
    });
    candidates.push({
      path: windowsPath.join(resourcesPath, 'scripts', `${scriptFileBase}.ps1`),
      kind: 'powershell',
    });
  }

  candidates.push({
    path: windowsPath.join(dirname, '..', 'scripts', `${scriptFileBase}.exe`),
    kind: 'native',
  });
  candidates.push({
    path: windowsPath.join(dirname, '..', 'scripts', `${scriptFileBase}.ps1`),
    kind: 'powershell',
  });
  candidates.push({
    path: windowsPath.join(dirname, '..', '..', 'scripts', `${scriptFileBase}.exe`),
    kind: 'native',
  });
  candidates.push({
    path: windowsPath.join(dirname, '..', '..', 'scripts', `${scriptFileBase}.ps1`),
    kind: 'powershell',
  });

  return candidates;
}

export function parseWindowTrackerHelperOutput(output: string): WindowGeometry | null {
  const result = output.trim();
  if (!result || result === 'not-found') {
    return null;
  }

  const parts = result.split(',');
  if (parts.length !== 4) {
    return null;
  }

  const [xText, yText, widthText, heightText] = parts;
  const x = Number.parseInt(xText!, 10);
  const y = Number.parseInt(yText!, 10);
  const width = Number.parseInt(widthText!, 10);
  const height = Number.parseInt(heightText!, 10);
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

  return { x, y, width, height };
}

export function parseWindowTrackerHelperFocusState(output: string): boolean | null {
  const focusLine = output
    .split(/\r?\n/)
    .map((line) => line.trim())
    .find((line) => line.startsWith('focus='));

  if (!focusLine) {
    return null;
  }

  const value = focusLine.slice('focus='.length).trim().toLowerCase();
  if (value === 'focused') {
    return true;
  }
  if (value === 'not-focused') {
    return false;
  }

  return null;
}

export function parseWindowTrackerHelperState(output: string): 'visible' | 'minimized' | null {
  const stateLine = output
    .split(/\r?\n/)
    .map((line) => line.trim())
    .find((line) => line.startsWith('state='));

  if (!stateLine) {
    return null;
  }

  const value = stateLine.slice('state='.length).trim().toLowerCase();
  if (value === 'visible') {
    return 'visible';
  }
  if (value === 'minimized') {
    return 'minimized';
  }

  return null;
}

export function parseWindowTrackerHelperForegroundProcess(output: string): string | null {
  const processLine = output
    .split(/\r?\n/)
    .map((line) => line.trim())
    .find((line) => line.startsWith('process='));

  if (!processLine) {
    return null;
  }

  const value = processLine.slice('process='.length).trim();
  return value.length > 0 ? value : null;
}

type WindowsTrackerHelperRunnerResult = {
  stdout: string;
  stderr: string;
};

function runWindowsTrackerHelperWithExecFile(
  spec: WindowsTrackerHelperLaunchSpec,
  mode: WindowsTrackerHelperRunMode,
  extraArgs: string[] = [],
): Promise<WindowsTrackerHelperRunnerResult> {
  return new Promise((resolve, reject) => {
    const modeArgs = spec.kind === 'native' ? ['--mode', mode] : ['-Mode', mode];
    execFile(
      spec.command,
      [...spec.args, ...modeArgs, ...extraArgs],
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

export async function queryWindowsForegroundProcessName(deps: {
  resolveHelper?: () => WindowsTrackerHelperLaunchSpec | null;
  runHelper?: (
    spec: WindowsTrackerHelperLaunchSpec,
    mode: WindowsTrackerHelperRunMode,
    extraArgs?: string[],
  ) => Promise<WindowsTrackerHelperRunnerResult>;
} = {}): Promise<string | null> {
  const spec =
    deps.resolveHelper?.() ??
    resolveWindowsTrackerHelper({
      helperModeEnv: 'powershell',
    });
  if (!spec || spec.kind !== 'powershell') {
    return null;
  }

  const runHelper = deps.runHelper ?? runWindowsTrackerHelperWithExecFile;
  const { stdout } = await runHelper(spec, 'foreground-process');
  return parseWindowTrackerHelperForegroundProcess(stdout);
}

export async function syncWindowsOverlayToMpvZOrder(deps: {
  overlayWindowHandle: string;
  targetMpvSocketPath?: string | null;
  resolveHelper?: () => WindowsTrackerHelperLaunchSpec | null;
  runHelper?: (
    spec: WindowsTrackerHelperLaunchSpec,
    mode: WindowsTrackerHelperRunMode,
    extraArgs?: string[],
  ) => Promise<WindowsTrackerHelperRunnerResult>;
}): Promise<boolean> {
  const spec =
    deps.resolveHelper?.() ??
    resolveWindowsTrackerHelper({
      helperModeEnv: 'powershell',
    });
  if (!spec || spec.kind !== 'powershell') {
    return false;
  }

  const runHelper = deps.runHelper ?? runWindowsTrackerHelperWithExecFile;
  const extraArgs = [deps.targetMpvSocketPath ?? '', deps.overlayWindowHandle];
  const { stdout } = await runHelper(spec, 'bind-overlay', extraArgs);
  return stdout.trim() === 'ok';
}

export async function lowerWindowsOverlayInZOrder(deps: {
  overlayWindowHandle: string;
  resolveHelper?: () => WindowsTrackerHelperLaunchSpec | null;
  runHelper?: (
    spec: WindowsTrackerHelperLaunchSpec,
    mode: WindowsTrackerHelperRunMode,
    extraArgs?: string[],
  ) => Promise<WindowsTrackerHelperRunnerResult>;
}): Promise<boolean> {
  const spec =
    deps.resolveHelper?.() ??
    resolveWindowsTrackerHelper({
      helperModeEnv: 'powershell',
    });
  if (!spec || spec.kind !== 'powershell') {
    return false;
  }

  const runHelper = deps.runHelper ?? runWindowsTrackerHelperWithExecFile;
  const { stdout } = await runHelper(spec, 'lower-overlay', [deps.overlayWindowHandle]);
  return stdout.trim() === 'ok';
}

export function setWindowsOverlayOwnerNative(overlayHwnd: number, mpvHwnd: number): boolean {
  try {
    const win32 = require('./win32') as typeof import('./win32');
    win32.setOverlayOwner(overlayHwnd, mpvHwnd);
    return true;
  } catch {
    return false;
  }
}

export function ensureWindowsOverlayTransparencyNative(overlayHwnd: number): boolean {
  try {
    const win32 = require('./win32') as typeof import('./win32');
    win32.ensureOverlayTransparency(overlayHwnd);
    return true;
  } catch {
    return false;
  }
}

export function bindWindowsOverlayAboveMpvNative(overlayHwnd: number): boolean {
  try {
    const win32 = require('./win32') as typeof import('./win32');
    const poll = win32.findMpvWindows();
    const focused = poll.matches.find((m) => m.isForeground);
    const best = focused ?? poll.matches.sort((a, b) => b.area - a.area)[0];
    if (!best) return false;
    win32.bindOverlayAboveMpv(overlayHwnd, best.hwnd);
    win32.ensureOverlayTransparency(overlayHwnd);
    return true;
  } catch {
    return false;
  }
}

export function clearWindowsOverlayOwnerNative(overlayHwnd: number): boolean {
  try {
    const win32 = require('./win32') as typeof import('./win32');
    win32.clearOverlayOwner(overlayHwnd);
    return true;
  } catch {
    return false;
  }
}

export function getWindowsForegroundProcessNameNative(): string | null {
  try {
    const win32 = require('./win32') as typeof import('./win32');
    return win32.getForegroundProcessName();
  } catch {
    return null;
  }
}

export function resolveWindowsTrackerHelper(
  options: ResolveWindowsTrackerHelperOptions = {},
): WindowsTrackerHelperLaunchSpec | null {
  const existsSync = options.existsSync ?? fs.existsSync;
  const mkdirSync = options.mkdirSync ?? fs.mkdirSync;
  const copyFileSync = options.copyFileSync ?? fs.copyFileSync;
  const dirname = options.dirname ?? __dirname;
  const resourcesPath = options.resourcesPath ?? process.resourcesPath;
  const mode = normalizeHelperMode(
    options.helperModeEnv ?? process.env.SUBMINER_WINDOWS_TRACKER_HELPER,
  );
  const override = normalizeHelperPathOverride(
    options.helperPathEnv ?? process.env.SUBMINER_WINDOWS_TRACKER_HELPER_PATH,
    mode,
  );

  if (override) {
    if (!existsSync(override.path)) {
      log.warn(`Configured Windows tracker helper path does not exist: ${override.path}`);
      return null;
    }
    const helperPath = materializeAsarHelper(override.path, override.kind, {
      mkdirSync,
      copyFileSync,
    });
    return helperPath ? createLaunchSpec(helperPath, override.kind) : null;
  }

  const candidates = getHelperCandidates(dirname, resourcesPath);
  const orderedCandidates =
    mode === 'powershell'
      ? candidates.filter((candidate) => candidate.kind === 'powershell')
      : mode === 'native'
        ? candidates.filter((candidate) => candidate.kind === 'native')
        : candidates;

  for (const candidate of orderedCandidates) {
    if (!existsSync(candidate.path)) {
      continue;
    }

    const helperPath = materializeAsarHelper(candidate.path, candidate.kind, {
      mkdirSync,
      copyFileSync,
    });
    if (!helperPath) {
      continue;
    }

    log.info(`Using Windows helper (${candidate.kind}): ${helperPath}`);
    return createLaunchSpec(helperPath, candidate.kind);
  }

  if (mode === 'native') {
    log.warn('Windows native tracker helper requested but no helper was found.');
  } else if (mode === 'powershell') {
    log.warn('Windows PowerShell tracker helper requested but no helper was found.');
  } else {
    log.warn('Windows tracker helper not found.');
  }

  return null;
}
