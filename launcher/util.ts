import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { spawn } from 'node:child_process';
import type { LogLevel, CommandExecOptions, CommandExecResult } from './types.js';
import { log } from './log.js';

export function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export function isExecutable(filePath: string): boolean {
  try {
    fs.accessSync(filePath, fs.constants.X_OK);
    return true;
  } catch {
    return false;
  }
}

function isRunnableFile(filePath: string): boolean {
  try {
    if (!fs.statSync(filePath).isFile()) return false;
    return process.platform === 'win32' ? true : isExecutable(filePath);
  } catch {
    return false;
  }
}

function isPathLikeCommand(command: string): boolean {
  return (
    command.includes('/') ||
    command.includes('\\') ||
    /^[A-Za-z]:[\\/]/.test(command) ||
    command.startsWith('.')
  );
}

function getWindowsPathExts(): string[] {
  const raw = process.env.PATHEXT ?? '.COM;.EXE;.BAT;.CMD';
  return raw
    .split(';')
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0);
}

export function getPathEnv(): string {
  const pathKey = Object.keys(process.env).find((key) => key.toLowerCase() === 'path');
  return pathKey ? (process.env[pathKey] ?? '') : '';
}

function resolveExecutablePath(command: string): string | null {
  const tryCandidate = (candidate: string): string | null =>
    isRunnableFile(candidate) ? candidate : null;

  const resolveWindowsCandidate = (candidate: string): string | null => {
    const direct = tryCandidate(candidate);
    if (direct) return direct;
    if (path.extname(candidate)) return null;
    for (const ext of getWindowsPathExts()) {
      const withExt = tryCandidate(`${candidate}${ext}`);
      if (withExt) return withExt;
    }
    return null;
  };

  if (isPathLikeCommand(command)) {
    const resolved = path.resolve(resolvePathMaybe(command));
    return process.platform === 'win32' ? resolveWindowsCandidate(resolved) : tryCandidate(resolved);
  }

  const pathEnv = getPathEnv();
  for (const dir of pathEnv.split(path.delimiter)) {
    if (!dir) continue;
    const candidate = path.join(dir, command);
    const resolved =
      process.platform === 'win32' ? resolveWindowsCandidate(candidate) : tryCandidate(candidate);
    if (resolved) return resolved;
  }
  return null;
}

function normalizeWindowsBashArg(value: string): string {
  const normalized = value.replace(/\\/g, '/');
  const driveMatch = normalized.match(/^([A-Za-z]):\/(.*)$/);
  if (!driveMatch) {
    return normalized;
  }

  const [, driveLetter, remainder] = driveMatch;
  return `/mnt/${driveLetter!.toLowerCase()}/${remainder}`;
}

function resolveGitBashExecutable(): string | null {
  const directCandidates = [
    'C:\\Program Files\\Git\\bin\\bash.exe',
    'C:\\Program Files\\Git\\usr\\bin\\bash.exe',
  ];
  for (const candidate of directCandidates) {
    if (isRunnableFile(candidate)) return candidate;
  }

  const gitExecutable = resolveExecutablePath('git');
  if (!gitExecutable) return null;
  const gitDir = path.dirname(gitExecutable);
  const inferredCandidates = [
    path.resolve(gitDir, '..', 'bin', 'bash.exe'),
    path.resolve(gitDir, '..', 'usr', 'bin', 'bash.exe'),
  ];
  for (const candidate of inferredCandidates) {
    if (isRunnableFile(candidate)) return candidate;
  }
  return null;
}

function resolveWindowsBashTarget(): {
  command: string;
  flavor: 'git' | 'wsl';
} {
  const gitBash = resolveGitBashExecutable();
  if (gitBash) {
    return { command: gitBash, flavor: 'git' };
  }
  return {
    command: resolveExecutablePath('bash') ?? 'bash',
    flavor: 'wsl',
  };
}

function normalizeWindowsShellArg(value: string, flavor: 'git' | 'wsl'): string {
  if (!isPathLikeCommand(value)) {
    return value;
  }
  return flavor === 'git' ? value.replace(/\\/g, '/') : normalizeWindowsBashArg(value);
}

function readShebang(filePath: string): string {
  try {
    const fd = fs.openSync(filePath, 'r');
    try {
      const buffer = Buffer.alloc(160);
      const bytesRead = fs.readSync(fd, buffer, 0, buffer.length, 0);
      return buffer.toString('utf8', 0, bytesRead).split(/\r?\n/, 1)[0] ?? '';
    } finally {
      fs.closeSync(fd);
    }
  } catch {
    return '';
  }
}

export function commandExists(command: string): boolean {
  return resolveExecutablePath(command) !== null;
}

export function resolvePathMaybe(input: string): string {
  if (input.startsWith('~')) {
    return path.join(os.homedir(), input.slice(1));
  }
  return input;
}

export function resolveBinaryPathCandidate(input: string): string {
  const trimmed = input.trim();
  if (!trimmed) return '';
  const unquoted = trimmed.replace(/^"(.*)"$/, '$1').replace(/^'(.*)'$/, '$1');
  return resolvePathMaybe(unquoted);
}

export function realpathMaybe(filePath: string): string {
  try {
    return fs.realpathSync(filePath);
  } catch {
    return path.resolve(filePath);
  }
}

export function isUrlTarget(target: string): boolean {
  return /^https?:\/\//.test(target) || /^ytsearch:/.test(target);
}

export function isYoutubeTarget(target: string): boolean {
  return /^ytsearch:/.test(target) || /^https?:\/\/(www\.)?(youtube\.com|youtu\.be)\//.test(target);
}

export function sanitizeToken(value: string): string {
  return String(value)
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '');
}

export function normalizeBasename(value: string, fallback: string): string {
  const safe = sanitizeToken(value.replace(/[\\/]+/g, '-'));
  if (safe) return safe;
  const fallbackSafe = sanitizeToken(fallback);
  if (fallbackSafe) return fallbackSafe;
  return `${Date.now()}`;
}

export function normalizeLangCode(value: string): string {
  return value
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9-]+/g, '');
}

export function uniqueNormalizedLangCodes(values: string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const value of values) {
    const normalized = normalizeLangCode(value);
    if (!normalized || seen.has(normalized)) continue;
    seen.add(normalized);
    out.push(normalized);
  }
  return out;
}

export function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function parseBoolLike(value: string): boolean | null {
  const normalized = value.trim().toLowerCase();
  if (normalized === 'yes' || normalized === 'true' || normalized === '1' || normalized === 'on') {
    return true;
  }
  if (normalized === 'no' || normalized === 'false' || normalized === '0' || normalized === 'off') {
    return false;
  }
  return null;
}

export function inferWhisperLanguage(langCodes: string[], fallback: string): string {
  for (const lang of uniqueNormalizedLangCodes(langCodes)) {
    if (lang === 'jpn') return 'ja';
    if (lang.length >= 2) return lang.slice(0, 2);
  }
  return fallback;
}

export function resolveCommandInvocation(
  executable: string,
  args: string[],
): { command: string; args: string[] } {
  if (process.platform !== 'win32') {
    return { command: executable, args };
  }

  const resolvedExecutable = resolveExecutablePath(executable) ?? executable;
  const extension = path.extname(resolvedExecutable).toLowerCase();
  if (extension === '.ps1') {
    return {
      command: 'powershell.exe',
      args: ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', resolvedExecutable, ...args],
    };
  }

  if (extension === '.sh') {
    const bashTarget = resolveWindowsBashTarget();
    return {
      command: bashTarget.command,
      args: [
        normalizeWindowsShellArg(resolvedExecutable, bashTarget.flavor),
        ...args.map((arg) => normalizeWindowsShellArg(arg, bashTarget.flavor)),
      ],
    };
  }

  if (!extension) {
    const shebang = readShebang(resolvedExecutable);
    if (/^#!.*\b(?:sh|bash)\b/i.test(shebang)) {
      const bashTarget = resolveWindowsBashTarget();
      return {
        command: bashTarget.command,
        args: [
          normalizeWindowsShellArg(resolvedExecutable, bashTarget.flavor),
          ...args.map((arg) => normalizeWindowsShellArg(arg, bashTarget.flavor)),
        ],
      };
    }
  }

  return { command: resolvedExecutable, args };
}

export function runExternalCommand(
  executable: string,
  args: string[],
  opts: CommandExecOptions = {},
  childTracker?: Set<ReturnType<typeof spawn>>,
): Promise<CommandExecResult> {
  const allowFailure = opts.allowFailure === true;
  const captureStdout = opts.captureStdout === true;
  const configuredLogLevel = opts.logLevel ?? 'info';
  const commandLabel = opts.commandLabel || executable;
  const streamOutput = opts.streamOutput === true;

  return new Promise((resolve, reject) => {
    const target = resolveCommandInvocation(executable, args);
    log(
      'debug',
      configuredLogLevel,
      `[${commandLabel}] spawn: ${target.command} ${target.args.join(' ')}`,
    );
    const child = spawn(target.command, target.args, {
      stdio: ['ignore', 'pipe', 'pipe'],
      env: { ...process.env, ...opts.env },
    });
    childTracker?.add(child);

    let stdout = '';
    let stderr = '';
    let stdoutBuffer = '';
    let stderrBuffer = '';
    const flushLines = (
      buffer: string,
      level: LogLevel,
      sink: (remaining: string) => void,
    ): void => {
      const lines = buffer.split(/\r?\n/);
      const remaining = lines.pop() ?? '';
      for (const line of lines) {
        const trimmed = line.trim();
        if (trimmed.length > 0) {
          log(level, configuredLogLevel, `[${commandLabel}] ${trimmed}`);
        }
      }
      sink(remaining);
    };

    child.stdout.on('data', (chunk: Buffer) => {
      const text = chunk.toString();
      if (captureStdout) stdout += text;
      if (streamOutput) {
        stdoutBuffer += text;
        flushLines(stdoutBuffer, 'debug', (remaining) => {
          stdoutBuffer = remaining;
        });
      }
    });

    child.stderr.on('data', (chunk: Buffer) => {
      const text = chunk.toString();
      stderr += text;
      if (streamOutput) {
        stderrBuffer += text;
        flushLines(stderrBuffer, 'debug', (remaining) => {
          stderrBuffer = remaining;
        });
      }
    });

    child.on('error', (error) => {
      childTracker?.delete(child);
      reject(new Error(`Failed to start "${executable}": ${error.message}`));
    });

    child.on('close', (code) => {
      childTracker?.delete(child);
      if (streamOutput) {
        const trailingOut = stdoutBuffer.trim();
        if (trailingOut.length > 0) {
          log('debug', configuredLogLevel, `[${commandLabel}] ${trailingOut}`);
        }
        const trailingErr = stderrBuffer.trim();
        if (trailingErr.length > 0) {
          log('debug', configuredLogLevel, `[${commandLabel}] ${trailingErr}`);
        }
      }
      log(
        code === 0 ? 'debug' : 'warn',
        configuredLogLevel,
        `[${commandLabel}] exit code ${code ?? 1}`,
      );
      if (code !== 0 && !allowFailure) {
        const commandString = `${target.command} ${target.args.join(' ')}`;
        reject(
          new Error(`Command failed (${commandString}): ${stderr.trim() || `exit code ${code}`}`),
        );
        return;
      }
      resolve({ code: code ?? 1, stdout, stderr });
    });
  });
}
