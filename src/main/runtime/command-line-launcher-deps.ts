import { spawn } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

export interface RunCommandResult {
  exitCode: number | null;
  stdout: string;
  stderr: string;
}

export type RunCommand = (
  command: string,
  args: string[],
  options?: { timeoutMs?: number; env?: NodeJS.ProcessEnv },
) => Promise<RunCommandResult>;

export type FsDeps = {
  existsSync?: (candidate: string) => boolean;
  accessSync?: (candidate: string, mode?: number) => void;
  mkdirSync?: (candidate: string, options?: { recursive?: boolean }) => unknown;
  copyFileSync?: (from: string, to: string) => void;
  writeFileSync?: (candidate: string, content: string, encoding?: BufferEncoding) => void;
  readFileSync?: (candidate: string, encoding?: BufferEncoding) => string;
  chmodSync?: (candidate: string, mode: number) => void;
};

export type CommonOptions = FsDeps & {
  platform?: NodeJS.Platform;
  env?: Record<string, string | undefined>;
  homeDir?: string;
  cwd?: string;
  resourcesPath?: string;
  appExePath?: string;
  launcherResourcePath?: string;
  runCommand?: RunCommand;
};

export type WindowsPathOptions = {
  localAppData?: string;
  userProfile?: string;
  getUserPath?: () => string;
  setUserPath?: (nextPath: string) => void | Promise<void>;
  broadcastEnvironmentChange?: () => void | Promise<void>;
};

export function platformOf(options: CommonOptions): NodeJS.Platform {
  return options.platform ?? process.platform;
}

export function envOf(options: CommonOptions): Record<string, string | undefined> {
  return options.env ?? process.env;
}

export function pathModuleFor(platform: NodeJS.Platform): typeof path.posix | typeof path.win32 {
  return platform === 'win32' ? path.win32 : path.posix;
}

export function existsSyncOf(options: FsDeps): (candidate: string) => boolean {
  return options.existsSync ?? fs.existsSync;
}

export function accessSyncOf(options: FsDeps): (candidate: string, mode?: number) => void {
  return options.accessSync ?? fs.accessSync;
}

export function splitPath(value: string | undefined, platform: NodeJS.Platform): string[] {
  if (!value) return [];
  const delimiter = platform === 'win32' ? ';' : ':';
  return value
    .split(delimiter)
    .map((entry) => entry.trim())
    .filter(Boolean);
}

export function normalizePathForCompare(
  candidate: string,
  platform: NodeJS.Platform,
  platformPath = pathModuleFor(platform),
): string {
  const normalized = platformPath.normalize(candidate).replace(/[\\/]+$/, '');
  return platform === 'win32' ? normalized.toLowerCase() : normalized;
}

export function pathEntriesContain(
  entries: string[],
  dir: string,
  platform: NodeJS.Platform,
): boolean {
  const normalizedDir = normalizePathForCompare(dir, platform);
  return entries.some((entry) => normalizePathForCompare(entry, platform) === normalizedDir);
}

function isExecutableFile(candidate: string, options: CommonOptions): boolean {
  try {
    if (!existsSyncOf(options)(candidate)) return false;
    if (options.existsSync && !options.accessSync) return true;
    accessSyncOf(options)(candidate, fs.constants.X_OK);
    return true;
  } catch {
    return false;
  }
}

export function findCommand(command: string, options: CommonOptions): string | null {
  const platform = platformOf(options);
  const platformPath = pathModuleFor(platform);
  const entries = splitPath(envOf(options).PATH, platform);
  const hasExtension = platformPath.extname(command) !== '';
  const extensions =
    platform === 'win32'
      ? hasExtension
        ? ['']
        : (envOf(options).PATHEXT?.split(';').filter(Boolean) ?? [
            '.exe',
            '.cmd',
            '.bat',
            '.EXE',
            '.CMD',
            '.BAT',
          ])
      : [''];

  for (const entry of entries) {
    for (const extension of extensions) {
      const candidate = platformPath.join(entry, `${command}${extension}`);
      if (isExecutableFile(candidate, options)) return candidate;
    }
  }
  return null;
}

export function tail(value: string, max = 1200): string {
  const clean = value.trim();
  return clean.length > max ? clean.slice(clean.length - max) : clean;
}

export function failureMessage(result: RunCommandResult, fallback: string): string {
  const detail = tail(result.stderr || result.stdout);
  return detail ? `${fallback}: ${detail}` : fallback;
}

function needsWindowsShell(command: string): boolean {
  return process.platform === 'win32' && /\.(cmd|bat)$/i.test(command);
}

function quoteForWindowsShell(value: string): string {
  return /\s/.test(value) ? `"${value}"` : value;
}

function createDefaultRunCommand(): RunCommand {
  return (command, args, options = {}) =>
    new Promise((resolve) => {
      const useShell = needsWindowsShell(command);
      let child: ReturnType<typeof spawn>;
      try {
        child = useShell
          ? spawn(quoteForWindowsShell(command), args.map(quoteForWindowsShell), {
              env: options.env ?? process.env,
              windowsHide: false,
              shell: true,
            })
          : spawn(command, args, {
              env: options.env ?? process.env,
              windowsHide: false,
            });
      } catch (error) {
        resolve({
          exitCode: 1,
          stdout: '',
          stderr: error instanceof Error ? error.message : String(error),
        });
        return;
      }
      let stdout = '';
      let stderr = '';
      const timeout = setTimeout(() => {
        child.kill();
      }, options.timeoutMs ?? 15_000);
      child.stdout?.on('data', (chunk) => {
        stdout = tail(stdout + String(chunk), 4000);
      });
      child.stderr?.on('data', (chunk) => {
        stderr = tail(stderr + String(chunk), 4000);
      });
      child.on('error', (error) => {
        clearTimeout(timeout);
        resolve({ exitCode: 1, stdout, stderr: error.message });
      });
      child.on('close', (code) => {
        clearTimeout(timeout);
        resolve({ exitCode: code, stdout, stderr });
      });
    });
}

export function getRunCommand(options: CommonOptions): RunCommand {
  return options.runCommand ?? createDefaultRunCommand();
}
