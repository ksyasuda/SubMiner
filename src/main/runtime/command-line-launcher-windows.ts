import os from 'node:os';
import path from 'node:path';
import {
  envOf,
  getRunCommand,
  pathEntriesContain,
  splitPath,
  type CommonOptions,
  type WindowsPathOptions,
} from './command-line-launcher-deps';

const WINDOWS_PATH_MAX = 32767;

export function windowsLauncherPaths(options: CommonOptions & WindowsPathOptions): {
  binDir: string;
  installPath: string;
} {
  const localAppData = options.localAppData ?? envOf(options).LOCALAPPDATA;
  const base = localAppData ?? path.win32.join(options.homeDir ?? os.homedir(), 'AppData', 'Local');
  const binDir = path.win32.join(base, 'SubMiner', 'bin');
  return { binDir, installPath: path.win32.join(binDir, 'subminer.cmd') };
}

export function windowsShimContent(appExePath: string, launcherResourcePath: string): string {
  return [
    '@echo off',
    'setlocal',
    `set "SUBMINER_BINARY_PATH=${appExePath}"`,
    `bun "${launcherResourcePath}" %*`,
    '',
  ].join('\r\n');
}

export function shimMatchesCurrentInstall(
  content: string,
  appExePath: string,
  launcherResourcePath: string,
): boolean {
  const normalized = content.replaceAll('/', '\\').toLowerCase();
  return (
    normalized.includes(appExePath.replaceAll('/', '\\').toLowerCase()) &&
    normalized.includes(launcherResourcePath.replaceAll('/', '\\').toLowerCase())
  );
}

export function getUserPath(options: CommonOptions & WindowsPathOptions): string {
  return options.getUserPath?.() ?? envOf(options).Path ?? envOf(options).PATH ?? '';
}

async function setWindowsUserPath(options: CommonOptions & WindowsPathOptions, nextPath: string) {
  if (options.setUserPath) {
    await options.setUserPath(nextPath);
    return;
  }
  const escaped = nextPath.replaceAll("'", "''");
  await getRunCommand(options)('powershell', [
    '-NoProfile',
    '-ExecutionPolicy',
    'Bypass',
    '-Command',
    `[Environment]::SetEnvironmentVariable('Path', '${escaped}', 'User')`,
  ]);
}

async function broadcastEnvironmentChange(options: CommonOptions & WindowsPathOptions) {
  if (options.broadcastEnvironmentChange) {
    await options.broadcastEnvironmentChange();
    return;
  }
  await getRunCommand(options)('powershell', [
    '-NoProfile',
    '-ExecutionPolicy',
    'Bypass',
    '-Command',
    "$signature='[DllImport(\"user32.dll\",SetLastError=true,CharSet=CharSet.Auto)] public static extern IntPtr SendMessageTimeout(IntPtr hWnd,uint Msg,UIntPtr wParam,string lParam,uint fuFlags,uint uTimeout,out UIntPtr lpdwResult);'; Add-Type -MemberDefinition $signature -Name NativeMethods -Namespace Win32; $result=[UIntPtr]::Zero; [Win32.NativeMethods]::SendMessageTimeout([IntPtr]0xffff,0x1A,[UIntPtr]::Zero,'Environment',2,5000,[ref]$result) | Out-Null",
  ]);
}

export async function appendWindowsUserPathDir(
  dir: string,
  options: CommonOptions & WindowsPathOptions,
): Promise<string | null> {
  const current = getUserPath(options);
  const entries = splitPath(current, 'win32');
  if (pathEntriesContain(entries, dir, 'win32')) return null;
  const next = current.trim() ? `${current};${dir}` : dir;
  if (next.length > WINDOWS_PATH_MAX) {
    throw new Error('User PATH is too long to append the SubMiner launcher directory safely.');
  }
  await setWindowsUserPath(options, next);
  await broadcastEnvironmentChange(options);
  return next;
}

export function defaultBunRepairPath(options: CommonOptions & WindowsPathOptions): string {
  const userProfile =
    options.userProfile ?? envOf(options).USERPROFILE ?? options.homeDir ?? os.homedir();
  return path.win32.join(userProfile, '.bun', 'bin');
}
