import { randomUUID } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { windowsLauncherBootstrapContent } from './windows-launcher-bootstrap';
import { MANAGED_LAUNCHER_MARKER, posixLauncherBootstrapContent } from './posix-launcher-bootstrap';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  envOf,
  existsSyncOf,
  pathModuleFor,
  platformOf,
  type CommonOptions,
  type WindowsPathOptions,
} from './command-line-launcher-deps';

export { MANAGED_LAUNCHER_MARKER, shellQuote } from './posix-launcher-bootstrap';

export function isManagedLauncher(content: string): boolean {
  const lines = content.split(/\r?\n/, 3);
  return (
    (lines[0] === '#!/bin/sh' && lines[1] === `# ${MANAGED_LAUNCHER_MARKER}`) ||
    (lines[0] === '@echo off' && lines[1] === `rem ${MANAGED_LAUNCHER_MARKER}`)
  );
}

export function managedLauncherContent(options: {
  platform: NodeJS.Platform;
  appPath: string;
}): string {
  if (options.platform === 'win32') return windowsLauncherBootstrapContent(options.appPath);
  return posixLauncherBootstrapContent(options.appPath);
}

export function managedLauncherPaths(options: CommonOptions) {
  const platform = platformOf(options);
  const platformPath = pathModuleFor(platform);
  const env = envOf(options);
  const home = options.homeDir ?? os.homedir();
  const dataHome = env.XDG_DATA_HOME;
  const directory = platformPath.join(
    dataHome && platformPath.isAbsolute(dataHome)
      ? dataHome
      : platformPath.join(home, '.local', 'share'),
    'SubMiner',
    'launcher',
  );
  return {
    directory,
    bunPath: platformPath.join(directory, 'bun'),
    scriptPath: platformPath.join(directory, 'subminer'),
    versionPath: platformPath.join(directory, 'version'),
    fingerprintPath: platformPath.join(directory, 'fingerprint'),
    appPathFile: platformPath.join(directory, 'app-path'),
  };
}

function absoluteWindowsPath(candidate: string, label: string): string {
  const normalized = path.win32.normalize(candidate);
  if (!path.win32.isAbsolute(normalized)) {
    throw new Error(`${label} must be an absolute Windows path: ${candidate}`);
  }
  return normalized;
}

export function windowsManagedRuntimePaths(options: CommonOptions & WindowsPathOptions) {
  const env = envOf(options);
  const userProfile = options.userProfile?.trim() || env.USERPROFILE?.trim() || os.homedir();
  const localAppData =
    options.localAppData?.trim() ||
    env.LOCALAPPDATA?.trim() ||
    path.win32.join(userProfile, 'AppData', 'Local');
  const rootDirectory = path.win32.join(
    absoluteWindowsPath(localAppData, 'Windows local app data directory'),
    'SubMiner',
    'launcher-runtime',
  );
  const version = options.appVersion?.trim() || 'development';
  if (
    version === '.' ||
    version === '..' ||
    /[<>:"/\\|?*\u0000-\u001f]/.test(version) ||
    /[ .]$/.test(version)
  ) {
    throw new Error(`SubMiner version is not a valid directory name: ${version}`);
  }
  const directory = path.win32.join(rootDirectory, version);
  return {
    rootDirectory,
    directory,
    bunPath: path.win32.join(directory, 'bun.exe'),
  };
}

export function cleanupOldWindowsManagedRuntimes(
  options: CommonOptions & WindowsPathOptions,
): void {
  const paths = windowsManagedRuntimePaths(options);
  try {
    for (const entry of fs.readdirSync(paths.rootDirectory, { withFileTypes: true })) {
      if (!entry.isDirectory() || entry.name === path.win32.basename(paths.directory)) continue;
      const oldDirectory = path.win32.join(paths.rootDirectory, entry.name);
      try {
        fs.rmSync(path.win32.join(oldDirectory, 'bun.exe'), { force: true });
      } catch {
        continue;
      }
      try {
        fs.rmSync(oldDirectory, { recursive: true, force: true });
      } catch {
        // Cleanup is best-effort after proving the old runtime is not locked.
      }
    }
  } catch {
    // The runtime root can be absent before the first managed launcher install.
  }
}

function stageWindowsManagedRuntime(
  options: CommonOptions &
    WindowsPathOptions & {
      bundledBunPath: string;
      force?: boolean;
    },
): string {
  const paths = windowsManagedRuntimePaths(options);
  const exists = existsSyncOf(options);
  const mkdir = options.mkdirSync ?? fs.mkdirSync;
  const copy = options.copyFileSync ?? fs.copyFileSync;

  mkdir(paths.directory, { recursive: true });
  if (options.force || !exists(paths.bunPath)) {
    const stagingPath = `${paths.bunPath}.${randomUUID()}.tmp`;
    try {
      copy(options.bundledBunPath, stagingPath);
      if (exists(paths.bunPath)) fs.rmSync(paths.bunPath, { force: true });
      fs.renameSync(stagingPath, paths.bunPath);
    } finally {
      fs.rmSync(stagingPath, { force: true });
    }
  }

  const packagedLicenses = path.join(path.dirname(options.bundledBunPath), 'licenses');
  const cachedLicenses = path.win32.join(paths.directory, 'licenses');
  if (exists(packagedLicenses) && (options.force || !exists(cachedLicenses))) {
    fs.cpSync(packagedLicenses, cachedLicenses, { recursive: true, force: true });
  }

  return paths.bunPath;
}

// Keep ephemeral or replaceable executables outside the installed app. Linux
// also copies the script because AppImage resources disappear when the app exits.
export function stageManagedLauncher(
  options: CommonOptions &
    WindowsPathOptions & {
      bundledBunPath: string;
      launcherResourcePath: string;
      force?: boolean;
    },
) {
  const platform = platformOf(options);
  if (platform === 'win32') {
    return {
      bunPath: stageWindowsManagedRuntime(options),
      scriptPath: options.launcherResourcePath,
    };
  }
  if (platform !== 'linux') {
    return { bunPath: options.bundledBunPath, scriptPath: options.launcherResourcePath };
  }
  const paths = managedLauncherPaths(options);
  const exists = existsSyncOf(options);
  const read = options.readFileSync ?? fs.readFileSync;
  const write = options.writeFileSync ?? fs.writeFileSync;
  const copy = options.copyFileSync ?? fs.copyFileSync;
  const mkdir = options.mkdirSync ?? fs.mkdirSync;
  const chmod = options.chmodSync ?? fs.chmodSync;
  const version = options.appVersion ?? 'development';
  const appPath = envOf(options).APPIMAGE ?? options.appExePath;
  const fingerprint = appPath
    ? execFileSync('stat', ['-Lc', '%d:%i:%s:%y:%z', '--', appPath], {
        encoding: 'utf8',
        env: { ...envOf(options), PATH: `/usr/bin:/bin:${envOf(options).PATH ?? ''}` },
      }).trim()
    : '';
  if (
    !options.force &&
    exists(paths.versionPath) &&
    read(paths.versionPath, 'utf8') === version &&
    exists(paths.fingerprintPath) &&
    read(paths.fingerprintPath, 'utf8') === `${fingerprint}\n` &&
    exists(paths.appPathFile) &&
    read(paths.appPathFile, 'utf8') === `${appPath ?? ''}\n` &&
    exists(paths.bunPath) &&
    exists(paths.scriptPath)
  )
    return paths;

  mkdir(paths.directory, { recursive: true });
  const staging = fs.mkdtempSync(path.join(paths.directory, '.stage-'));
  try {
    copy(options.bundledBunPath, path.join(staging, 'bun'));
    chmod(path.join(staging, 'bun'), 0o755);
    copy(options.launcherResourcePath, path.join(staging, 'subminer'));
    const notices = path.join(path.dirname(options.bundledBunPath), 'licenses');
    if (exists(notices))
      fs.cpSync(notices, path.join(paths.directory, 'licenses'), { recursive: true });
    write(path.join(staging, 'version'), version);
    write(path.join(staging, 'app-path'), `${appPath ?? ''}\n`);
    write(path.join(staging, 'fingerprint'), `${fingerprint}\n`);
    for (const name of ['bun', 'subminer', 'version', 'app-path', 'fingerprint']) {
      fs.renameSync(path.join(staging, name), path.join(paths.directory, name));
    }
  } finally {
    fs.rmSync(staging, { recursive: true, force: true });
  }
  return paths;
}
