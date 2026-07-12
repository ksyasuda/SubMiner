import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  accessSyncOf,
  envOf,
  existsSyncOf,
  failureMessage,
  findCommand,
  getRunCommand,
  normalizePathForCompare,
  pathEntriesContain,
  pathModuleFor,
  platformOf,
  splitPath,
  type CommonOptions,
  type WindowsPathOptions,
} from './command-line-launcher-deps';
import {
  appendWindowsUserPathDir,
  defaultBunRepairPath,
  shimMatchesCurrentInstall,
  windowsLauncherPaths,
  windowsShimContent,
} from './command-line-launcher-windows';

export type { RunCommand, RunCommandResult } from './command-line-launcher-deps';

export type ToolStatus = 'ready' | 'missing' | 'installing' | 'failed';

export type LauncherInstallStatus =
  | 'ready'
  | 'installed_bun_missing'
  | 'not_installed'
  | 'not_on_path'
  | 'shadowed'
  | 'not_installable'
  | 'failed';

export interface BunSnapshot {
  status: ToolStatus;
  commandPath: string | null;
  version: string | null;
  installMethod: 'winget' | 'scoop' | 'homebrew' | 'official-script' | null;
  installCommand: string[] | null;
  message: string | null;
}

export interface LauncherSnapshot {
  status: LauncherInstallStatus;
  commandPath: string | null;
  installPath: string | null;
  pathDir: string | null;
  shadowedBy: string | null;
  message: string | null;
}

export interface CommandLineLauncherSnapshot {
  supported: boolean;
  bun: BunSnapshot;
  launcher: LauncherSnapshot;
}

const BUN_OFFICIAL_POSIX_COMMAND = ['bash', '-lc', 'curl -fsSL https://bun.com/install | bash'];
const BUN_OFFICIAL_WINDOWS_COMMAND = [
  'powershell',
  '-NoProfile',
  '-ExecutionPolicy',
  'Bypass',
  '-Command',
  'irm bun.sh/install.ps1 | iex',
];
const INSTALL_TIMEOUT_MS = 10 * 60 * 1000;
const COMMAND_TIMEOUT_MS = 15 * 1000;
const MACOS_HOMEBREW_PATH_DIRS = ['/opt/homebrew/bin'];

function installMethodForCommand(command: string[] | null): BunSnapshot['installMethod'] {
  if (!command) return null;
  const executablePath = command[0];
  if (!executablePath) return null;
  const executable = path.basename(executablePath).toLowerCase();
  const windowsExecutable = path.win32.basename(executablePath).toLowerCase();
  if (windowsExecutable === 'winget.exe') return 'winget';
  if (windowsExecutable === 'scoop.cmd') return 'scoop';
  if (executable === 'brew' || windowsExecutable === 'brew') return 'homebrew';
  return 'official-script';
}

export function resolveBunInstallCommand(
  options: CommonOptions = {},
): BunSnapshot['installCommand'] {
  const platform = platformOf(options);
  if (platform === 'win32') {
    const winget = findCommand('winget.exe', options);
    if (winget) {
      return [
        winget,
        'install',
        '--id',
        'Oven-sh.Bun',
        '--exact',
        '--accept-package-agreements',
        '--accept-source-agreements',
      ];
    }
    const scoop = findCommand('scoop.cmd', options);
    if (scoop) return [scoop, 'install', 'bun'];
    return BUN_OFFICIAL_WINDOWS_COMMAND;
  }

  const brew = findCommand('brew', options);
  if (platform === 'darwin' && brew) return [brew, 'install', 'bun'];
  if (platform === 'linux' && brew) return [brew, 'install', 'bun'];
  return BUN_OFFICIAL_POSIX_COMMAND;
}

export async function detectBun(options: CommonOptions = {}): Promise<BunSnapshot> {
  const bunPath = findCommand('bun', options);
  const installCommand = resolveBunInstallCommand(options);
  if (!bunPath) {
    return {
      status: 'missing',
      commandPath: null,
      version: null,
      installMethod: installMethodForCommand(installCommand),
      installCommand,
      message: null,
    };
  }

  const result = await getRunCommand(options)(bunPath, ['--version'], {
    timeoutMs: COMMAND_TIMEOUT_MS,
    env: envOf(options) as NodeJS.ProcessEnv,
  });
  if (result.exitCode === 0) {
    return {
      status: 'ready',
      commandPath: bunPath,
      version: result.stdout.trim() || null,
      installMethod: null,
      installCommand: null,
      message: null,
    };
  }

  return {
    status: 'failed',
    commandPath: bunPath,
    version: null,
    installMethod: installMethodForCommand(installCommand),
    installCommand,
    message: failureMessage(result, 'bun --version failed'),
  };
}

export function resolveLauncherResourcePath(options: CommonOptions): string {
  const platformPath = pathModuleFor(platformOf(options));
  if (options.launcherResourcePath) return options.launcherResourcePath;
  const resourcesPath =
    options.resourcesPath ?? (process as typeof process & { resourcesPath?: string }).resourcesPath;
  const packaged = resourcesPath ? platformPath.join(resourcesPath, 'launcher', 'subminer') : null;
  if (packaged && existsSyncOf(options)(packaged)) return packaged;
  return platformPath.join(options.cwd ?? process.cwd(), 'dist', 'launcher', 'subminer');
}

function isWritableDir(candidate: string, options: CommonOptions): boolean {
  try {
    if (!existsSyncOf(options)(candidate)) return false;
    accessSyncOf(options)(candidate, fs.constants.W_OK);
    return true;
  } catch {
    return false;
  }
}

function collectPathDirs(options: CommonOptions): string[] {
  const platform = platformOf(options);
  const dirs: string[] = [];
  const add = (dir: string) => {
    if (!pathEntriesContain(dirs, dir, platform)) dirs.push(dir);
  };
  splitPath(envOf(options).PATH, platform).forEach(add);
  return dirs;
}

export async function resolveLauncherInstallTarget(
  options: CommonOptions & WindowsPathOptions = {},
): Promise<LauncherSnapshot> {
  const platform = platformOf(options);
  if (platform === 'win32') {
    const { binDir, installPath } = windowsLauncherPaths(options);
    return {
      status: existsSyncOf(options)(installPath) ? 'ready' : 'not_installed',
      commandPath: existsSyncOf(options)(installPath) ? installPath : null,
      installPath,
      pathDir: binDir,
      shadowedBy: null,
      message: null,
    };
  }

  const homeDir = options.homeDir ?? os.homedir();
  const pathDirs = collectPathDirs(options);
  const preferred =
    platform === 'darwin'
      ? [
          '/opt/homebrew/bin',
          '/usr/local/bin',
          path.posix.join(homeDir, '.local', 'bin'),
          path.posix.join(homeDir, 'bin'),
        ]
      : [
          path.posix.join(homeDir, '.local', 'bin'),
          path.posix.join(homeDir, 'bin'),
          '/usr/local/bin',
        ];
  const manualPreferred =
    platform === 'darwin'
      ? [
          path.posix.join(homeDir, '.local', 'bin'),
          path.posix.join(homeDir, 'bin'),
          '/usr/local/bin',
        ]
      : preferred;
  const installCandidates = [...manualPreferred, ...pathDirs].filter(
    (dir, index, all) =>
      all.findIndex(
        (other) =>
          normalizePathForCompare(other, platform) === normalizePathForCompare(dir, platform),
      ) === index,
  );
  const installedPreferred = pathDirs.find((dir) => {
    if (!pathEntriesContain(preferred, dir, platform)) return false;
    return existsSyncOf(options)(path.posix.join(dir, 'subminer'));
  });
  if (installedPreferred) {
    const installPath = path.posix.join(installedPreferred, 'subminer');
    return {
      status: 'ready',
      commandPath: installPath,
      installPath,
      pathDir: installedPreferred,
      shadowedBy: null,
      message: null,
    };
  }
  const selected = installCandidates.find(
    (dir) =>
      (platform !== 'darwin' || !pathEntriesContain(MACOS_HOMEBREW_PATH_DIRS, dir, platform)) &&
      pathEntriesContain(pathDirs, dir, platform) &&
      isWritableDir(dir, options),
  );
  if (!selected) {
    return {
      status: 'not_installable',
      commandPath: null,
      installPath: null,
      pathDir: null,
      shadowedBy: null,
      message: 'No writable directory was found on your command-line PATH.',
    };
  }
  const installPath = path.posix.join(selected, 'subminer');
  return {
    status: existsSyncOf(options)(installPath) ? 'ready' : 'not_installed',
    commandPath: existsSyncOf(options)(installPath) ? installPath : null,
    installPath,
    pathDir: selected,
    shadowedBy: null,
    message: null,
  };
}

export async function detectLauncher(
  options: CommonOptions & WindowsPathOptions & { bunSnapshot?: BunSnapshot } = {},
): Promise<LauncherSnapshot> {
  const platform = platformOf(options);
  const target = await resolveLauncherInstallTarget(options);
  if (target.status === 'not_installable') return target;
  const expectedPath = target.installPath;
  if (!expectedPath) return target;
  const platformPath = pathModuleFor(platform);
  const launcherResourcePath = resolveLauncherResourcePath(options);
  const appExePath = options.appExePath ?? process.execPath;

  if (platform === 'win32' && existsSyncOf(options)(expectedPath)) {
    const content = String((options.readFileSync ?? fs.readFileSync)(expectedPath, 'utf8'));
    if (!shimMatchesCurrentInstall(content, appExePath, launcherResourcePath)) {
      return {
        ...target,
        status: 'not_installed',
        commandPath: null,
        message: 'Installed launcher points at a previous SubMiner install; reinstall to refresh.',
      };
    }
  }

  const commandPath = findCommand('subminer', options);
  const expectedNormalized = normalizePathForCompare(expectedPath, platform, platformPath);
  if (
    commandPath &&
    normalizePathForCompare(commandPath, platform, platformPath) !== expectedNormalized
  ) {
    return { ...target, status: 'shadowed', commandPath: expectedPath, shadowedBy: commandPath };
  }
  if (!existsSyncOf(options)(expectedPath))
    return { ...target, status: 'not_installed', commandPath: null };
  if (!commandPath) {
    return {
      ...target,
      status: 'not_on_path',
      commandPath: expectedPath,
      message: 'Launcher exists but its directory is not on PATH.',
    };
  }

  const bunSnapshot = options.bunSnapshot ?? (await detectBun(options));
  if (bunSnapshot.status !== 'ready') {
    return {
      ...target,
      status: 'installed_bun_missing',
      commandPath,
      message: 'Launcher is installed, but Bun is missing. Install Bun, then open a new terminal.',
    };
  }

  const result = await getRunCommand(options)(commandPath, ['--help'], {
    timeoutMs: COMMAND_TIMEOUT_MS,
    env: envOf(options) as NodeJS.ProcessEnv,
  });
  if (result.exitCode !== 0) {
    return {
      ...target,
      status: 'failed',
      commandPath,
      message: failureMessage(result, 'subminer --help failed'),
    };
  }
  return { ...target, status: 'ready', commandPath, message: null };
}

export async function installLauncher(
  options: CommonOptions & WindowsPathOptions & { bunSnapshot?: BunSnapshot } = {},
): Promise<LauncherSnapshot> {
  const platform = platformOf(options);
  const target = await resolveLauncherInstallTarget(options);
  if (!target.installPath || !target.pathDir) return target;
  const launcherResourcePath = resolveLauncherResourcePath(options);
  if (!existsSyncOf(options)(launcherResourcePath)) {
    return {
      ...target,
      status: 'failed',
      message: `Packaged launcher resource is missing: ${launcherResourcePath}`,
    };
  }

  if (platform === 'win32') {
    (options.mkdirSync ?? fs.mkdirSync)(target.pathDir, { recursive: true });
    (options.writeFileSync ?? fs.writeFileSync)(
      target.installPath,
      windowsShimContent(options.appExePath ?? process.execPath, launcherResourcePath),
      'utf8',
    );
    try {
      const nextPath = await appendWindowsUserPathDir(target.pathDir, options);
      if (nextPath && options.env) {
        options.env.PATH = nextPath;
        options.env.Path = nextPath;
      }
    } catch (error) {
      return {
        ...target,
        status: 'failed',
        message: error instanceof Error ? error.message : String(error),
      };
    }
  } else {
    (options.copyFileSync ?? fs.copyFileSync)(launcherResourcePath, target.installPath);
    (options.chmodSync ?? fs.chmodSync)(target.installPath, 0o755);
  }
  return detectLauncher(options);
}

export async function installBun(
  options: CommonOptions & WindowsPathOptions = {},
): Promise<BunSnapshot> {
  const platform = platformOf(options);
  if (platform === 'win32') {
    const bunDir = defaultBunRepairPath(options);
    const bunExe = path.win32.join(bunDir, 'bun.exe');
    if (existsSyncOf(options)(bunExe) && !findCommand('bun.exe', options)) {
      try {
        await appendWindowsUserPathDir(bunDir, options);
        return {
          status: 'ready',
          commandPath: bunExe,
          version: null,
          installMethod: null,
          installCommand: null,
          message: 'Bun PATH repaired. Open a new terminal.',
        };
      } catch (error) {
        return {
          status: 'failed',
          commandPath: bunExe,
          version: null,
          installMethod: null,
          installCommand: null,
          message: error instanceof Error ? error.message : String(error),
        };
      }
    }
  }

  const installCommand = resolveBunInstallCommand(options);
  if (!installCommand || installCommand.length === 0) {
    return {
      status: 'missing',
      commandPath: null,
      version: null,
      installMethod: null,
      installCommand: null,
      message: 'No Bun install command is available for this platform.',
    };
  }

  const command = installCommand[0]!;
  const args = installCommand.slice(1);
  const result = await getRunCommand(options)(command, args, {
    timeoutMs: INSTALL_TIMEOUT_MS,
    env: envOf(options) as NodeJS.ProcessEnv,
  });
  if (result.exitCode !== 0) {
    return {
      status: 'failed',
      commandPath: null,
      version: null,
      installMethod: installMethodForCommand(installCommand),
      installCommand,
      message: failureMessage(result, 'Bun install failed'),
    };
  }

  const detected = await detectBun(options);
  if (detected.status === 'ready') {
    return { ...detected, message: 'Bun installed. Open a new terminal.' };
  }
  return {
    ...detected,
    status: 'missing',
    message:
      platform === 'win32'
        ? 'Bun installed, but this process cannot see it on PATH yet. Open a new terminal.'
        : 'Bun installed, but is not on PATH for this shell. Add ~/.bun/bin to PATH if needed.',
  };
}

export async function detectCommandLineLauncher(
  options: CommonOptions & WindowsPathOptions = {},
): Promise<CommandLineLauncherSnapshot> {
  const platform = platformOf(options);
  const supported = platform === 'win32' || platform === 'linux' || platform === 'darwin';
  if (!supported) {
    return {
      supported: false,
      bun: {
        status: 'missing',
        commandPath: null,
        version: null,
        installMethod: null,
        installCommand: null,
        message: 'Command-line launcher setup is not supported on this platform.',
      },
      launcher: {
        status: 'not_installable',
        commandPath: null,
        installPath: null,
        pathDir: null,
        shadowedBy: null,
        message: 'Command-line launcher setup is not supported on this platform.',
      },
    };
  }
  const bun = await detectBun(options);
  const launcher = await detectLauncher({ ...options, bunSnapshot: bun });
  return { supported, bun, launcher };
}
