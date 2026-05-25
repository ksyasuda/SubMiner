import fs from 'node:fs';
import os from 'node:os';
import { CliArgs, hasExplicitCommand, parseArgs, shouldStartApp } from './cli/args';
import { resolveConfigDir } from './config/path-resolution';

const BACKGROUND_ARG = '--background';
const START_ARG = '--start';
const PASSWORD_STORE_ARG = '--password-store';
const DEFAULT_LINUX_PASSWORD_STORE = 'gnome-libsecret';
const BACKGROUND_CHILD_ENV = 'SUBMINER_BACKGROUND_CHILD';
const TRANSPORTED_APP_ARGC_ENV = 'SUBMINER_APP_ARGC';
const TRANSPORTED_APP_ARG_PREFIX = 'SUBMINER_APP_ARG_';
const MAX_TRANSPORTED_APP_ARGS = 256;
const APP_NAME = 'SubMiner';
const MPV_LONG_OPTIONS_WITH_SEPARATE_VALUES = new Set([
  '--alang',
  '--audio-file',
  '--input-ipc-server',
  '--log-file',
  '--msg-level',
  '--profile',
  '--script',
  '--script-opts',
  '--scripts',
  '--slang',
  '--sub-file',
  '--sub-file-paths',
  '--title',
  '--volume',
  '--ytdl-format',
]);

type EarlyAppLike = {
  setName: (name: string) => void;
  setPath: (name: 'userData', value: string) => void;
};

type CommandLineLike = {
  appendSwitch: (name: string, value?: string) => void;
};

type EarlyAppPathOptions = {
  platform?: NodeJS.Platform;
  appDataDir?: string;
  xdgConfigHome?: string;
  homeDir?: string;
  existsSync?: (candidate: string) => boolean;
};

function removeLsfgLayer(env: NodeJS.ProcessEnv): void {
  if (typeof env.VK_INSTANCE_LAYERS === 'string' && /lsfg/i.test(env.VK_INSTANCE_LAYERS)) {
    delete env.VK_INSTANCE_LAYERS;
  }
}

function removePassiveStartupArgs(argv: string[]): string[] {
  const filtered: string[] = [];

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg) continue;

    if (arg === PASSWORD_STORE_ARG) {
      const value = argv[i + 1];
      if (value && !value.startsWith('--')) {
        i += 1;
      }
      continue;
    }

    if (arg.startsWith(`${PASSWORD_STORE_ARG}=`)) {
      continue;
    }

    filtered.push(arg);
  }

  return filtered;
}

function getPasswordStoreArg(argv: string[]): string | null {
  let resolved: string | null = null;
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg?.startsWith(PASSWORD_STORE_ARG)) {
      continue;
    }

    if (arg === PASSWORD_STORE_ARG) {
      const value = argv[i + 1];
      if (value && !value.startsWith('--')) {
        resolved = value.trim();
        i += 1;
      }
      continue;
    }

    const [prefix, value] = arg.split('=', 2);
    if (prefix === PASSWORD_STORE_ARG && value && value.trim().length > 0) {
      resolved = value.trim();
    }
  }
  return resolved;
}

function normalizePasswordStoreArg(value: string): string {
  const normalized = value.trim();
  if (normalized.toLowerCase() === 'gnome') {
    return DEFAULT_LINUX_PASSWORD_STORE;
  }
  return normalized;
}

export function resolveLinuxPasswordStoreValue(
  argv: string[],
  platform: NodeJS.Platform = process.platform,
): string | null {
  if (platform !== 'linux') return null;
  return normalizePasswordStoreArg(getPasswordStoreArg(argv) ?? DEFAULT_LINUX_PASSWORD_STORE);
}

export function applyEarlyLinuxCommandLineSwitches(
  commandLine: CommandLineLike,
  argv: string[],
  platform: NodeJS.Platform = process.platform,
): void {
  if (platform !== 'linux') return;
  commandLine.appendSwitch('enable-features', 'GlobalShortcutsPortal');
  commandLine.appendSwitch(
    'password-store',
    resolveLinuxPasswordStoreValue(argv, platform) ?? DEFAULT_LINUX_PASSWORD_STORE,
  );
}

function consumesLaunchMpvValue(token: string): boolean {
  return (
    token.startsWith('--') &&
    token !== '--' &&
    !token.includes('=') &&
    MPV_LONG_OPTIONS_WITH_SEPARATE_VALUES.has(token)
  );
}

function parseCliArgs(argv: string[]): CliArgs {
  return parseArgs(argv);
}

export function hasTransportedStartupArgs(env: NodeJS.ProcessEnv): boolean {
  return typeof env[TRANSPORTED_APP_ARGC_ENV] === 'string';
}

export function shouldForwardStartupArgvViaAppControl(
  argv: string[],
  env: NodeJS.ProcessEnv,
): boolean {
  if (env.ELECTRON_RUN_AS_NODE === '1') return false;

  const args = parseCliArgs(argv);
  if (args.help || args.appPing || args.launchMpv || args.generateConfig) return false;
  if (resolveStatsDaemonCommandAction(argv) !== null) return false;

  return hasExplicitCommand(args);
}

function readTransportedStartupArgs(env: NodeJS.ProcessEnv): string[] | null {
  const rawCount = env[TRANSPORTED_APP_ARGC_ENV];
  if (rawCount === undefined) {
    return null;
  }

  const count = Number(rawCount);
  if (!Number.isInteger(count) || count < 0 || count > MAX_TRANSPORTED_APP_ARGS) {
    return null;
  }

  const args: string[] = [];
  for (let index = 0; index < count; index += 1) {
    const value = env[`${TRANSPORTED_APP_ARG_PREFIX}${index}`];
    if (typeof value !== 'string') {
      return null;
    }
    args.push(value);
  }
  return args;
}

export function normalizeStartupArgv(argv: string[], env: NodeJS.ProcessEnv): string[] {
  if (env.ELECTRON_RUN_AS_NODE === '1') return argv;

  const transportedArgs = readTransportedStartupArgs(env);
  if (transportedArgs) {
    if (removePassiveStartupArgs(transportedArgs).length === 0) {
      if (process.platform === 'win32') {
        return [argv[0] ?? APP_NAME, ...transportedArgs, START_ARG];
      }
      return [argv[0] ?? APP_NAME, ...transportedArgs, START_ARG, BACKGROUND_ARG];
    }
    return [argv[0] ?? APP_NAME, ...transportedArgs];
  }

  const effectiveArgs = removePassiveStartupArgs(argv.slice(1));
  if (effectiveArgs.length === 0) {
    if (process.platform === 'win32') {
      return [...argv, START_ARG];
    }
    return [...argv, START_ARG, BACKGROUND_ARG];
  }

  if (
    effectiveArgs.length === 1 &&
    effectiveArgs[0] === BACKGROUND_ARG &&
    !argv.includes(START_ARG)
  ) {
    return [...argv, START_ARG];
  }

  return argv;
}

export function configureEarlyAppPaths(app: EarlyAppLike, options?: EarlyAppPathOptions): string {
  const userDataPath = resolveConfigDir({
    platform: options?.platform ?? process.platform,
    appDataDir: options?.appDataDir ?? process.env.APPDATA,
    xdgConfigHome: options?.xdgConfigHome ?? process.env.XDG_CONFIG_HOME,
    homeDir: options?.homeDir ?? os.homedir(),
    existsSync: options?.existsSync ?? fs.existsSync,
  });

  app.setName(APP_NAME);
  app.setPath('userData', userDataPath);

  return userDataPath;
}

export function shouldDetachBackgroundLaunch(argv: string[], env: NodeJS.ProcessEnv): boolean {
  if (env.ELECTRON_RUN_AS_NODE === '1') return false;
  if (!argv.includes(BACKGROUND_ARG)) return false;
  if (env[BACKGROUND_CHILD_ENV] === '1') return false;
  return true;
}

export function shouldHandleHelpOnlyAtEntry(argv: string[], env: NodeJS.ProcessEnv): boolean {
  if (env.ELECTRON_RUN_AS_NODE === '1') return false;
  const args = parseCliArgs(argv);
  return args.help && !shouldStartApp(args);
}

export function shouldHandleLaunchMpvAtEntry(argv: string[], env: NodeJS.ProcessEnv): boolean {
  if (env.ELECTRON_RUN_AS_NODE === '1') return false;
  return parseCliArgs(argv).launchMpv;
}

export function shouldHandleStatsDaemonCommandAtEntry(
  argv: string[],
  env: NodeJS.ProcessEnv,
): boolean {
  if (env.ELECTRON_RUN_AS_NODE === '1') return false;
  return resolveStatsDaemonCommandAction(argv) !== null;
}

export function resolveStatsDaemonCommandAction(argv: string[]): 'start' | 'stop' | null {
  if (argv.includes('--stats-daemon-stop') || argv.includes('--stats-stop')) {
    return 'stop';
  }
  if (argv.includes('--stats-daemon-start') || argv.includes('--stats-background')) {
    return 'start';
  }
  return null;
}

export function normalizeLaunchMpvTargets(argv: string[]): string[] {
  const launchMpvIndex = argv.findIndex((arg) => arg === '--launch-mpv');
  if (launchMpvIndex < 0) {
    return [];
  }

  const targets: string[] = [];

  let parsingTargets = false;
  for (let i = launchMpvIndex + 1; i < argv.length; i += 1) {
    const token = argv[i];
    if (!token) continue;

    if (parsingTargets) {
      targets.push(token);
      continue;
    }

    if (token === '--') {
      parsingTargets = true;
      continue;
    }

    if (token.startsWith('--')) {
      if (consumesLaunchMpvValue(token) && i + 1 < argv.length) {
        const value = argv[i + 1];
        if (value && !value.startsWith('-')) {
          i += 1;
        }
      }
      continue;
    }

    if (token.startsWith('-')) {
      continue;
    }

    parsingTargets = true;
    targets.push(token);
  }

  return targets;
}

export function normalizeLaunchMpvExtraArgs(argv: string[]): string[] {
  const launchMpvIndex = argv.findIndex((arg) => arg === '--launch-mpv');
  if (launchMpvIndex < 0) {
    return [];
  }

  const extraArgs: string[] = [];
  for (let i = launchMpvIndex + 1; i < argv.length; i += 1) {
    const token = argv[i];
    if (!token) continue;
    if (token === '--') {
      break;
    }
    if (token.startsWith('--')) {
      extraArgs.push(token);
      if (consumesLaunchMpvValue(token) && i + 1 < argv.length) {
        const value = argv[i + 1];
        if (value && !value.startsWith('-')) {
          extraArgs.push(value);
          i += 1;
        }
      }
      continue;
    }
    if (token.startsWith('-')) {
      extraArgs.push(token);
      continue;
    }
    if (!token.startsWith('-')) {
      break;
    }
  }
  return extraArgs;
}

export function sanitizeStartupEnv(baseEnv: NodeJS.ProcessEnv): NodeJS.ProcessEnv {
  const env = { ...baseEnv };
  if (!env.NODE_NO_WARNINGS) {
    env.NODE_NO_WARNINGS = '1';
  }
  removeLsfgLayer(env);
  return env;
}

export function sanitizeHelpEnv(baseEnv: NodeJS.ProcessEnv): NodeJS.ProcessEnv {
  return sanitizeStartupEnv(baseEnv);
}

export function sanitizeLaunchMpvEnv(baseEnv: NodeJS.ProcessEnv): NodeJS.ProcessEnv {
  return sanitizeStartupEnv(baseEnv);
}

export function sanitizeBackgroundEnv(baseEnv: NodeJS.ProcessEnv): NodeJS.ProcessEnv {
  const env = sanitizeStartupEnv(baseEnv);
  env[BACKGROUND_CHILD_ENV] = '1';
  return env;
}
