import fs from 'node:fs';
import os from 'node:os';
import { CliArgs, parseArgs, shouldStartApp } from './cli/args';
import { resolveConfigDir } from './config/path-resolution';

const BACKGROUND_ARG = '--background';
const START_ARG = '--start';
const PASSWORD_STORE_ARG = '--password-store';
const BACKGROUND_CHILD_ENV = 'SUBMINER_BACKGROUND_CHILD';
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

export function normalizeStartupArgv(argv: string[], env: NodeJS.ProcessEnv): string[] {
  if (env.ELECTRON_RUN_AS_NODE === '1') return argv;

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
  return argv.includes('--stats-daemon-start') || argv.includes('--stats-daemon-stop');
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
