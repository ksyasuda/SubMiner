import { CliArgs, parseArgs, shouldStartApp } from './cli/args';

const BACKGROUND_ARG = '--background';
const START_ARG = '--start';
const PASSWORD_STORE_ARG = '--password-store';
const BACKGROUND_CHILD_ENV = 'SUBMINER_BACKGROUND_CHILD';

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

function parseCliArgs(argv: string[]): CliArgs {
  return parseArgs(argv);
}

export function normalizeStartupArgv(argv: string[], env: NodeJS.ProcessEnv): string[] {
  if (env.ELECTRON_RUN_AS_NODE === '1') return argv;

  const effectiveArgs = removePassiveStartupArgs(argv.slice(1));
  if (effectiveArgs.length === 0) {
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

export function sanitizeBackgroundEnv(baseEnv: NodeJS.ProcessEnv): NodeJS.ProcessEnv {
  const env = sanitizeStartupEnv(baseEnv);
  env[BACKGROUND_CHILD_ENV] = '1';
  return env;
}
