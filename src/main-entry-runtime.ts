import { CliArgs, parseArgs, shouldStartApp } from './cli/args';

const BACKGROUND_ARG = '--background';
const BACKGROUND_CHILD_ENV = 'SUBMINER_BACKGROUND_CHILD';

function removeLsfgLayer(env: NodeJS.ProcessEnv): void {
  if (typeof env.VK_INSTANCE_LAYERS === 'string' && /lsfg/i.test(env.VK_INSTANCE_LAYERS)) {
    delete env.VK_INSTANCE_LAYERS;
  }
}

function parseCliArgs(argv: string[]): CliArgs {
  return parseArgs(argv);
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

export function sanitizeHelpEnv(baseEnv: NodeJS.ProcessEnv): NodeJS.ProcessEnv {
  const env = { ...baseEnv };
  if (!env.NODE_NO_WARNINGS) {
    env.NODE_NO_WARNINGS = '1';
  }
  removeLsfgLayer(env);
  return env;
}

export function sanitizeBackgroundEnv(baseEnv: NodeJS.ProcessEnv): NodeJS.ProcessEnv {
  const env = sanitizeHelpEnv(baseEnv);
  env[BACKGROUND_CHILD_ENV] = '1';
  return env;
}
