import { spawn } from 'node:child_process';

const BACKGROUND_ARG = '--background';
const BACKGROUND_CHILD_ENV = 'SUBMINER_BACKGROUND_CHILD';

function shouldDetachBackgroundLaunch(argv: string[], env: NodeJS.ProcessEnv): boolean {
  if (env.ELECTRON_RUN_AS_NODE === '1') return false;
  if (!argv.includes(BACKGROUND_ARG)) return false;
  if (env[BACKGROUND_CHILD_ENV] === '1') return false;
  return true;
}

function sanitizeBackgroundEnv(baseEnv: NodeJS.ProcessEnv): NodeJS.ProcessEnv {
  const env = { ...baseEnv };
  env[BACKGROUND_CHILD_ENV] = '1';
  if (!env.NODE_NO_WARNINGS) {
    env.NODE_NO_WARNINGS = '1';
  }
  if (typeof env.VK_INSTANCE_LAYERS === 'string' && /lsfg/i.test(env.VK_INSTANCE_LAYERS)) {
    delete env.VK_INSTANCE_LAYERS;
  }
  return env;
}

if (shouldDetachBackgroundLaunch(process.argv, process.env)) {
  const child = spawn(process.execPath, process.argv.slice(1), {
    detached: true,
    stdio: 'ignore',
    env: sanitizeBackgroundEnv(process.env),
  });
  child.unref();
  process.exit(0);
}

require('./main.js');
