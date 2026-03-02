import { spawn } from 'node:child_process';
import { printHelp } from './cli/help';
import {
  sanitizeBackgroundEnv,
  sanitizeHelpEnv,
  shouldDetachBackgroundLaunch,
  shouldHandleHelpOnlyAtEntry,
} from './main-entry-runtime';

const DEFAULT_TEXTHOOKER_PORT = 5174;

if (shouldDetachBackgroundLaunch(process.argv, process.env)) {
  const child = spawn(process.execPath, process.argv.slice(1), {
    detached: true,
    stdio: 'ignore',
    env: sanitizeBackgroundEnv(process.env),
  });
  child.unref();
  process.exit(0);
}

if (shouldHandleHelpOnlyAtEntry(process.argv, process.env)) {
  const sanitizedEnv = sanitizeHelpEnv(process.env);
  process.env.NODE_NO_WARNINGS = sanitizedEnv.NODE_NO_WARNINGS;
  if (!sanitizedEnv.VK_INSTANCE_LAYERS) {
    delete process.env.VK_INSTANCE_LAYERS;
  }
  printHelp(DEFAULT_TEXTHOOKER_PORT);
  process.exit(0);
}

require('./main.js');
