import { spawn } from 'node:child_process';
import { app, dialog } from 'electron';
import { printHelp } from './cli/help';
import {
  normalizeLaunchMpvTargets,
  normalizeStartupArgv,
  sanitizeStartupEnv,
  sanitizeBackgroundEnv,
  sanitizeHelpEnv,
  sanitizeLaunchMpvEnv,
  shouldDetachBackgroundLaunch,
  shouldHandleHelpOnlyAtEntry,
  shouldHandleLaunchMpvAtEntry,
} from './main-entry-runtime';
import { createWindowsMpvLaunchDeps, launchWindowsMpv } from './main/runtime/windows-mpv-launch';

const DEFAULT_TEXTHOOKER_PORT = 5174;

function applySanitizedEnv(sanitizedEnv: NodeJS.ProcessEnv): void {
  if (sanitizedEnv.NODE_NO_WARNINGS) {
    process.env.NODE_NO_WARNINGS = sanitizedEnv.NODE_NO_WARNINGS;
  }

  if (sanitizedEnv.VK_INSTANCE_LAYERS) {
    process.env.VK_INSTANCE_LAYERS = sanitizedEnv.VK_INSTANCE_LAYERS;
  } else {
    delete process.env.VK_INSTANCE_LAYERS;
  }
}

process.argv = normalizeStartupArgv(process.argv, process.env);
applySanitizedEnv(sanitizeStartupEnv(process.env));

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

if (shouldHandleLaunchMpvAtEntry(process.argv, process.env)) {
  const sanitizedEnv = sanitizeLaunchMpvEnv(process.env);
  applySanitizedEnv(sanitizedEnv);
  void app.whenReady().then(() => {
    const result = launchWindowsMpv(
      normalizeLaunchMpvTargets(process.argv),
      createWindowsMpvLaunchDeps({
        getEnv: (name) => process.env[name],
        showError: (title, content) => {
          dialog.showErrorBox(title, content);
        },
      }),
    );
    app.exit(result.ok ? 0 : 1);
  });
} else {
  require('./main.js');
}
