import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { stopProcess, type EnvDelta } from './process';

export const E2E_SCREEN = { width: 1280, height: 720 } as const;

export type E2eDisplay = {
  kind: 'xvfb' | 'host';
  env: EnvDelta;
  stop: () => Promise<void>;
};

// Linux gets a private Xvfb server so no window ever reaches the real desktop.
// macOS and Windows have no offscreen display server, so windows open in the
// current GUI session there. SUBMINER_E2E_DISPLAY=host forces that on Linux too.
export async function startDisplay(): Promise<E2eDisplay> {
  const requested = process.env.SUBMINER_E2E_DISPLAY?.trim().toLowerCase();
  if (requested && requested !== 'host' && requested !== 'xvfb') {
    throw new Error(`SUBMINER_E2E_DISPLAY must be "host" or "xvfb", got "${requested}"`);
  }
  const kind = requested ?? (process.platform === 'linux' ? 'xvfb' : 'host');
  if (kind === 'host') {
    return { kind, env: { set: {}, unset: [] }, stop: async () => {} };
  }
  return startXvfb();
}

async function startXvfb(): Promise<E2eDisplay> {
  // -displayfd makes Xvfb pick a free display number and write it to fd 3 once
  // it accepts connections, so there is no port guessing and no startup sleep.
  const child = spawn(
    'Xvfb',
    [
      '-displayfd',
      '3',
      '-screen',
      '0',
      `${E2E_SCREEN.width}x${E2E_SCREEN.height}x24`,
      '-nolisten',
      'tcp',
    ],
    { stdio: ['ignore', 'ignore', 'ignore', 'pipe'] },
  );
  const displayFd = child.stdio[3];
  if (!displayFd) throw new Error('Xvfb did not expose its display fd');

  const displayNumber = await Promise.race([
    once(displayFd, 'data').then(([chunk]) => String(chunk).trim()),
    once(child, 'error').then(([error]) => {
      throw new Error(
        `Could not start Xvfb (${String(error)}). Install it or set SUBMINER_E2E_DISPLAY=host.`,
      );
    }),
    once(child, 'exit').then(([code]) => {
      throw new Error(`Xvfb exited during startup with code ${String(code)}`);
    }),
  ]);

  return {
    kind: 'xvfb',
    env: {
      set: { DISPLAY: `:${displayNumber}`, XDG_SESSION_TYPE: 'x11' },
      // Without these the app and mpv would pick the real Wayland session, or
      // take the compositor-specific code paths of the desktop running the tests.
      unset: [
        'WAYLAND_DISPLAY',
        'HYPRLAND_INSTANCE_SIGNATURE',
        'SWAYSOCK',
        'XDG_CURRENT_DESKTOP',
        'XDG_SESSION_DESKTOP',
      ],
    },
    stop: () => stopProcess(child),
  };
}
