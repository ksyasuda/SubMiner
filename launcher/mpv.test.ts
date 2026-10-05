import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import net from 'node:net';
import { EventEmitter } from 'node:events';
import type { Args } from './types';
import { getAppControlSocketPath } from '../src/shared/app-control';
import { makeLauncherArgs } from './test-support/args.js';
import { withEnv } from './test-support/env.js';
import { withProcessExitIntercept } from './test-support/exit-intercept.js';
import {
  buildConfiguredMpvDefaultArgs,
  buildRuntimeExtraScriptOptParts,
  buildMpvBackendArgs,
  buildMpvEnv,
  cleanupPlaybackSession,
  detectBackend,
  launchAppBackgroundDetached,
  findAppBinary,
  launchAppCommandDetached,
  launchTexthookerOnly,
  parseMpvArgString,
  runAppCommandCaptureOutput,
  resolveLauncherRuntimePluginPath,
  resolveLauncherRuntimePluginPlan,
  stopOverlay,
  startOverlay,
  state,
  waitForUnixSocketReady,
} from './mpv';

function makeArgs(overrides: Partial<Args> = {}): Args {
  return makeLauncherArgs({ backend: 'x11', logLevel: 'error', ...overrides });
}

// ── shared helpers ───────────────────────────────────────────────────────────

interface TempCase {
  dir: string;
  socketPath: string;
}

/** Runs `run` in a fresh temp dir (with an mpv socket path inside it) and removes it afterwards. */
async function withTempCase<T>(run: (tempCase: TempCase) => T | Promise<T>): Promise<T> {
  const baseDir = path.join(process.cwd(), '.tmp', 'launcher-mpv-tests');
  fs.mkdirSync(baseDir, { recursive: true });
  const dir = fs.mkdtempSync(path.join(baseDir, 'case-'));
  try {
    return await run({ dir, socketPath: path.join(dir, 'mpv.sock') });
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

function withPlatform<T>(platform: NodeJS.Platform, callback: () => T): T {
  const originalDescriptor = Object.getOwnPropertyDescriptor(process, 'platform');
  Object.defineProperty(process, 'platform', {
    configurable: true,
    value: platform,
  });

  try {
    return callback();
  } finally {
    if (originalDescriptor) {
      Object.defineProperty(process, 'platform', originalDescriptor);
    }
  }
}

/** Writes an executable shell script `name` into `dir` and returns its path. */
function writeFakeApp(dir: string, lines: string[], name = 'fake-subminer.sh'): string {
  const appPath = path.join(dir, name);
  fs.writeFileSync(appPath, ['#!/bin/sh', ...lines, ''].join('\n'));
  fs.chmodSync(appPath, 0o755);
  return appPath;
}

/** Fake app that logs every invocation's argv and answers `--app-ping` with `pingExitCode`. */
function writeInvocationRecorder(dir: string, pingExitCode: number) {
  const invocationsPath = path.join(dir, 'app-invocations.log');
  const appPath = writeFakeApp(dir, [
    `printf '%s\\n' "$@" >> ${JSON.stringify(invocationsPath)}`,
    `if [ "$1" = "--app-ping" ]; then exit ${pingExitCode}; fi`,
    'exit 0',
  ]);
  return {
    appPath,
    readInvocations: () =>
      fs.existsSync(invocationsPath) ? fs.readFileSync(invocationsPath, 'utf8') : '',
  };
}

/** Makes every `net.createConnection` look like an mpv socket that connects after `delayMs`. */
async function withConnectableSockets<T>(delayMs: number, run: () => Promise<T>): Promise<T> {
  const originalCreateConnection = net.createConnection;
  net.createConnection = (() => {
    const socket = new EventEmitter() as net.Socket;
    socket.destroy = (() => socket) as net.Socket['destroy'];
    socket.setTimeout = (() => socket) as net.Socket['setTimeout'];
    setTimeout(() => socket.emit('connect'), delayMs);
    return socket;
  }) as typeof net.createConnection;
  try {
    return await run();
  } finally {
    net.createConnection = originalCreateConnection;
  }
}

function resetLauncherState(): void {
  state.overlayProc = null;
  state.overlayManagedByLauncher = false;
  state.appPath = '';
}

const listen = (server: net.Server, socketPath: string) =>
  new Promise<void>((resolve, reject) => {
    server.once('error', reject);
    server.listen(socketPath, resolve);
  });

const closeServer = (server: net.Server) =>
  new Promise<void>((resolve) => server.close(() => resolve()));

/** Fake app control server: records each request's argv and answers with `reply`. */
function createControlServer(reply: { ok: boolean; error?: string } = { ok: true }) {
  const receivedArgv: string[][] = [];
  const server = net.createServer((socket) => {
    let buffer = '';
    socket.on('data', (chunk) => {
      buffer += chunk.toString('utf8');
      const newlineIndex = buffer.indexOf('\n');
      if (newlineIndex < 0) return;
      const payload = JSON.parse(buffer.slice(0, newlineIndex)) as { argv?: unknown };
      if (Array.isArray(payload.argv)) {
        receivedArgv.push(
          payload.argv.filter((value): value is string => typeof value === 'string'),
        );
      }
      socket.end(JSON.stringify(reply) + '\n');
    });
  });
  return { server, receivedArgv };
}

// ── app command execution ────────────────────────────────────────────────────

test('runAppCommandCaptureOutput captures status and stdio', () => {
  const result = runAppCommandCaptureOutput(process.execPath, [
    '-e',
    'process.stdout.write("stdout-line"); process.stderr.write("stderr-line");',
  ]);

  assert.equal(result.status, 0);
  assert.equal(result.stdout, 'stdout-line');
  assert.equal(result.stderr, 'stderr-line');
  assert.equal(result.error, undefined);
});

test('runAppCommandCaptureOutput strips ELECTRON_RUN_AS_NODE from app child env', async () => {
  await withEnv({ ELECTRON_RUN_AS_NODE: '1' }, () => {
    const result = runAppCommandCaptureOutput(process.execPath, [
      '-e',
      'process.stdout.write(String(process.env.ELECTRON_RUN_AS_NODE ?? ""));',
    ]);

    assert.equal(result.status, 0);
    assert.equal(result.stdout, '');
  });
});

test(
  'runAppCommandCaptureOutput transports Linux AppImage args through environment',
  { skip: process.platform !== 'linux' },
  async () => {
    await withTempCase(({ dir }) => {
      const appPath = writeFakeApp(
        dir,
        [
          'printf "args:%s\\n" "$*"',
          'printf "argc:%s\\n" "$SUBMINER_APP_ARGC"',
          'printf "arg0:%s\\n" "$SUBMINER_APP_ARG_0"',
          'printf "arg1:%s\\n" "$SUBMINER_APP_ARG_1"',
        ],
        'SubMiner.AppImage',
      );

      const result = runAppCommandCaptureOutput(appPath, ['--app-ping', '--socket']);

      assert.equal(result.status, 0);
      assert.match(result.stdout, /^args:\n/m);
      assert.match(result.stdout, /^argc:2\n/m);
      assert.match(result.stdout, /^arg0:--app-ping\n/m);
      assert.match(result.stdout, /^arg1:--socket\n/m);
    });
  },
);

test('runAppCommandCaptureOutput runs Linux AppImage sync in Node-only mode', async () => {
  await withTempCase(({ dir }) => {
    const appPath = writeFakeApp(
      dir,
      [
        'printf "args:%s\\n" "$*"',
        'printf "electron-node:%s\\n" "$ELECTRON_RUN_AS_NODE"',
        'printf "argc:%s\\n" "$SUBMINER_APP_ARGC"',
        'printf "arg0:%s\\n" "$SUBMINER_APP_ARG_0"',
      ],
      'SubMiner.AppImage',
    );

    const result = withPlatform('linux', () =>
      runAppCommandCaptureOutput(appPath, ['--sync-cli', 'sync', '--snapshot', '/tmp/out']),
    );

    assert.equal(result.status, 0);
    assert.match(result.stdout, /^args:-e /m);
    assert.match(result.stdout, /^electron-node:1$/m);
    assert.match(result.stdout, /^argc:4$/m);
    assert.match(result.stdout, /^arg0:--sync-cli$/m);
  });
});

// ── args, env, and plugin resolution ─────────────────────────────────────────

test('parseMpvArgString preserves empty quoted tokens', () => {
  assert.deepEqual(parseMpvArgString('--title "" --force-media-title \'\' --pause'), [
    '--title',
    '',
    '--force-media-title',
    '',
    '--pause',
  ]);
});

test('detectBackend resolves windows on win32 auto mode', () => {
  withPlatform('win32', () => {
    assert.equal(detectBackend('auto'), 'windows');
  });
});

test('buildMpvEnv forces X11 by dropping Wayland hints when backend resolves to x11', () => {
  withPlatform('linux', () => {
    const env = buildMpvEnv(makeArgs({ backend: 'x11' }), {
      DISPLAY: ':1',
      WAYLAND_DISPLAY: 'wayland-0',
      XDG_SESSION_TYPE: 'wayland',
      HYPRLAND_INSTANCE_SIGNATURE: 'hypr',
      SWAYSOCK: '/tmp/sway.sock',
    });

    assert.equal(env.DISPLAY, ':1');
    assert.equal(env.WAYLAND_DISPLAY, undefined);
    assert.equal(env.XDG_SESSION_TYPE, 'x11');
    // The overlay inherits this env and needs hyprctl for its XWayland window.
    assert.equal(env.HYPRLAND_INSTANCE_SIGNATURE, 'hypr');
    assert.equal(env.SWAYSOCK, undefined);
  });
});

test('buildMpvEnv auto mode falls back to X11 when no supported Wayland tracker backend is detected', () => {
  withPlatform('linux', () => {
    const env = buildMpvEnv(makeArgs({ backend: 'auto' }), {
      DISPLAY: ':1',
      WAYLAND_DISPLAY: 'wayland-0',
      XDG_SESSION_TYPE: 'wayland',
      XDG_CURRENT_DESKTOP: 'KDE',
      XDG_SESSION_DESKTOP: 'plasma',
    });

    assert.equal(env.DISPLAY, ':1');
    assert.equal(env.WAYLAND_DISPLAY, undefined);
    assert.equal(env.XDG_SESSION_TYPE, 'x11');
  });
});

test('buildMpvEnv preserves native Wayland env for supported Hyprland and Sway auto backends', () => {
  withPlatform('linux', () => {
    const hyprEnv = buildMpvEnv(makeArgs({ backend: 'auto' }), {
      DISPLAY: ':1',
      WAYLAND_DISPLAY: 'wayland-0',
      XDG_SESSION_TYPE: 'wayland',
      HYPRLAND_INSTANCE_SIGNATURE: 'hypr',
    });
    assert.equal(hyprEnv.WAYLAND_DISPLAY, 'wayland-0');
    assert.equal(hyprEnv.XDG_SESSION_TYPE, 'wayland');

    const swayEnv = buildMpvEnv(makeArgs({ backend: 'auto' }), {
      DISPLAY: ':1',
      WAYLAND_DISPLAY: 'wayland-0',
      XDG_SESSION_TYPE: 'wayland',
      SWAYSOCK: '/tmp/sway.sock',
    });
    assert.equal(swayEnv.WAYLAND_DISPLAY, 'wayland-0');
    assert.equal(swayEnv.XDG_SESSION_TYPE, 'wayland');
  });
});

test('buildMpvBackendArgs pins the X11 window context when backend resolves to x11', () => {
  withPlatform('linux', () => {
    assert.deepEqual(
      buildMpvBackendArgs(makeArgs({ backend: 'x11' }), {
        DISPLAY: ':1',
        WAYLAND_DISPLAY: 'wayland-0',
        XDG_SESSION_TYPE: 'wayland',
      }),
      ['--gpu-context=x11vk,x11egl,x11'],
    );
  });
});

test('buildMpvBackendArgs pins the same X11 window context for unsupported Wayland auto fallback', () => {
  withPlatform('linux', () => {
    assert.deepEqual(
      buildMpvBackendArgs(makeArgs({ backend: 'auto' }), {
        DISPLAY: ':1',
        WAYLAND_DISPLAY: 'wayland-0',
        XDG_SESSION_TYPE: 'wayland',
        XDG_CURRENT_DESKTOP: 'KDE',
        XDG_SESSION_DESKTOP: 'plasma',
      }),
      ['--gpu-context=x11vk,x11egl,x11'],
    );
  });
});

test('buildMpvBackendArgs keeps supported Hyprland and Sway auto backends unchanged', () => {
  withPlatform('linux', () => {
    assert.deepEqual(
      buildMpvBackendArgs(makeArgs({ backend: 'auto' }), {
        DISPLAY: ':1',
        WAYLAND_DISPLAY: 'wayland-0',
        XDG_SESSION_TYPE: 'wayland',
        HYPRLAND_INSTANCE_SIGNATURE: 'hypr',
      }),
      [],
    );
    assert.deepEqual(
      buildMpvBackendArgs(makeArgs({ backend: 'auto' }), {
        DISPLAY: ':1',
        WAYLAND_DISPLAY: 'wayland-0',
        XDG_SESSION_TYPE: 'wayland',
        SWAYSOCK: '/tmp/sway.sock',
      }),
      [],
    );
  });
});

test('buildConfiguredMpvDefaultArgs appends maximized launch mode to configured defaults', () => {
  withPlatform('linux', () => {
    const args = buildConfiguredMpvDefaultArgs(makeArgs({ launchMode: 'maximized' }), {
      DISPLAY: ':1',
      XDG_SESSION_TYPE: 'x11',
    });

    assert.equal(args.at(-1), '--window-maximized=yes');
  });
});

test('buildConfiguredMpvDefaultArgs passes configured mpv profile before SubMiner defaults', () => {
  withPlatform('linux', () => {
    assert.deepEqual(
      buildConfiguredMpvDefaultArgs(makeArgs({ profile: 'anime,hdr' }), {
        DISPLAY: ':1',
        XDG_SESSION_TYPE: 'x11',
      }).slice(0, 2),
      ['--profile=anime,hdr', '--sub-auto=fuzzy'],
    );
  });
});

test('buildConfiguredMpvDefaultArgs disables macOS menu shortcuts so SubMiner bindings reach mpv', () => {
  withPlatform('darwin', () => {
    assert.equal(
      buildConfiguredMpvDefaultArgs(makeArgs()).includes('--macos-menu-shortcuts=no'),
      true,
    );
  });
});

test('resolveLauncherRuntimePluginPath finds bundled plugin from explicit environment path', () => {
  const pluginDir = '/opt/SubMiner/plugin/subminer';
  assert.equal(
    resolveLauncherRuntimePluginPath({
      appPath: '/opt/SubMiner/SubMiner.AppImage',
      env: { SUBMINER_MPV_PLUGIN_PATH: pluginDir },
      existsSync: (candidate) => candidate === path.join(pluginDir, 'main.lua'),
    }),
    path.join(pluginDir, 'main.lua'),
  );
});

test('resolveLauncherRuntimePluginPath finds Linux app-support plugin assets', () => {
  const homeDir = '/home/tester';
  const expected = path.join(
    homeDir,
    '.local',
    'share',
    'SubMiner',
    'plugin',
    'subminer',
    'main.lua',
  );

  assert.equal(
    resolveLauncherRuntimePluginPath({
      appPath: '/home/tester/.local/bin/SubMiner.AppImage',
      scriptPath: '/home/tester/.local/bin/subminer',
      platform: 'linux',
      homeDir,
      env: {},
      existsSync: (candidate) => candidate === expected,
    }),
    expected,
  );
});

test('resolveLauncherRuntimePluginPlan injects bundled plugin when no installed plugin exists', () => {
  const plan = resolveLauncherRuntimePluginPlan({
    runtimePluginPath: '/opt/SubMiner/plugin/subminer/main.lua',
    platform: 'linux',
    homeDir: '/home/tester',
    existsSync: () => false,
  });

  assert.equal(plan.scriptPath, '/opt/SubMiner/plugin/subminer/main.lua');
  assert.equal(plan.installedPlugin.installed, false);
  assert.equal(plan.warningMessage, null);
  assert.equal(plan.errorMessage, null);
});

test('resolveLauncherRuntimePluginPlan uses installed plugin instead of bundled injection', () => {
  const installedPath = '/home/tester/.config/mpv/scripts/subminer/main.lua';
  const versionPath = '/home/tester/.config/mpv/scripts/subminer/version.lua';
  const existing = new Set([installedPath, versionPath]);
  const plan = resolveLauncherRuntimePluginPlan({
    runtimePluginPath: '/opt/SubMiner/plugin/subminer/main.lua',
    platform: 'linux',
    homeDir: '/home/tester',
    existsSync: (candidate) => existing.has(candidate),
    readFileSync: () => 'return { version = "0.12.0" }',
  });

  assert.equal(plan.scriptPath, null);
  assert.equal(plan.installedPlugin.path, installedPath);
  assert.equal(plan.installedPlugin.version, '0.12.0');
  assert.match(plan.warningMessage ?? '', /This mpv session will use the installed plugin/);
  assert.equal(plan.errorMessage, null);
});

test('resolveLauncherRuntimePluginPlan reports missing bundled plugin when no installed plugin exists', () => {
  const plan = resolveLauncherRuntimePluginPlan({
    runtimePluginPath: null,
    platform: 'linux',
    homeDir: '/home/tester',
    existsSync: () => false,
  });

  assert.equal(plan.scriptPath, null);
  assert.equal(plan.installedPlugin.installed, false);
  assert.match(plan.errorMessage ?? '', /Packaged mpv plugin assets were not found/);
});

test('buildRuntimeExtraScriptOptParts marks launcher-owned startup pause gate', () => {
  assert.deepEqual(
    buildRuntimeExtraScriptOptParts('/tmp/video.mkv', 'file', {
      startPaused: true,
      runtimePluginConfig: {
        socketPath: '/tmp/subminer.sock',
        binaryPath: '',
        backend: 'auto',
        autoStart: true,
        autoStartVisibleOverlay: true,
        autoStartPauseUntilReady: true,
        osdMessages: false,
        texthookerEnabled: false,
      },
    }),
    ['subminer-auto_start_pause_until_ready_owns_initial_pause=yes'],
  );
});

test('launchTexthookerOnly exits non-zero when app binary cannot be spawned', () => {
  const error = withProcessExitIntercept(() => {
    launchTexthookerOnly('/definitely-missing-subminer-binary', makeArgs());
  });

  assert.equal(error.code, 1);
  assert.match(error.stderr, /Failed to launch texthooker mode/);
});

test('launchTexthookerOnly forwards browser-open request to app command', async () => {
  await withTempCase(({ dir }) => {
    const argsPath = path.join(dir, 'args.txt');
    const openedUrls: string[] = [];
    const appPath = writeFakeApp(dir, [`printf '%s\\n' "$@" > "${argsPath}"`, 'exit 0']);

    const error = withProcessExitIntercept(() => {
      launchTexthookerOnly(appPath, makeArgs({ logLevel: 'info', texthookerOpenBrowser: true }), {
        openBrowser: (url) => openedUrls.push(url),
      });
    });

    assert.equal(error.code, 0);
    assert.deepEqual(fs.readFileSync(argsPath, 'utf8').trim().split('\n'), [
      '--texthooker',
      '--open-browser',
    ]);
    assert.deepEqual(openedUrls, ['http://127.0.0.1:5174']);
  });
});

test('launchAppCommandDetached handles child process spawn errors', async () => {
  let uncaughtError: Error | null = null;
  const onUncaughtException = (error: Error) => {
    uncaughtError = error;
  };
  process.once('uncaughtException', onUncaughtException);
  try {
    launchAppCommandDetached(
      '/definitely-missing-subminer-binary',
      [],
      makeArgs({ logLevel: 'warn' }).logLevel,
      'test',
    );
    await new Promise((resolve) => setTimeout(resolve, 50));
    assert.equal(uncaughtError, null);
  } finally {
    process.removeListener('uncaughtException', onUncaughtException);
  }
});

test('launchAppBackgroundDetached starts background child directly', async () => {
  await withTempCase(async ({ dir }) => {
    const argsPath = path.join(dir, 'args.txt');
    const envPath = path.join(dir, 'env.txt');
    const appPath = writeFakeApp(dir, [
      `printf '%s\\n' "$@" > ${JSON.stringify(argsPath)}`,
      `printf '%s\\n' "$SUBMINER_BACKGROUND_CHILD" > ${JSON.stringify(envPath)}`,
    ]);

    launchAppBackgroundDetached(appPath, 'info');

    const deadline = Date.now() + 1000;
    while ((!fs.existsSync(argsPath) || !fs.existsSync(envPath)) && Date.now() < deadline) {
      await new Promise((resolve) => setTimeout(resolve, 20));
    }

    assert.equal(fs.readFileSync(argsPath, 'utf8').trim(), '--start\n--background');
    assert.equal(fs.readFileSync(envPath, 'utf8').trim(), '1');
  });
});

test('stopOverlay logs a warning when stop command cannot be spawned', () => {
  const originalWrite = process.stdout.write;
  const writes: string[] = [];
  const overlayProc = {
    killed: false,
    kill: () => true,
  } as unknown as NonNullable<typeof state.overlayProc>;

  try {
    process.stdout.write = ((chunk: string | Uint8Array) => {
      writes.push(Buffer.isBuffer(chunk) ? chunk.toString('utf8') : String(chunk));
      return true;
    }) as typeof process.stdout.write;
    state.stopRequested = false;
    state.overlayManagedByLauncher = true;
    state.appPath = '/definitely-missing-subminer-binary';
    state.overlayProc = overlayProc;

    stopOverlay(makeArgs({ logLevel: 'warn' }));

    assert.ok(writes.some((text) => text.includes('Failed to stop SubMiner overlay')));
  } finally {
    process.stdout.write = originalWrite;
    state.stopRequested = false;
    resetLauncherState();
  }
});

// ── socket readiness ─────────────────────────────────────────────────────────

test('waitForUnixSocketReady returns false when socket never appears', async () => {
  await withTempCase(async ({ socketPath }) => {
    assert.equal(await waitForUnixSocketReady(socketPath, 120), false);
  });
});

test('waitForUnixSocketReady returns false when path exists but is not socket', async () => {
  await withTempCase(async ({ socketPath }) => {
    fs.writeFileSync(socketPath, 'not-a-socket');
    assert.equal(await waitForUnixSocketReady(socketPath, 200), false);
  });
});

test('waitForUnixSocketReady returns true when socket becomes connectable before timeout', async () => {
  await withTempCase(async ({ socketPath }) => {
    fs.writeFileSync(socketPath, '');
    const ready = await withConnectableSockets(25, () => waitForUnixSocketReady(socketPath, 400));
    assert.equal(ready, true);
  });
});

// ── startOverlay ownership and attach ────────────────────────────────────────

/**
 * Runs `startOverlay` against a recording fake app and a connectable mpv socket, and returns
 * what the app saw plus the launcher ownership state, captured before state is reset.
 */
async function runStartOverlay(options: { pingExitCode: number; alreadyManaged?: boolean }) {
  return withTempCase(async ({ dir, socketPath }) => {
    const app = writeInvocationRecorder(dir, options.pingExitCode);
    fs.writeFileSync(socketPath, '');
    if (options.alreadyManaged) {
      state.appPath = app.appPath;
      state.overlayManagedByLauncher = true;
    }
    try {
      await withConnectableSockets(10, () => startOverlay(app.appPath, makeArgs(), socketPath));
      return {
        invocations: app.readInvocations(),
        managedByLauncher: state.overlayManagedByLauncher,
        ownedAppPath: state.appPath,
        appPath: app.appPath,
      };
    } finally {
      resetLauncherState();
    }
  });
}

test('startOverlay captures app stdout and stderr into app log', async () => {
  await withTempCase(async ({ dir, socketPath }) => {
    const appLogPath = path.join(dir, 'app.log');
    const appPath = writeFakeApp(dir, [
      'printf "hello from stdout\\n"',
      'printf "hello from stderr\\n" >&2',
      'exit 0',
    ]);
    fs.writeFileSync(socketPath, '');

    try {
      await withEnv({ SUBMINER_APP_LOG: appLogPath }, () =>
        withConnectableSockets(10, () => startOverlay(appPath, makeArgs(), socketPath)),
      );

      const logText = fs.readFileSync(appLogPath, 'utf8');
      assert.match(logText, /\[STDOUT\] hello from stdout/);
      assert.match(logText, /\[STDERR\] hello from stderr/);
    } finally {
      resetLauncherState();
    }
  });
});

test('startOverlay starts launcher-owned playback in background managed mode', async () => {
  const result = await runStartOverlay({ pingExitCode: 1 });

  assert.match(result.invocations, /--background/);
  assert.match(result.invocations, /--managed-playback/);
  assert.equal(result.managedByLauncher, true);
  assert.equal(result.ownedAppPath, result.appPath);
});

test('startOverlay borrows an already-running background app instead of owning its lifecycle', async () => {
  const result = await runStartOverlay({ pingExitCode: 0 });

  assert.match(result.invocations, /--app-ping/);
  assert.match(result.invocations, /--start/);
  assert.doesNotMatch(result.invocations, /--background/);
  assert.equal(result.managedByLauncher, false);
  assert.equal(result.ownedAppPath, '');
});

test('startOverlay keeps lifecycle ownership for its already-managed app', async () => {
  const result = await runStartOverlay({ pingExitCode: 0, alreadyManaged: true });

  assert.equal(result.managedByLauncher, true);
  assert.equal(result.ownedAppPath, result.appPath);
});

const controlSocketCases = [
  {
    name: 'startOverlay attaches through the running app control socket without spawning another app command',
    // The control socket is found through SUBMINER_APP_CONTROL_SOCKET.
    useConfigDir: false,
  },
  {
    name: 'startOverlay uses caller config dir for app control socket discovery',
    useConfigDir: true,
  },
];

for (const c of controlSocketCases) {
  test(c.name, { skip: process.platform === 'win32' }, async () => {
    await withTempCase(async ({ dir, socketPath }) => {
      const configDir = path.join(dir, 'launcher-config');
      fs.mkdirSync(configDir, { recursive: true });
      const controlSocketPath = c.useConfigDir
        ? getAppControlSocketPath({ configDir, platform: 'linux' })
        : path.join(dir, 'control.sock');
      const app = writeInvocationRecorder(dir, 0);
      const mpvServer = net.createServer((socket) => socket.end());
      const control = createControlServer();

      try {
        await listen(mpvServer, socketPath);
        await listen(control.server, controlSocketPath);

        await withEnv(
          { SUBMINER_APP_CONTROL_SOCKET: c.useConfigDir ? undefined : controlSocketPath },
          () =>
            c.useConfigDir
              ? startOverlay(app.appPath, makeArgs(), socketPath, [], configDir)
              : startOverlay(app.appPath, makeArgs(), socketPath),
        );

        assert.equal(app.readInvocations(), '');
        assert.equal(control.receivedArgv.length, 1);
        assert.deepEqual(control.receivedArgv[0]?.slice(0, 6), [
          '--start',
          '--managed-playback',
          '--backend',
          'x11',
          '--socket',
          socketPath,
        ]);
        assert.equal(state.overlayManagedByLauncher, false);
        assert.equal(state.appPath, '');
      } finally {
        await closeServer(mpvServer);
        await closeServer(control.server);
        resetLauncherState();
      }
    });
  });
}

test(
  'startOverlay falls back to legacy app startup when control command fails',
  { skip: process.platform === 'win32' },
  async () => {
    await withTempCase(async ({ dir, socketPath }) => {
      const controlSocketPath = path.join(dir, 'control.sock');
      const app = writeInvocationRecorder(dir, 0);
      const control = createControlServer({ ok: false, error: 'boom' });

      try {
        await listen(control.server, controlSocketPath);

        await withEnv({ SUBMINER_APP_CONTROL_SOCKET: controlSocketPath }, () =>
          startOverlay(app.appPath, makeArgs(), socketPath),
        );

        const invocations = app.readInvocations();
        assert.match(invocations, /--app-ping/);
        assert.match(invocations, /--start/);
      } finally {
        await closeServer(control.server);
        resetLauncherState();
      }
    });
  },
);

test('cleanupPlaybackSession stops launcher-managed overlay app and mpv-owned children', async () => {
  await withTempCase(async ({ dir }) => {
    const app = writeInvocationRecorder(dir, 0);
    const calls: string[] = [];
    const overlayProc = {
      killed: false,
      kill: () => {
        calls.push('overlay-kill');
        return true;
      },
    } as unknown as NonNullable<typeof state.overlayProc>;
    const mpvProc = {
      killed: false,
      kill: () => {
        calls.push('mpv-kill');
        return true;
      },
    } as unknown as NonNullable<typeof state.mpvProc>;

    state.stopRequested = false;
    state.appPath = app.appPath;
    state.overlayManagedByLauncher = true;
    state.overlayProc = overlayProc;
    state.mpvProc = mpvProc;

    try {
      await cleanupPlaybackSession(makeArgs());

      assert.deepEqual(calls.sort(), ['mpv-kill', 'overlay-kill']);
      assert.match(app.readInvocations(), /--stop/);
    } finally {
      state.mpvProc = null;
      state.stopRequested = false;
      resetLauncherState();
    }
  });
});

// ── findAppBinary ────────────────────────────────────────────────────────────

interface FindAppBinaryScenario {
  platform: NodeJS.Platform;
  home: string;
  /** Extra env; SUBMINER_APPIMAGE_PATH / SUBMINER_BINARY_PATH are always cleared first. */
  env?: Record<string, string | undefined>;
  /** When set, `fs.accessSync` accepts only these paths. */
  executables?: string[];
  /** When set, `fs.existsSync` / `fs.statSync` see only these paths. */
  existing?: string[];
  directories?: string[];
  /** When set, replaces `fs.realpathSync`. */
  realpath?: (filePath: string) => string;
}

/**
 * Runs `run` with platform, home dir, env, and (optionally) fs probes faked for
 * `findAppBinary`, which reads them from process globals. Everything is restored afterwards.
 */
async function withFindAppBinaryScenario(
  scenario: FindAppBinaryScenario,
  run: (pathModule: typeof path) => void,
): Promise<void> {
  const restores: Array<() => void> = [];
  const stub = (target: object, key: string, value: unknown): void => {
    const original = Object.getOwnPropertyDescriptor(target, key);
    Object.defineProperty(target, key, { value, configurable: true, writable: true });
    restores.push(() => {
      if (original) Object.defineProperty(target, key, original);
      else Reflect.deleteProperty(target, key);
    });
  };
  const missing = (code: string, filePath: string) =>
    Object.assign(new Error(`${code}: ${filePath}`), { code });

  try {
    stub(process, 'platform', scenario.platform);
    stub(os, 'homedir', () => scenario.home);
    if (scenario.executables) {
      const executables = new Set(scenario.executables);
      stub(fs, 'accessSync', (filePath: string): void => {
        if (!executables.has(filePath)) throw missing('EACCES', filePath);
      });
    }
    if (scenario.existing || scenario.directories) {
      const directories = new Set(scenario.directories ?? []);
      const files = new Set(scenario.existing ?? []);
      stub(
        fs,
        'existsSync',
        (filePath: string) => files.has(filePath) || directories.has(filePath),
      );
      stub(fs, 'statSync', (filePath: string) => {
        if (directories.has(filePath)) return { isDirectory: () => true };
        if (files.has(filePath)) return { isDirectory: () => false };
        throw missing('ENOENT', filePath);
      });
    }
    if (scenario.realpath) stub(fs, 'realpathSync', scenario.realpath);

    await withEnv(
      { SUBMINER_APPIMAGE_PATH: undefined, SUBMINER_BINARY_PATH: undefined, ...scenario.env },
      () => run(scenario.platform === 'win32' ? (path.win32 as typeof path) : path),
    );
  } finally {
    for (const restore of restores.reverse()) restore();
  }
}

function makeExecutable(filePath: string): void {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, '#!/bin/sh\nexit 0\n');
  fs.chmodSync(filePath, 0o755);
}

test(
  'findAppBinary resolves ~/.local/bin/SubMiner.AppImage when it exists',
  { concurrency: false },
  async () => {
    const baseDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-test-home-'));
    try {
      const appImage = path.join(baseDir, '.local/bin/SubMiner.AppImage');
      makeExecutable(appImage);

      await withFindAppBinaryScenario({ platform: 'linux', home: baseDir }, (pathModule) => {
        assert.equal(findAppBinary('/some/other/path/subminer', pathModule), appImage);
      });
    } finally {
      fs.rmSync(baseDir, { recursive: true, force: true });
    }
  },
);

test(
  'findAppBinary resolves /opt/SubMiner/SubMiner.AppImage when ~/.local/bin candidate does not exist',
  { concurrency: false },
  async () => {
    await withFindAppBinaryScenario(
      {
        platform: 'linux',
        home: '/home/tester',
        executables: ['/opt/SubMiner/SubMiner.AppImage'],
      },
      (pathModule) => {
        assert.equal(
          findAppBinary('/some/other/path/subminer', pathModule),
          '/opt/SubMiner/SubMiner.AppImage',
        );
      },
    );
  },
);

test(
  'findAppBinary finds subminer on PATH when AppImage candidates do not exist',
  { concurrency: false },
  async () => {
    const binDir = '/home/tester/bin';
    const wrapperPath = path.join(binDir, 'subminer');

    await withFindAppBinaryScenario(
      {
        platform: 'linux',
        home: '/home/tester',
        env: { PATH: binDir },
        executables: [wrapperPath],
      },
      (pathModule) => {
        // selfPath must differ from wrapperPath so the self-check does not exclude it
        assert.equal(findAppBinary('/home/tester/launcher/subminer', pathModule), wrapperPath);
      },
    );
  },
);

test(
  'findAppBinary excludes PATH matches that canonicalize to the launcher path',
  { concurrency: false },
  async () => {
    const binDir = '/home/tester/bin';
    const wrapperPath = path.join(binDir, 'subminer');
    const canonicalPath = '/home/tester/launch/subminer';

    await withFindAppBinaryScenario(
      {
        platform: 'linux',
        home: '/home/tester',
        env: { PATH: binDir },
        executables: [wrapperPath],
        realpath: (filePath) =>
          filePath === canonicalPath || filePath === wrapperPath ? canonicalPath : filePath,
      },
      (pathModule) => {
        assert.equal(findAppBinary(canonicalPath, pathModule), null);
      },
    );
  },
);

const windowsHome = 'C:\\Users\\tester';
const windowsLauncherPath = path.win32.join(windowsHome, 'launcher', 'SubMiner.exe');

test(
  'findAppBinary resolves Windows install paths when present',
  { concurrency: false },
  async () => {
    const localAppData = path.win32.join(windowsHome, 'AppData', 'Local');
    const appExe = path.win32.join(localAppData, 'Programs', 'SubMiner', 'SubMiner.exe');

    await withFindAppBinaryScenario(
      {
        platform: 'win32',
        home: windowsHome,
        env: { LOCALAPPDATA: localAppData },
        executables: [appExe],
      },
      (pathModule) => {
        assert.equal(findAppBinary(windowsLauncherPath, pathModule), appExe);
      },
    );
  },
);

test('findAppBinary resolves SubMiner.exe on PATH on Windows', { concurrency: false }, async () => {
  const binDir = path.win32.join(windowsHome, 'bin');
  const wrapperPath = path.win32.join(binDir, 'SubMiner.exe');

  await withFindAppBinaryScenario(
    { platform: 'win32', home: windowsHome, env: { PATH: binDir }, executables: [wrapperPath] },
    (pathModule) => {
      assert.equal(findAppBinary(windowsLauncherPath, pathModule), wrapperPath);
    },
  );
});

test(
  'findAppBinary resolves a Windows install directory to SubMiner.exe',
  { concurrency: false },
  async () => {
    const installDir = path.win32.join(windowsHome, 'Programs', 'SubMiner');
    const appExe = path.win32.join(installDir, 'SubMiner.exe');

    await withFindAppBinaryScenario(
      {
        platform: 'win32',
        home: windowsHome,
        env: { SUBMINER_BINARY_PATH: installDir },
        existing: [appExe],
        directories: [installDir],
        executables: [appExe],
      },
      (pathModule) => {
        assert.equal(findAppBinary(windowsLauncherPath, pathModule), appExe);
      },
    );
  },
);
