import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import net from 'node:net';
import { EventEmitter } from 'node:events';
import type { Args } from './types';
import { getAppControlSocketPath } from '../src/shared/app-control';
import { withProcessExitIntercept } from './test-support/exit-intercept.js';
import {
  buildConfiguredMpvDefaultArgs,
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
  shouldResolveAniSkipMetadata,
  stopOverlay,
  startOverlay,
  state,
  waitForUnixSocketReady,
} from './mpv';
import * as mpvModule from './mpv';

function createTempSocketPath(): { dir: string; socketPath: string } {
  const baseDir = path.join(process.cwd(), '.tmp', 'launcher-mpv-tests');
  fs.mkdirSync(baseDir, { recursive: true });
  const dir = fs.mkdtempSync(path.join(baseDir, 'case-'));
  return { dir, socketPath: path.join(dir, 'mpv.sock') };
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

test('mpv module exposes only canonical socket readiness helper', () => {
  assert.equal('waitForSocket' in mpvModule, false);
});

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

test('runAppCommandCaptureOutput strips ELECTRON_RUN_AS_NODE from app child env', () => {
  const original = process.env.ELECTRON_RUN_AS_NODE;
  try {
    process.env.ELECTRON_RUN_AS_NODE = '1';
    const result = runAppCommandCaptureOutput(process.execPath, [
      '-e',
      'process.stdout.write(String(process.env.ELECTRON_RUN_AS_NODE ?? ""));',
    ]);

    assert.equal(result.status, 0);
    assert.equal(result.stdout, '');
  } finally {
    if (original === undefined) {
      delete process.env.ELECTRON_RUN_AS_NODE;
    } else {
      process.env.ELECTRON_RUN_AS_NODE = original;
    }
  }
});

test('runAppCommandCaptureOutput transports Linux AppImage args through environment', () => {
  if (process.platform !== 'linux') return;
  const { dir } = createTempSocketPath();
  const appPath = path.join(dir, 'SubMiner.AppImage');
  fs.writeFileSync(
    appPath,
    [
      '#!/bin/sh',
      'printf "args:%s\\n" "$*"',
      'printf "argc:%s\\n" "$SUBMINER_APP_ARGC"',
      'printf "arg0:%s\\n" "$SUBMINER_APP_ARG_0"',
      'printf "arg1:%s\\n" "$SUBMINER_APP_ARG_1"',
      '',
    ].join('\n'),
  );
  fs.chmodSync(appPath, 0o755);

  try {
    const result = runAppCommandCaptureOutput(appPath, ['--app-ping', '--socket']);

    assert.equal(result.status, 0);
    assert.match(result.stdout, /^args:\n/m);
    assert.match(result.stdout, /^argc:2\n/m);
    assert.match(result.stdout, /^arg0:--app-ping\n/m);
    assert.match(result.stdout, /^arg1:--socket\n/m);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

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
    assert.equal(env.HYPRLAND_INSTANCE_SIGNATURE, undefined);
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

test('buildMpvBackendArgs forces an explicit X11 renderer stack when backend resolves to x11', () => {
  withPlatform('linux', () => {
    assert.deepEqual(
      buildMpvBackendArgs(makeArgs({ backend: 'x11' }), {
        DISPLAY: ':1',
        WAYLAND_DISPLAY: 'wayland-0',
        XDG_SESSION_TYPE: 'wayland',
      }),
      ['--vo=gpu', '--gpu-api=opengl', '--gpu-context=x11egl,x11'],
    );
  });
});

test('buildMpvBackendArgs forces the same X11 renderer stack for unsupported Wayland auto fallback', () => {
  withPlatform('linux', () => {
    assert.deepEqual(
      buildMpvBackendArgs(makeArgs({ backend: 'auto' }), {
        DISPLAY: ':1',
        WAYLAND_DISPLAY: 'wayland-0',
        XDG_SESSION_TYPE: 'wayland',
        XDG_CURRENT_DESKTOP: 'KDE',
        XDG_SESSION_DESKTOP: 'plasma',
      }),
      ['--vo=gpu', '--gpu-api=opengl', '--gpu-context=x11egl,x11'],
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
    assert.deepEqual(
      buildConfiguredMpvDefaultArgs(makeArgs({ launchMode: 'maximized' }), {
        DISPLAY: ':1',
        WAYLAND_DISPLAY: 'wayland-0',
        XDG_SESSION_TYPE: 'wayland',
        XDG_CURRENT_DESKTOP: 'KDE',
        XDG_SESSION_DESKTOP: 'plasma',
      }),
      [
        '--sub-auto=fuzzy',
        '--sub-file-paths=.;subs;subtitles',
        '--sid=auto',
        '--secondary-sid=auto',
        '--sub-visibility=no',
        '--secondary-sub-visibility=no',
        '--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
        '--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
        '--vo=gpu',
        '--gpu-api=opengl',
        '--gpu-context=x11egl,x11',
        '--window-maximized=yes',
      ],
    );
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

test('launchTexthookerOnly exits non-zero when app binary cannot be spawned', () => {
  const error = withProcessExitIntercept(() => {
    launchTexthookerOnly('/definitely-missing-subminer-binary', makeArgs());
  });

  assert.equal(error.code, 1);
  assert.match(error.stderr, /Failed to launch texthooker mode/);
});

test('launchTexthookerOnly forwards browser-open request to app command', () => {
  const { dir } = createTempSocketPath();
  const appPath = path.join(dir, 'fake-subminer.sh');
  const argsPath = path.join(dir, 'args.txt');
  const openedUrls: string[] = [];
  fs.writeFileSync(appPath, `#!/bin/sh\nprintf '%s\\n' "$@" > "${argsPath}"\nexit 0\n`);
  fs.chmodSync(appPath, 0o755);

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
  fs.rmSync(dir, { recursive: true, force: true });
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
  const { dir } = createTempSocketPath();
  const appPath = path.join(dir, 'fake-subminer.sh');
  const argsPath = path.join(dir, 'args.txt');
  const envPath = path.join(dir, 'env.txt');
  fs.writeFileSync(
    appPath,
    [
      '#!/bin/sh',
      `printf '%s\\n' "$@" > ${JSON.stringify(argsPath)}`,
      `printf '%s\\n' "$SUBMINER_BACKGROUND_CHILD" > ${JSON.stringify(envPath)}`,
      '',
    ].join('\n'),
  );
  fs.chmodSync(appPath, 0o755);

  launchAppBackgroundDetached(appPath, 'info');

  const deadline = Date.now() + 1000;
  while ((!fs.existsSync(argsPath) || !fs.existsSync(envPath)) && Date.now() < deadline) {
    await new Promise((resolve) => setTimeout(resolve, 20));
  }

  assert.equal(fs.readFileSync(argsPath, 'utf8').trim(), '--start\n--background');
  assert.equal(fs.readFileSync(envPath, 'utf8').trim(), '1');
  fs.rmSync(dir, { recursive: true, force: true });
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
    state.overlayManagedByLauncher = false;
    state.appPath = '';
    state.overlayProc = null;
  }
});

test('waitForUnixSocketReady returns false when socket never appears', async () => {
  const { dir, socketPath } = createTempSocketPath();
  try {
    const ready = await waitForUnixSocketReady(socketPath, 120);
    assert.equal(ready, false);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('waitForUnixSocketReady returns false when path exists but is not socket', async () => {
  const { dir, socketPath } = createTempSocketPath();
  try {
    fs.writeFileSync(socketPath, 'not-a-socket');
    const ready = await waitForUnixSocketReady(socketPath, 200);
    assert.equal(ready, false);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('waitForUnixSocketReady returns true when socket becomes connectable before timeout', async () => {
  const { dir, socketPath } = createTempSocketPath();
  fs.writeFileSync(socketPath, '');
  const originalCreateConnection = net.createConnection;
  try {
    net.createConnection = (() => {
      const socket = new EventEmitter() as net.Socket;
      socket.destroy = (() => socket) as net.Socket['destroy'];
      socket.setTimeout = (() => socket) as net.Socket['setTimeout'];
      setTimeout(() => socket.emit('connect'), 25);
      return socket;
    }) as typeof net.createConnection;

    const ready = await waitForUnixSocketReady(socketPath, 400);
    assert.equal(ready, true);
  } finally {
    net.createConnection = originalCreateConnection;
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('shouldResolveAniSkipMetadata skips URL and YouTube-preloaded playback', () => {
  assert.equal(shouldResolveAniSkipMetadata('/media/show.mkv', 'file'), true);
  assert.equal(
    shouldResolveAniSkipMetadata('https://www.youtube.com/watch?v=test123', 'url'),
    false,
  );
  assert.equal(
    shouldResolveAniSkipMetadata('/tmp/video123.webm', 'file', {
      primaryPath: '/tmp/video123.ja.srt',
    }),
    false,
  );
});

function makeArgs(overrides: Partial<Args> = {}): Args {
  return {
    backend: 'x11',
    directory: '.',
    recursive: false,
    profile: '',
    startOverlay: false,
    whisperBin: '',
    whisperModel: '',
    whisperVadModel: '',
    whisperThreads: 4,
    youtubeSubgenOutDir: '',
    youtubeSubgenAudioFormat: 'wav',
    youtubeSubgenKeepTemp: false,
    youtubeFixWithAi: false,
    youtubePrimarySubLangs: [],
    youtubeSecondarySubLangs: [],
    youtubeAudioLangs: [],
    youtubeWhisperSourceLanguage: 'ja',
    aiConfig: {},
    useTexthooker: false,
    autoStartOverlay: false,
    texthookerOnly: false,
    texthookerOpenBrowser: false,
    useRofi: false,
    logLevel: 'error',
    passwordStore: '',
    target: '',
    targetKind: '',
    jimakuApiKey: '',
    jimakuApiKeyCommand: '',
    jimakuApiBaseUrl: '',
    jimakuLanguagePreference: 'none',
    jimakuMaxEntryResults: 10,
    jellyfin: false,
    jellyfinLogin: false,
    jellyfinLogout: false,
    jellyfinPlay: false,
    jellyfinDiscovery: false,
    dictionary: false,
    dictionaryCandidates: false,
    dictionarySelect: false,
    stats: false,
    doctor: false,
    doctorRefreshKnownWords: false,
    version: false,
    settings: false,
    configPath: false,
    configShow: false,
    mpvIdle: false,
    mpvSocket: false,
    mpvStatus: false,
    mpvArgs: '',
    appPassthrough: false,
    appArgs: [],
    jellyfinServer: '',
    jellyfinUsername: '',
    jellyfinPassword: '',
    launchMode: 'normal',
    ...overrides,
  };
}

test('startOverlay resolves without fixed 2s sleep when readiness signals arrive quickly', async () => {
  const { dir, socketPath } = createTempSocketPath();
  const appPath = path.join(dir, 'fake-subminer.sh');
  fs.writeFileSync(appPath, '#!/bin/sh\nexit 0\n');
  fs.chmodSync(appPath, 0o755);
  fs.writeFileSync(socketPath, '');
  const originalCreateConnection = net.createConnection;
  try {
    net.createConnection = (() => {
      const socket = new EventEmitter() as net.Socket;
      socket.destroy = (() => socket) as net.Socket['destroy'];
      socket.setTimeout = (() => socket) as net.Socket['setTimeout'];
      setTimeout(() => socket.emit('connect'), 10);
      return socket;
    }) as typeof net.createConnection;

    const startedAt = Date.now();
    await startOverlay(appPath, makeArgs(), socketPath);
    const elapsedMs = Date.now() - startedAt;

    assert.ok(elapsedMs < 1200, `expected startOverlay <1200ms, got ${elapsedMs}ms`);
  } finally {
    net.createConnection = originalCreateConnection;
    state.overlayProc = null;
    state.overlayManagedByLauncher = false;
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('startOverlay captures app stdout and stderr into app log', async () => {
  const { dir, socketPath } = createTempSocketPath();
  const appPath = path.join(dir, 'fake-subminer.sh');
  const appLogPath = path.join(dir, 'app.log');
  const originalAppLog = process.env.SUBMINER_APP_LOG;
  fs.writeFileSync(
    appPath,
    '#!/bin/sh\nprintf "hello from stdout\\n"\nprintf "hello from stderr\\n" >&2\nexit 0\n',
  );
  fs.chmodSync(appPath, 0o755);
  fs.writeFileSync(socketPath, '');
  const originalCreateConnection = net.createConnection;
  try {
    process.env.SUBMINER_APP_LOG = appLogPath;
    net.createConnection = (() => {
      const socket = new EventEmitter() as net.Socket;
      socket.destroy = (() => socket) as net.Socket['destroy'];
      socket.setTimeout = (() => socket) as net.Socket['setTimeout'];
      setTimeout(() => socket.emit('connect'), 10);
      return socket;
    }) as typeof net.createConnection;

    await startOverlay(appPath, makeArgs(), socketPath);

    const logText = fs.readFileSync(appLogPath, 'utf8');
    assert.match(logText, /\[STDOUT\] hello from stdout/);
    assert.match(logText, /\[STDERR\] hello from stderr/);
  } finally {
    net.createConnection = originalCreateConnection;
    state.overlayProc = null;
    state.overlayManagedByLauncher = false;
    if (originalAppLog === undefined) {
      delete process.env.SUBMINER_APP_LOG;
    } else {
      process.env.SUBMINER_APP_LOG = originalAppLog;
    }
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('startOverlay starts launcher-owned playback in background managed mode', async () => {
  const { dir, socketPath } = createTempSocketPath();
  const appPath = path.join(dir, 'fake-subminer.sh');
  const appInvocationsPath = path.join(dir, 'app-invocations.log');
  fs.writeFileSync(
    appPath,
    [
      '#!/bin/sh',
      `printf '%s\\n' "$@" >> ${JSON.stringify(appInvocationsPath)}`,
      'if [ "$1" = "--app-ping" ]; then exit 1; fi',
      'exit 0',
      '',
    ].join('\n'),
  );
  fs.chmodSync(appPath, 0o755);
  fs.writeFileSync(socketPath, '');
  const originalCreateConnection = net.createConnection;
  try {
    net.createConnection = (() => {
      const socket = new EventEmitter() as net.Socket;
      socket.destroy = (() => socket) as net.Socket['destroy'];
      socket.setTimeout = (() => socket) as net.Socket['setTimeout'];
      setTimeout(() => socket.emit('connect'), 10);
      return socket;
    }) as typeof net.createConnection;

    await startOverlay(appPath, makeArgs(), socketPath);

    const invocationText = fs.readFileSync(appInvocationsPath, 'utf8');
    assert.match(invocationText, /--background/);
    assert.match(invocationText, /--managed-playback/);
    assert.equal(state.overlayManagedByLauncher, true);
    assert.equal(state.appPath, appPath);
  } finally {
    net.createConnection = originalCreateConnection;
    state.overlayProc = null;
    state.overlayManagedByLauncher = false;
    state.appPath = '';
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('startOverlay borrows an already-running background app instead of owning its lifecycle', async () => {
  const { dir, socketPath } = createTempSocketPath();
  const appPath = path.join(dir, 'fake-subminer.sh');
  const appInvocationsPath = path.join(dir, 'app-invocations.log');
  fs.writeFileSync(
    appPath,
    [
      '#!/bin/sh',
      `printf '%s\\n' "$@" >> ${JSON.stringify(appInvocationsPath)}`,
      'if [ "$1" = "--app-ping" ]; then exit 0; fi',
      'exit 0',
      '',
    ].join('\n'),
  );
  fs.chmodSync(appPath, 0o755);
  fs.writeFileSync(socketPath, '');
  const originalCreateConnection = net.createConnection;
  try {
    net.createConnection = (() => {
      const socket = new EventEmitter() as net.Socket;
      socket.destroy = (() => socket) as net.Socket['destroy'];
      socket.setTimeout = (() => socket) as net.Socket['setTimeout'];
      setTimeout(() => socket.emit('connect'), 10);
      return socket;
    }) as typeof net.createConnection;

    await startOverlay(appPath, makeArgs(), socketPath);

    const invocationText = fs.readFileSync(appInvocationsPath, 'utf8');
    assert.match(invocationText, /--app-ping/);
    assert.match(invocationText, /--start/);
    assert.doesNotMatch(invocationText, /--background/);
    assert.equal(state.overlayManagedByLauncher, false);
    assert.equal(state.appPath, '');
  } finally {
    net.createConnection = originalCreateConnection;
    state.overlayProc = null;
    state.overlayManagedByLauncher = false;
    state.appPath = '';
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('startOverlay attaches through the running app control socket without spawning another app command', async () => {
  if (process.platform === 'win32') return;

  const { dir, socketPath } = createTempSocketPath();
  const controlSocketPath = path.join(dir, 'control.sock');
  const appPath = path.join(dir, 'fake-subminer.sh');
  const appInvocationsPath = path.join(dir, 'app-invocations.log');
  const receivedControlArgv: string[][] = [];
  const originalControlSocket = process.env.SUBMINER_APP_CONTROL_SOCKET;

  fs.writeFileSync(
    appPath,
    [
      '#!/bin/sh',
      `printf '%s\\n' "$@" >> ${JSON.stringify(appInvocationsPath)}`,
      'if [ "$1" = "--app-ping" ]; then exit 0; fi',
      'exit 0',
      '',
    ].join('\n'),
  );
  fs.chmodSync(appPath, 0o755);

  const mpvServer = net.createServer((socket) => socket.end());
  const controlServer = net.createServer((socket) => {
    let buffer = '';
    socket.on('data', (chunk) => {
      buffer += chunk.toString('utf8');
      const newlineMatch = buffer.match(/\r?\n/);
      if (!newlineMatch || newlineMatch.index === undefined) return;
      const line = buffer.slice(0, newlineMatch.index).trim();
      buffer = buffer.slice(newlineMatch.index + newlineMatch[0].length);
      if (!line) return;
      const payload = JSON.parse(line) as { argv?: unknown };
      if (Array.isArray(payload.argv)) {
        receivedControlArgv.push(
          payload.argv.filter((value): value is string => typeof value === 'string'),
        );
      }
      socket.end(JSON.stringify({ ok: true }) + '\n');
    });
  });

  try {
    process.env.SUBMINER_APP_CONTROL_SOCKET = controlSocketPath;
    await new Promise<void>((resolve, reject) => {
      mpvServer.once('error', reject);
      mpvServer.listen(socketPath, resolve);
    });
    await new Promise<void>((resolve, reject) => {
      controlServer.once('error', reject);
      controlServer.listen(controlSocketPath, resolve);
    });

    await startOverlay(appPath, makeArgs(), socketPath);

    const invocationText = fs.existsSync(appInvocationsPath)
      ? fs.readFileSync(appInvocationsPath, 'utf8')
      : '';
    assert.equal(invocationText, '');
    assert.equal(receivedControlArgv.length, 1);
    assert.deepEqual(receivedControlArgv[0]?.slice(0, 7), [
      '--start',
      '--managed-playback',
      '--backend',
      'x11',
      '--socket',
      socketPath,
      '--log-level',
    ]);
    assert.equal(state.overlayManagedByLauncher, false);
    assert.equal(state.appPath, '');
  } finally {
    if (originalControlSocket === undefined) {
      delete process.env.SUBMINER_APP_CONTROL_SOCKET;
    } else {
      process.env.SUBMINER_APP_CONTROL_SOCKET = originalControlSocket;
    }
    await new Promise<void>((resolve) => mpvServer.close(() => resolve()));
    await new Promise<void>((resolve) => controlServer.close(() => resolve()));
    state.overlayProc = null;
    state.overlayManagedByLauncher = false;
    state.appPath = '';
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('startOverlay uses caller config dir for app control socket discovery', async () => {
  if (process.platform === 'win32') return;

  const { dir, socketPath } = createTempSocketPath();
  const configDir = path.join(dir, 'launcher-config');
  const controlSocketPath = getAppControlSocketPath({ configDir, platform: 'linux' });
  fs.mkdirSync(configDir, { recursive: true });
  const appPath = path.join(dir, 'fake-subminer.sh');
  const appInvocationsPath = path.join(dir, 'app-invocations.log');
  const receivedControlArgv: string[][] = [];
  const originalControlSocket = process.env.SUBMINER_APP_CONTROL_SOCKET;

  fs.writeFileSync(
    appPath,
    [
      '#!/bin/sh',
      `printf '%s\\n' "$@" >> ${JSON.stringify(appInvocationsPath)}`,
      'if [ "$1" = "--app-ping" ]; then exit 0; fi',
      'exit 0',
      '',
    ].join('\n'),
  );
  fs.chmodSync(appPath, 0o755);

  const mpvServer = net.createServer((socket) => socket.end());
  const controlServer = net.createServer((socket) => {
    let buffer = '';
    socket.on('data', (chunk) => {
      buffer += chunk.toString('utf8');
      const newlineIndex = buffer.indexOf('\n');
      if (newlineIndex < 0) return;
      const payload = JSON.parse(buffer.slice(0, newlineIndex)) as { argv?: unknown };
      if (Array.isArray(payload.argv)) {
        receivedControlArgv.push(
          payload.argv.filter((value): value is string => typeof value === 'string'),
        );
      }
      socket.end(JSON.stringify({ ok: true }) + '\n');
    });
  });

  try {
    delete process.env.SUBMINER_APP_CONTROL_SOCKET;
    await new Promise<void>((resolve, reject) => {
      mpvServer.once('error', reject);
      mpvServer.listen(socketPath, resolve);
    });
    await new Promise<void>((resolve, reject) => {
      controlServer.once('error', reject);
      controlServer.listen(controlSocketPath, resolve);
    });

    await startOverlay(appPath, makeArgs(), socketPath, [], configDir);

    const invocationText = fs.existsSync(appInvocationsPath)
      ? fs.readFileSync(appInvocationsPath, 'utf8')
      : '';
    assert.equal(invocationText, '');
    assert.equal(receivedControlArgv.length, 1);
    assert.deepEqual(receivedControlArgv[0]?.slice(0, 6), [
      '--start',
      '--managed-playback',
      '--backend',
      'x11',
      '--socket',
      socketPath,
    ]);
  } finally {
    if (originalControlSocket === undefined) {
      delete process.env.SUBMINER_APP_CONTROL_SOCKET;
    } else {
      process.env.SUBMINER_APP_CONTROL_SOCKET = originalControlSocket;
    }
    await new Promise<void>((resolve) => mpvServer.close(() => resolve()));
    await new Promise<void>((resolve) => controlServer.close(() => resolve()));
    state.overlayProc = null;
    state.overlayManagedByLauncher = false;
    state.appPath = '';
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('startOverlay falls back to legacy app startup when control command fails', async () => {
  if (process.platform === 'win32') return;

  const { dir, socketPath } = createTempSocketPath();
  const controlSocketPath = path.join(dir, 'control.sock');
  const appPath = path.join(dir, 'fake-subminer.sh');
  const appInvocationsPath = path.join(dir, 'app-invocations.log');
  const originalControlSocket = process.env.SUBMINER_APP_CONTROL_SOCKET;

  fs.writeFileSync(
    appPath,
    [
      '#!/bin/sh',
      `printf '%s\\n' "$@" >> ${JSON.stringify(appInvocationsPath)}`,
      'if [ "$1" = "--app-ping" ]; then exit 0; fi',
      'exit 0',
      '',
    ].join('\n'),
  );
  fs.chmodSync(appPath, 0o755);

  const controlServer = net.createServer((socket) => {
    socket.on('data', () => {
      socket.end(JSON.stringify({ ok: false, error: 'boom' }) + '\n');
    });
  });

  try {
    process.env.SUBMINER_APP_CONTROL_SOCKET = controlSocketPath;
    await new Promise<void>((resolve, reject) => {
      controlServer.once('error', reject);
      controlServer.listen(controlSocketPath, resolve);
    });

    await startOverlay(appPath, makeArgs(), socketPath);

    const invocationText = fs.readFileSync(appInvocationsPath, 'utf8');
    assert.match(invocationText, /--app-ping/);
    assert.match(invocationText, /--start/);
  } finally {
    if (originalControlSocket === undefined) {
      delete process.env.SUBMINER_APP_CONTROL_SOCKET;
    } else {
      process.env.SUBMINER_APP_CONTROL_SOCKET = originalControlSocket;
    }
    await new Promise<void>((resolve) => controlServer.close(() => resolve()));
    state.overlayProc = null;
    state.overlayManagedByLauncher = false;
    state.appPath = '';
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('startOverlay keeps lifecycle ownership for its already-managed app', async () => {
  const { dir, socketPath } = createTempSocketPath();
  const appPath = path.join(dir, 'fake-subminer.sh');
  const appInvocationsPath = path.join(dir, 'app-invocations.log');
  fs.writeFileSync(
    appPath,
    [
      '#!/bin/sh',
      `printf '%s\\n' "$@" >> ${JSON.stringify(appInvocationsPath)}`,
      'if [ "$1" = "--app-ping" ]; then exit 0; fi',
      'exit 0',
      '',
    ].join('\n'),
  );
  fs.chmodSync(appPath, 0o755);
  fs.writeFileSync(socketPath, '');
  const originalCreateConnection = net.createConnection;
  try {
    state.appPath = appPath;
    state.overlayManagedByLauncher = true;
    net.createConnection = (() => {
      const socket = new EventEmitter() as net.Socket;
      socket.destroy = (() => socket) as net.Socket['destroy'];
      socket.setTimeout = (() => socket) as net.Socket['setTimeout'];
      setTimeout(() => socket.emit('connect'), 10);
      return socket;
    }) as typeof net.createConnection;

    await startOverlay(appPath, makeArgs(), socketPath);

    assert.equal(state.overlayManagedByLauncher, true);
    assert.equal(state.appPath, appPath);
  } finally {
    net.createConnection = originalCreateConnection;
    state.overlayProc = null;
    state.overlayManagedByLauncher = false;
    state.appPath = '';
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('cleanupPlaybackSession stops launcher-managed overlay app and mpv-owned children', async () => {
  const { dir } = createTempSocketPath();
  const appPath = path.join(dir, 'fake-subminer.sh');
  const appInvocationsPath = path.join(dir, 'app-invocations.log');
  fs.writeFileSync(
    appPath,
    `#!/bin/sh\necho \"$@\" >> ${JSON.stringify(appInvocationsPath)}\nexit 0\n`,
  );
  fs.chmodSync(appPath, 0o755);

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
  const helperProc = {
    killed: false,
    kill: () => {
      calls.push('helper-kill');
      return true;
    },
  } as unknown as NonNullable<typeof state.overlayProc>;

  state.stopRequested = false;
  state.appPath = appPath;
  state.overlayManagedByLauncher = true;
  state.overlayProc = overlayProc;
  state.mpvProc = mpvProc;
  state.youtubeSubgenChildren.add(helperProc);

  try {
    await cleanupPlaybackSession(makeArgs());

    assert.deepEqual(calls, ['overlay-kill', 'mpv-kill', 'helper-kill']);
    assert.match(fs.readFileSync(appInvocationsPath, 'utf8'), /--stop/);
  } finally {
    state.overlayProc = null;
    state.mpvProc = null;
    state.youtubeSubgenChildren.clear();
    state.overlayManagedByLauncher = false;
    state.appPath = '';
    state.stopRequested = false;
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

// ── findAppBinary: Linux packaged path discovery ──────────────────────────────

function makeExecutable(filePath: string): void {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, '#!/bin/sh\nexit 0\n');
  fs.chmodSync(filePath, 0o755);
}

function withFindAppBinaryEnvSandbox(run: () => void): void {
  const originalAppImagePath = process.env.SUBMINER_APPIMAGE_PATH;
  const originalBinaryPath = process.env.SUBMINER_BINARY_PATH;
  try {
    delete process.env.SUBMINER_APPIMAGE_PATH;
    delete process.env.SUBMINER_BINARY_PATH;
    run();
  } finally {
    if (originalAppImagePath === undefined) {
      delete process.env.SUBMINER_APPIMAGE_PATH;
    } else {
      process.env.SUBMINER_APPIMAGE_PATH = originalAppImagePath;
    }
    if (originalBinaryPath === undefined) {
      delete process.env.SUBMINER_BINARY_PATH;
    } else {
      process.env.SUBMINER_BINARY_PATH = originalBinaryPath;
    }
  }
}

function withFindAppBinaryPlatformSandbox(
  platform: NodeJS.Platform,
  run: (pathModule: typeof path) => void,
): void {
  const originalPlatform = process.platform;
  try {
    Object.defineProperty(process, 'platform', { value: platform, configurable: true });
    withFindAppBinaryEnvSandbox(() =>
      run(platform === 'win32' ? (path.win32 as typeof path) : path),
    );
  } finally {
    Object.defineProperty(process, 'platform', { value: originalPlatform, configurable: true });
  }
}

function withAccessSyncStub(
  isExecutablePath: (filePath: string) => boolean,
  run: () => void,
): void {
  const originalAccessSync = fs.accessSync;
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (fs as any).accessSync = (filePath: string): void => {
      if (isExecutablePath(filePath)) {
        return;
      }
      throw Object.assign(new Error(`EACCES: ${filePath}`), { code: 'EACCES' });
    };
    run();
  } finally {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (fs as any).accessSync = originalAccessSync;
  }
}

function withExistsAndStatSyncStubs(
  options: {
    existingPaths?: string[];
    directoryPaths?: string[];
  },
  run: () => void,
): void {
  const existingPaths = new Set(options.existingPaths ?? []);
  const directoryPaths = new Set(options.directoryPaths ?? []);
  const originalExistsSync = fs.existsSync;
  const originalStatSync = fs.statSync;
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (fs as any).existsSync = (filePath: string): boolean =>
      existingPaths.has(filePath) || directoryPaths.has(filePath);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (fs as any).statSync = (filePath: string) => {
      if (directoryPaths.has(filePath)) {
        return { isDirectory: () => true };
      }
      if (existingPaths.has(filePath)) {
        return { isDirectory: () => false };
      }
      throw Object.assign(new Error(`ENOENT: ${filePath}`), { code: 'ENOENT' });
    };
    run();
  } finally {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (fs as any).existsSync = originalExistsSync;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (fs as any).statSync = originalStatSync;
  }
}

function withRealpathSyncStub(resolvePath: (filePath: string) => string, run: () => void): void {
  const originalRealpathSync = fs.realpathSync;
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (fs as any).realpathSync = (filePath: string): string => resolvePath(filePath);
    run();
  } finally {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (fs as any).realpathSync = originalRealpathSync;
  }
}

function listRepoRootWindowsTempArtifacts(): string[] {
  return fs
    .readdirSync(process.cwd())
    .filter((entry) => /^\\tmp\\subminer-test-win-/.test(entry))
    .sort();
}

function runFindAppBinaryWindowsPathCase(): void {
  const baseDir = 'C:\\Users\\tester\\subminer-test-win-path';
  const originalHomedir = os.homedir;
  const originalPath = process.env.PATH;
  try {
    os.homedir = () => baseDir;
    const binDir = path.win32.join(baseDir, 'bin');
    const wrapperPath = path.win32.join(binDir, 'SubMiner.exe');
    process.env.PATH = `${binDir}${path.win32.delimiter}${originalPath ?? ''}`;

    withFindAppBinaryPlatformSandbox('win32', (pathModule) => {
      withAccessSyncStub(
        (filePath) => filePath === wrapperPath,
        () => {
          const result = findAppBinary(
            pathModule.join(baseDir, 'launcher', 'SubMiner.exe'),
            pathModule,
          );
          assert.equal(result, wrapperPath);
        },
      );
    });
  } finally {
    os.homedir = originalHomedir;
    process.env.PATH = originalPath;
  }
}

function runFindAppBinaryWindowsInstallDirCase(): void {
  const baseDir = 'C:\\Users\\tester\\subminer-test-win-dir';
  const originalHomedir = os.homedir;
  const originalSubminerBinaryPath = process.env.SUBMINER_BINARY_PATH;
  try {
    os.homedir = () => baseDir;
    const installDir = path.win32.join(baseDir, 'Programs', 'SubMiner');
    const appExe = path.win32.join(installDir, 'SubMiner.exe');
    process.env.SUBMINER_BINARY_PATH = installDir;

    withPlatform('win32', () => {
      withExistsAndStatSyncStubs({ existingPaths: [appExe], directoryPaths: [installDir] }, () => {
        withAccessSyncStub(
          (filePath) => filePath === appExe,
          () => {
            const result = findAppBinary(
              path.win32.join(baseDir, 'launcher', 'SubMiner.exe'),
              path.win32,
            );
            assert.equal(result, appExe);
          },
        );
      });
    });
  } finally {
    os.homedir = originalHomedir;
    if (originalSubminerBinaryPath === undefined) {
      delete process.env.SUBMINER_BINARY_PATH;
    } else {
      process.env.SUBMINER_BINARY_PATH = originalSubminerBinaryPath;
    }
  }
}

test(
  'findAppBinary resolves ~/.local/bin/SubMiner.AppImage when it exists',
  { concurrency: false },
  () => {
    const baseDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-test-home-'));
    const originalHomedir = os.homedir;
    try {
      os.homedir = () => baseDir;
      const appImage = path.join(baseDir, '.local/bin/SubMiner.AppImage');
      makeExecutable(appImage);

      withFindAppBinaryPlatformSandbox('linux', (pathModule) => {
        const result = findAppBinary('/some/other/path/subminer', pathModule);
        assert.equal(result, appImage);
      });
    } finally {
      os.homedir = originalHomedir;
      fs.rmSync(baseDir, { recursive: true, force: true });
    }
  },
);

test(
  'findAppBinary resolves /opt/SubMiner/SubMiner.AppImage when ~/.local/bin candidate does not exist',
  { concurrency: false },
  () => {
    const baseDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-test-home-'));
    const originalHomedir = os.homedir;
    try {
      os.homedir = () => baseDir;
      withFindAppBinaryPlatformSandbox('linux', (pathModule) => {
        withAccessSyncStub(
          (filePath) => filePath === '/opt/SubMiner/SubMiner.AppImage',
          () => {
            const result = findAppBinary('/some/other/path/subminer', pathModule);
            assert.equal(result, '/opt/SubMiner/SubMiner.AppImage');
          },
        );
      });
    } finally {
      os.homedir = originalHomedir;
      fs.rmSync(baseDir, { recursive: true, force: true });
    }
  },
);

test(
  'findAppBinary finds subminer on PATH when AppImage candidates do not exist',
  { concurrency: false },
  () => {
    const baseDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-test-path-'));
    const originalHomedir = os.homedir;
    const originalPath = process.env.PATH;
    try {
      os.homedir = () => baseDir;
      // No AppImage candidates in empty home dir; place subminer wrapper on PATH
      const binDir = path.join(baseDir, 'bin');
      const wrapperPath = path.join(binDir, 'subminer');
      makeExecutable(wrapperPath);
      process.env.PATH = `${binDir}${path.delimiter}${originalPath ?? ''}`;

      withFindAppBinaryPlatformSandbox('linux', (pathModule) => {
        withAccessSyncStub(
          (filePath) => filePath === wrapperPath,
          () => {
            // selfPath must differ from wrapperPath so the self-check does not exclude it
            const result = findAppBinary(path.join(baseDir, 'launcher', 'subminer'), pathModule);
            assert.equal(result, wrapperPath);
          },
        );
      });
    } finally {
      os.homedir = originalHomedir;
      process.env.PATH = originalPath;
      fs.rmSync(baseDir, { recursive: true, force: true });
    }
  },
);

test(
  'findAppBinary excludes PATH matches that canonicalize to the launcher path',
  { concurrency: false },
  () => {
    const baseDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-test-realpath-'));
    const originalHomedir = os.homedir;
    const originalPath = process.env.PATH;
    try {
      os.homedir = () => baseDir;
      const binDir = path.join(baseDir, 'bin');
      const wrapperPath = path.join(binDir, 'subminer');
      const canonicalPath = path.join(baseDir, 'launch', 'subminer');
      makeExecutable(wrapperPath);
      process.env.PATH = `${binDir}${path.delimiter}${originalPath ?? ''}`;

      withFindAppBinaryPlatformSandbox('linux', (pathModule) => {
        withAccessSyncStub(
          (filePath) => filePath === wrapperPath,
          () => {
            withRealpathSyncStub(
              (filePath) => {
                if (filePath === canonicalPath || filePath === wrapperPath) {
                  return canonicalPath;
                }
                return filePath;
              },
              () => {
                const result = findAppBinary(canonicalPath, pathModule);
                assert.equal(result, null);
              },
            );
          },
        );
      });
    } finally {
      os.homedir = originalHomedir;
      process.env.PATH = originalPath;
      fs.rmSync(baseDir, { recursive: true, force: true });
    }
  },
);

test('findAppBinary resolves Windows install paths when present', { concurrency: false }, () => {
  const baseDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-test-win-'));
  const originalHomedir = os.homedir;
  const originalLocalAppData = process.env.LOCALAPPDATA;
  try {
    os.homedir = () => baseDir;
    process.env.LOCALAPPDATA = path.win32.join(baseDir, 'AppData', 'Local');
    const appExe = path.win32.join(
      baseDir,
      'AppData',
      'Local',
      'Programs',
      'SubMiner',
      'SubMiner.exe',
    );

    withFindAppBinaryPlatformSandbox('win32', (pathModule) => {
      withAccessSyncStub(
        (filePath) => filePath === appExe,
        () => {
          const result = findAppBinary(
            pathModule.join(baseDir, 'launcher', 'SubMiner.exe'),
            pathModule,
          );
          assert.equal(result, appExe);
        },
      );
    });
  } finally {
    os.homedir = originalHomedir;
    if (originalLocalAppData === undefined) {
      delete process.env.LOCALAPPDATA;
    } else {
      process.env.LOCALAPPDATA = originalLocalAppData;
    }
    fs.rmSync(baseDir, { recursive: true, force: true });
  }
});

test(
  'findAppBinary Windows cases do not leak backslash temp artifacts on POSIX',
  { concurrency: false },
  () => {
    if (path.sep === '\\') {
      return;
    }

    const before = listRepoRootWindowsTempArtifacts();
    runFindAppBinaryWindowsPathCase();
    runFindAppBinaryWindowsInstallDirCase();
    const after = listRepoRootWindowsTempArtifacts();

    assert.deepEqual(after, before);
  },
);

test('findAppBinary resolves SubMiner.exe on PATH on Windows', { concurrency: false }, () => {
  runFindAppBinaryWindowsPathCase();
});

test(
  'findAppBinary resolves a Windows install directory to SubMiner.exe',
  { concurrency: false },
  () => {
    runFindAppBinaryWindowsInstallDirCase();
  },
);
