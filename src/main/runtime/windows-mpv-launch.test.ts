import test from 'node:test';
import assert from 'node:assert/strict';
import type { SubminerPluginRuntimeScriptOptConfig } from '../../shared/subminer-plugin-script-opts';
import {
  buildWindowsMpvLaunchArgs,
  launchWindowsMpv,
  resolveWindowsMpvPath,
  type WindowsMpvLaunchDeps,
  type WindowsMpvRuntimePluginPolicy,
} from './windows-mpv-launch';

const MPV_PATH = 'C:\\mpv\\mpv.exe';
const BINARY_PATH = 'C:\\SubMiner\\SubMiner.exe';
const PLUGIN_PATH = 'C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua';
const DEFAULT_PIPE = '\\\\.\\pipe\\subminer-socket';
const CUSTOM_PIPE = '\\\\.\\pipe\\custom-subminer-socket';

function createDeps(overrides: Partial<WindowsMpvLaunchDeps> = {}): WindowsMpvLaunchDeps {
  return {
    getEnv: () => undefined,
    runWhere: () => ({ status: 1, stdout: '' }),
    fileExists: () => false,
    spawnDetached: async () => undefined,
    showError: () => undefined,
    ...overrides,
  };
}

// Deps where SUBMINER_MPV_PATH resolves to an existing mpv.exe.
function createResolvedDeps(overrides: Partial<WindowsMpvLaunchDeps> = {}): WindowsMpvLaunchDeps {
  return createDeps({
    getEnv: (name) => (name === 'SUBMINER_MPV_PATH' ? MPV_PATH : undefined),
    fileExists: (candidate) => candidate === MPV_PATH,
    ...overrides,
  });
}

function pluginConfig(
  overrides: Partial<SubminerPluginRuntimeScriptOptConfig> = {},
): SubminerPluginRuntimeScriptOptConfig {
  return {
    socketPath: DEFAULT_PIPE,
    binaryPath: '',
    backend: 'windows',
    autoStart: true,
    autoStartVisibleOverlay: true,
    autoStartPauseUntilReady: true,
    osdMessages: false,
    texthookerEnabled: false,
    ...overrides,
  };
}

type BuildArgsOptions = {
  targets?: string[];
  extraArgs?: string[];
  binaryPath?: string;
  launchMode?: Parameters<typeof buildWindowsMpvLaunchArgs>[4];
  runtimeConfig?: SubminerPluginRuntimeScriptOptConfig;
};

function buildArgs(options: BuildArgsOptions = {}): string[] {
  return buildWindowsMpvLaunchArgs(
    options.targets ?? ['C:\\video.mkv'],
    options.extraArgs ?? [],
    'binaryPath' in options ? options.binaryPath : BINARY_PATH,
    PLUGIN_PATH,
    options.launchMode ?? 'normal',
    options.runtimeConfig,
  );
}

type LaunchOptions = {
  targets?: string[];
  extraArgs?: string[];
  policy?: WindowsMpvRuntimePluginPolicy;
  runtimeConfig?: SubminerPluginRuntimeScriptOptConfig;
};

function launch(deps: WindowsMpvLaunchDeps, options: LaunchOptions = {}) {
  return launchWindowsMpv(
    options.targets ?? ['C:\\video.mkv'],
    deps,
    options.extraArgs ?? [],
    BINARY_PATH,
    PLUGIN_PATH,
    '',
    'normal',
    options.policy,
    options.runtimeConfig,
  );
}

function scriptOptsOf(args: readonly string[] | undefined): string {
  return args?.find((arg) => arg.startsWith('--script-opts=')) ?? '';
}

test('resolveWindowsMpvPath prefers SUBMINER_MPV_PATH', () => {
  assert.equal(resolveWindowsMpvPath(createResolvedDeps()), MPV_PATH);
});

test('resolveWindowsMpvPath prefers configured executable path before environment and PATH', () => {
  const resolved = resolveWindowsMpvPath(
    createDeps({
      getEnv: () => 'C:\\other\\mpv.exe',
      runWhere: () => ({ status: 0, stdout: 'C:\\tools\\mpv.exe\r\n' }),
      fileExists: (candidate) => [MPV_PATH, 'C:\\other\\mpv.exe'].includes(candidate),
    }),
    '  C:\\mpv\\mpv.exe  ',
  );

  assert.equal(resolved, MPV_PATH);
});

test('resolveWindowsMpvPath falls back to where.exe output', () => {
  const resolved = resolveWindowsMpvPath(
    createDeps({
      runWhere: () => ({ status: 0, stdout: 'C:\\tools\\mpv.exe\r\nC:\\other\\mpv.exe\r\n' }),
      fileExists: (candidate) => candidate === 'C:\\tools\\mpv.exe',
    }),
  );

  assert.equal(resolved, 'C:\\tools\\mpv.exe');
});

test('resolveWindowsMpvPath ignores an invalid environment override but keeps config authoritative', () => {
  const deps = createDeps({
    getEnv: () => 'C:\\missing\\mpv.exe',
    runWhere: () => ({ status: 0, stdout: 'C:\\tools\\mpv.exe\r\n' }),
    fileExists: (candidate) => candidate === 'C:\\tools\\mpv.exe',
  });
  assert.equal(resolveWindowsMpvPath(deps), 'C:\\tools\\mpv.exe');
  assert.equal(resolveWindowsMpvPath(deps, 'C:\\missing\\mpv.exe'), '');
});

test('buildWindowsMpvLaunchArgs injects the plugin, default pipe and script opts before targets', () => {
  const args = buildArgs({ targets: ['C:\\a.mkv', 'C:\\b.mkv'] });

  assert.deepEqual(args.slice(0, 2), [
    '--player-operation-mode=pseudo-gui',
    '--force-window=immediate',
  ]);
  assert.ok(args.includes(`--script=${PLUGIN_PATH}`));
  assert.ok(args.includes(`--input-ipc-server=${DEFAULT_PIPE}`));
  assert.equal(
    scriptOptsOf(args),
    `--script-opts=subminer-binary_path=${BINARY_PATH},subminer-socket_path=${DEFAULT_PIPE}`,
  );
  assert.equal(args.includes('--idle=yes'), false);
  assert.deepEqual(args.slice(-2), ['C:\\a.mkv', 'C:\\b.mkv']);
});

test('buildWindowsMpvLaunchArgs keeps shortcut-only launches in idle mode', () => {
  assert.ok(buildArgs({ targets: [] }).includes('--idle=yes'));
});

test('buildWindowsMpvLaunchArgs inserts maximized launch mode before explicit extra args', () => {
  const args = buildArgs({ extraArgs: ['--window-maximized=no'], launchMode: 'maximized' });

  const maximized = args.indexOf('--window-maximized=yes');
  const explicit = args.indexOf('--window-maximized=no');
  assert.ok(maximized >= 0);
  assert.ok(explicit > maximized);
  assert.ok(args.indexOf('C:\\video.mkv') > explicit);
});

test('buildWindowsMpvLaunchArgs mirrors a custom input-ipc-server into script opts', () => {
  const extraArgs = ['--input-ipc-server', CUSTOM_PIPE];
  const args = buildArgs({ extraArgs });

  assert.ok(args.includes(`--input-ipc-server=${CUSTOM_PIPE}`));
  assert.equal(
    scriptOptsOf(args),
    `--script-opts=subminer-binary_path=${BINARY_PATH},subminer-socket_path=${CUSTOM_PIPE}`,
  );
  assert.deepEqual(args.slice(-3), [...extraArgs, 'C:\\video.mkv']);
});

test('buildWindowsMpvLaunchArgs includes socket script opts when plugin entrypoint is present without binary path', () => {
  const args = buildArgs({
    extraArgs: ['--input-ipc-server', CUSTOM_PIPE],
    binaryPath: undefined,
  });

  assert.equal(scriptOptsOf(args), `--script-opts=subminer-socket_path=${CUSTOM_PIPE}`);
});

test('buildWindowsMpvLaunchArgs uses runtime plugin config script opts', () => {
  const scriptOpts = scriptOptsOf(
    buildArgs({
      extraArgs: ['--input-ipc-server', CUSTOM_PIPE],
      runtimeConfig: pluginConfig({
        socketPath: '\\\\.\\pipe\\ignored-config-socket',
        binaryPath: 'C:\\Custom\\SubMiner.exe',
        autoStartVisibleOverlay: false,
        autoStartPauseUntilReady: false,
      }),
    }),
  );

  assert.match(scriptOpts, /subminer-binary_path=C:\\Custom\\SubMiner\.exe/);
  assert.match(scriptOpts, /subminer-socket_path=\\\\\.\\pipe\\custom-subminer-socket/);
  assert.match(scriptOpts, /subminer-backend=windows/);
  assert.match(scriptOpts, /subminer-auto_start=yes/);
  assert.match(scriptOpts, /subminer-auto_start_visible_overlay=no/);
  assert.match(scriptOpts, /subminer-auto_start_pause_until_ready=no/);
  assert.match(scriptOpts, /subminer-texthooker_enabled=no/);
});

test('buildWindowsMpvLaunchArgs keeps Windows ipc default unless explicitly overridden', () => {
  const args = buildArgs({
    runtimeConfig: pluginConfig({
      socketPath: 'C:\\Users\\tester\\AppData\\Local\\Temp\\subminer-smoke-sock\\subminer.sock',
    }),
  });

  assert.ok(args.includes(`--input-ipc-server=${DEFAULT_PIPE}`));
  assert.match(scriptOptsOf(args), /subminer-socket_path=\\\\\.\\pipe\\subminer-socket/);
});

test('launchWindowsMpv attaches a launched video to a running app and disables plugin auto-start', async () => {
  const spawnedArgs: string[][] = [];
  const controlArgv: string[][] = [];
  const waitedSockets: Array<{ socketPath: string; timeoutMs: number }> = [];
  const logs: string[] = [];
  const result = await launch(
    createResolvedDeps({
      isAppControlServerAvailable: async () => true,
      waitForSocketReady: async (socketPath, timeoutMs) => {
        waitedSockets.push({ socketPath, timeoutMs });
        return true;
      },
      sendAppControlCommand: async (argv) => {
        controlArgv.push(argv);
        return { ok: true };
      },
      logInfo: (message) => logs.push(message),
      spawnDetached: async (_command, args) => {
        spawnedArgs.push(args);
      },
    }),
    {
      extraArgs: ['--input-ipc-server', '\\\\.\\pipe\\warm-subminer'],
      runtimeConfig: pluginConfig({
        socketPath: '\\\\.\\pipe\\ignored-config-socket',
        logLevel: 'debug',
        logRotation: 7,
        texthookerEnabled: true,
      }),
    },
  );

  assert.equal(result.ok, true);
  const scriptOpts = scriptOptsOf(spawnedArgs[0]);
  assert.match(scriptOpts, /subminer-auto_start=no/);
  assert.match(scriptOpts, /subminer-socket_path=\\\\\.\\pipe\\warm-subminer/);
  assert.deepEqual(waitedSockets, [{ socketPath: '\\\\.\\pipe\\warm-subminer', timeoutMs: 10000 }]);
  assert.deepEqual(controlArgv, [
    [
      '--start',
      '--managed-playback',
      '--log-level',
      'debug',
      '--backend',
      'windows',
      '--socket',
      '\\\\.\\pipe\\warm-subminer',
      '--show-visible-overlay',
      '--texthooker',
    ],
  ]);
  assert.ok(logs.some((line) => line.includes('attachRunningApp=yes')));
  assert.ok(logs.some((line) => line.includes('Attached launched mpv session')));
});

test('launchWindowsMpv leaves plugin auto-start enabled when no running app control socket exists', async () => {
  const spawnedArgs: string[][] = [];
  let controlCalls = 0;
  let waitCalls = 0;
  const result = await launch(
    createResolvedDeps({
      isAppControlServerAvailable: async () => false,
      waitForSocketReady: async () => {
        waitCalls += 1;
        return true;
      },
      sendAppControlCommand: async () => {
        controlCalls += 1;
        return { ok: true };
      },
      spawnDetached: async (_command, args) => {
        spawnedArgs.push(args);
      },
    }),
    { runtimeConfig: pluginConfig() },
  );

  assert.equal(result.ok, true);
  assert.match(scriptOptsOf(spawnedArgs[0]), /subminer-auto_start=yes/);
  assert.equal(waitCalls, 0);
  assert.equal(controlCalls, 0);
});

test('launchWindowsMpv reports missing mpv path', async () => {
  const errors: string[] = [];
  const result = await launchWindowsMpv(
    [],
    createDeps({
      showError: (_title, content) => errors.push(content),
    }),
  );

  assert.equal(result.ok, false);
  assert.equal(result.mpvPath, '');
  assert.match(errors[0] ?? '', /mpv\.executablePath/i);
});

test('launchWindowsMpv spawns the resolved mpv with the bundled plugin and targets', async () => {
  const spawns: Array<{ command: string; args: string[] }> = [];
  const logs: string[] = [];
  const result = await launch(
    createResolvedDeps({
      logInfo: (message) => logs.push(message),
      spawnDetached: async (command, args) => {
        spawns.push({ command, args });
      },
    }),
  );

  assert.deepEqual(result, { ok: true, mpvPath: MPV_PATH });
  assert.equal(spawns.length, 1);
  assert.equal(spawns[0]!.command, MPV_PATH);
  assert.ok(spawns[0]!.args.includes(`--script=${PLUGIN_PATH}`));
  assert.equal(spawns[0]!.args.at(-1), 'C:\\video.mkv');
  assert.match(logs[0] ?? '', /mpvPath=C:\\mpv\\mpv\.exe/);
  assert.match(logs[0] ?? '', /inputIpcServer=\\\\\.\\pipe\\subminer-socket/);
  assert.match(
    logs[0] ?? '',
    /bundledPlugin=C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main\.lua/,
  );
  assert.match(logs[0] ?? '', /installedPlugin=none/);
});

test('launchWindowsMpv forwards runtime logging config to mpv and plugin', async () => {
  let spawned: { args: string[]; env?: NodeJS.ProcessEnv } | null = null;
  const result = await launch(
    createResolvedDeps({
      spawnDetached: async (_command, args, env) => {
        spawned = { args, env };
      },
    }),
    {
      extraArgs: [
        '--log-file=C:\\Users\\tester\\AppData\\Roaming\\SubMiner\\logs\\mpv-2026-05-26.log',
      ],
      runtimeConfig: pluginConfig({
        logLevel: 'debug',
        logRotation: 0,
        autoStartVisibleOverlay: false,
      }),
    },
  );

  assert.equal(result.ok, true);
  assert.ok(spawned);
  const { args, env } = spawned as { args: string[]; env?: NodeJS.ProcessEnv };
  assert.ok(args.includes('--msg-level=all=warn,subminer=debug'));
  assert.doesNotMatch(scriptOptsOf(args), /subminer-log_level=debug/);
  assert.equal(env?.SUBMINER_LOG_LEVEL, 'debug');
  assert.equal(env?.SUBMINER_LOG_ROTATION, '0');
});

test('launchWindowsMpv skips bundled script when installed plugin is detected', async () => {
  const spawnedArgs: string[][] = [];
  const notifications: string[] = [];
  const result = await launch(
    createResolvedDeps({
      spawnDetached: async (_command, args) => {
        spawnedArgs.push(args);
      },
    }),
    {
      policy: {
        detectInstalledMpvPlugin: () => ({
          installed: true,
          path: 'C:\\Users\\tester\\AppData\\Roaming\\mpv\\scripts\\subminer\\main.lua',
          version: null,
          source: 'default-config',
          message: null,
        }),
        notifyInstalledPluginDetected: (detection) => {
          notifications.push(detection.path ?? '');
        },
      },
    },
  );

  assert.equal(result.ok, true);
  assert.equal(spawnedArgs[0]?.includes(`--script=${PLUGIN_PATH}`), false);
  assert.match(scriptOptsOf(spawnedArgs[0]), /subminer-binary_path=C:\\SubMiner\\SubMiner\.exe/);
  assert.deepEqual(notifications, [
    'C:\\Users\\tester\\AppData\\Roaming\\mpv\\scripts\\subminer\\main.lua',
  ]);
});

test('launchWindowsMpv prompts before launch and injects bundled script after legacy plugin removal', async () => {
  const spawnedArgs: string[][] = [];
  const prompts: string[] = [];
  let detectCalls = 0;
  const result = await launch(
    createResolvedDeps({
      spawnDetached: async (_command, args) => {
        spawnedArgs.push(args);
      },
    }),
    {
      policy: {
        detectInstalledMpvPlugin: () => {
          detectCalls += 1;
          return detectCalls === 1
            ? {
                installed: true,
                path: 'C:\\Users\\tester\\AppData\\Roaming\\mpv\\scripts\\subminer\\main.lua',
                version: '0.12.0',
                source: 'default-config',
                message: null,
              }
            : { installed: false, path: null, version: null, source: null, message: null };
        },
        resolveInstalledPluginBeforeLaunch: async (detection) => {
          prompts.push(detection.path ?? '');
          return 'removed' as const;
        },
      },
    },
  );

  assert.equal(result.ok, true);
  assert.equal(detectCalls, 2);
  assert.deepEqual(prompts, [
    'C:\\Users\\tester\\AppData\\Roaming\\mpv\\scripts\\subminer\\main.lua',
  ]);
  assert.ok(spawnedArgs[0]?.includes(`--script=${PLUGIN_PATH}`));
});

const spawnFailureCases: Array<{
  name: string;
  spawnDetached: WindowsMpvLaunchDeps['spawnDetached'];
  message: RegExp;
}> = [
  {
    name: 'synchronous',
    spawnDetached: () => {
      throw new Error('spawn failed');
    },
    message: /spawn failed/,
  },
  {
    name: 'asynchronous',
    spawnDetached: () => Promise.reject(new Error('async spawn failed')),
    message: /async spawn failed/,
  },
];

for (const failure of spawnFailureCases) {
  test(`launchWindowsMpv reports ${failure.name} spawn failures with path context`, async () => {
    const errors: string[] = [];
    const result = await launchWindowsMpv(
      [],
      createResolvedDeps({
        spawnDetached: failure.spawnDetached,
        showError: (_title, content) => errors.push(content),
      }),
    );

    assert.deepEqual(result, { ok: false, mpvPath: MPV_PATH });
    assert.match(errors[0] ?? '', /Failed to launch mpv/i);
    assert.match(errors[0] ?? '', /C:\\mpv\\mpv\.exe/i);
    assert.match(errors[0] ?? '', failure.message);
  });
}
