import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createFirstRunSetupService, shouldAutoOpenFirstRunSetup } from './first-run-setup-service';
import type { CliArgs } from '../../cli/args';
import type { CommandLineLauncherSnapshot } from './command-line-launcher';

function withTempDir(fn: (dir: string) => Promise<void> | void): Promise<void> | void {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-first-run-service-test-'));
  const result = fn(dir);
  if (result instanceof Promise) {
    return result.finally(() => {
      fs.rmSync(dir, { recursive: true, force: true });
    });
  }
  fs.rmSync(dir, { recursive: true, force: true });
}

function makeArgs(overrides: Partial<CliArgs> = {}): CliArgs {
  return {
    background: false,
    managedPlayback: false,
    start: false,
    launchMpv: false,
    launchMpvTargets: [],
    stop: false,
    toggle: false,
    toggleVisibleOverlay: false,
    togglePrimarySubtitleBar: false,
    settings: false,
    setup: false,
    show: false,
    hide: false,
    showVisibleOverlay: false,
    hideVisibleOverlay: false,
    copySubtitle: false,
    copySubtitleMultiple: false,
    mineSentence: false,
    mineSentenceMultiple: false,
    updateLastCardFromClipboard: false,
    refreshKnownWords: false,
    toggleSecondarySub: false,
    triggerFieldGrouping: false,
    triggerSubsync: false,
    markAudioCard: false,
    toggleStatsOverlay: false,
    toggleSubtitleSidebar: false,
    openRuntimeOptions: false,
    openSessionHelp: false,
    openControllerSelect: false,
    openControllerDebug: false,
    openJimaku: false,
    openYoutubePicker: false,
    openPlaylistBrowser: false,
    openCharacterDictionary: false,
    replayCurrentSubtitle: false,
    playNextSubtitle: false,
    shiftSubDelayPrevLine: false,
    shiftSubDelayNextLine: false,
    cycleRuntimeOptionId: undefined,
    cycleRuntimeOptionDirection: undefined,
    anilistStatus: false,
    anilistLogout: false,
    anilistSetup: false,
    anilistRetryQueue: false,
    dictionary: false,
    dictionaryCandidates: false,
    dictionarySelect: false,
    dictionaryAnilistId: undefined,
    stats: false,
    jellyfin: false,
    jellyfinLogin: false,
    jellyfinLogout: false,
    jellyfinLibraries: false,
    jellyfinItems: false,
    jellyfinSubtitles: false,
    jellyfinSubtitleUrlsOnly: false,
    jellyfinPlay: false,
    jellyfinRemoteAnnounce: false,
    jellyfinPreviewAuth: false,
    texthooker: false,
    texthookerOpenBrowser: false,
    update: false,
    help: false,
    autoStartOverlay: false,
    generateConfig: false,
    backupOverwrite: false,
    debug: false,
    ...overrides,
  };
}

function createCommandLineLauncherSnapshot(
  overrides: Partial<CommandLineLauncherSnapshot> = {},
): CommandLineLauncherSnapshot {
  return {
    supported: true,
    bun: {
      status: 'missing',
      commandPath: null,
      version: null,
      installMethod: 'official-script',
      installCommand: ['bash', '-lc', 'curl -fsSL https://bun.com/install | bash'],
      message: null,
    },
    launcher: {
      status: 'not_installed',
      commandPath: null,
      installPath: '/home/tester/.local/bin/subminer',
      pathDir: '/home/tester/.local/bin',
      shadowedBy: null,
      message: null,
    },
    ...overrides,
  };
}

test('shouldAutoOpenFirstRunSetup only for startup/setup intents', () => {
  assert.equal(shouldAutoOpenFirstRunSetup(makeArgs({ start: true, background: true })), true);
  assert.equal(shouldAutoOpenFirstRunSetup(makeArgs({ background: true, setup: true })), true);
  assert.equal(
    shouldAutoOpenFirstRunSetup(makeArgs({ background: true, jellyfinRemoteAnnounce: true })),
    false,
  );
  assert.equal(shouldAutoOpenFirstRunSetup(makeArgs({ settings: true })), false);
  assert.equal(shouldAutoOpenFirstRunSetup(makeArgs({ start: true, update: true })), false);
});

test('shouldAutoOpenFirstRunSetup treats numeric startup counts as explicit commands', () => {
  assert.equal(shouldAutoOpenFirstRunSetup(makeArgs({ start: true, copySubtitleCount: 2 })), false);
  assert.equal(
    shouldAutoOpenFirstRunSetup(makeArgs({ background: true, mineSentenceCount: 1 })),
    false,
  );
});

test('shouldAutoOpenFirstRunSetup treats session and stats startup commands as explicit commands', () => {
  assert.equal(
    shouldAutoOpenFirstRunSetup(makeArgs({ start: true, toggleSubtitleSidebar: true })),
    false,
  );
  assert.equal(
    shouldAutoOpenFirstRunSetup(makeArgs({ background: true, openSessionHelp: true })),
    false,
  );
  assert.equal(
    shouldAutoOpenFirstRunSetup(makeArgs({ start: true, openControllerSelect: true })),
    false,
  );
  assert.equal(
    shouldAutoOpenFirstRunSetup(makeArgs({ background: true, openControllerDebug: true })),
    false,
  );
  assert.equal(shouldAutoOpenFirstRunSetup(makeArgs({ start: true, stats: true })), false);
  assert.equal(
    shouldAutoOpenFirstRunSetup(makeArgs({ background: true, jellyfinSubtitleUrlsOnly: true })),
    false,
  );
});

test('setup service auto-completes legacy installs with config and dictionaries', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');

    const service = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => 2,
      detectPluginInstalled: () => true,
      installPlugin: async () => ({
        ok: true,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: '/tmp/mpv',
        message: 'installed',
      }),
      onStateChanged: () => undefined,
    });

    const snapshot = await service.ensureSetupStateInitialized();
    assert.equal(snapshot.state.status, 'completed');
    assert.equal(snapshot.state.completionSource, 'legacy_auto_detected');
    assert.equal(snapshot.dictionaryCount, 2);
    assert.equal(snapshot.canFinish, true);
  });
});

test('setup service allows finish without global mpv plugin once dictionaries are ready', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');
    let dictionaryCount = 0;

    const service = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => dictionaryCount,
      detectPluginInstalled: () => false,
      installPlugin: async () => ({
        ok: true,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: '/tmp/mpv',
        message: 'installed',
      }),
      onStateChanged: () => undefined,
    });

    const initial = await service.ensureSetupStateInitialized();
    assert.equal(initial.state.status, 'incomplete');
    assert.equal(initial.canFinish, false);

    dictionaryCount = 1;
    const refreshed = await service.refreshStatus();
    assert.equal(refreshed.canFinish, true);

    const completed = await service.markSetupCompleted();
    assert.equal(completed.state.status, 'completed');
    assert.equal(completed.state.completionSource, 'user');
    assert.equal(completed.state.yomitanSetupMode, 'internal');
  });
});

test('setup service allows completion without internal dictionaries when external yomitan is configured', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');

    const service = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => 0,
      isExternalYomitanConfigured: () => true,
      detectPluginInstalled: () => true,
      installPlugin: async () => ({
        ok: true,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: null,
        message: 'ok',
      }),
      onStateChanged: () => undefined,
    });

    const initial = await service.ensureSetupStateInitialized();
    assert.equal(initial.canFinish, true);

    const completed = await service.markSetupCompleted();
    assert.equal(completed.state.status, 'completed');
    assert.equal(completed.state.yomitanSetupMode, 'external');
    assert.equal(completed.dictionaryCount, 0);
  });
});

test('setup service does not probe internal dictionaries when external yomitan is configured', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');

    const service = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => {
        throw new Error('should not probe internal dictionaries in external mode');
      },
      isExternalYomitanConfigured: () => true,
      detectPluginInstalled: () => true,
      installPlugin: async () => ({
        ok: true,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: null,
        message: 'ok',
      }),
      onStateChanged: () => undefined,
    });

    const snapshot = await service.ensureSetupStateInitialized();
    assert.equal(snapshot.state.status, 'completed');
    assert.equal(snapshot.canFinish, true);
    assert.equal(snapshot.externalYomitanConfigured, true);
    assert.equal(snapshot.dictionaryCount, 0);
  });
});

test('setup service reopens when external-yomitan completion later has no external profile and no internal dictionaries', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');

    const service = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => 0,
      isExternalYomitanConfigured: () => true,
      detectPluginInstalled: () => true,
      installPlugin: async () => ({
        ok: true,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: null,
        message: 'ok',
      }),
      onStateChanged: () => undefined,
    });

    await service.ensureSetupStateInitialized();
    await service.markSetupCompleted();

    const relaunched = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => 0,
      isExternalYomitanConfigured: () => false,
      detectPluginInstalled: () => true,
      installPlugin: async () => ({
        ok: true,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: null,
        message: 'ok',
      }),
      onStateChanged: () => undefined,
    });

    const snapshot = await relaunched.ensureSetupStateInitialized();
    assert.equal(snapshot.state.status, 'incomplete');
    assert.equal(snapshot.state.yomitanSetupMode, null);
    assert.equal(snapshot.canFinish, false);
  });
});

test('setup service keeps completed when a global mpv plugin is removed later', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');

    const completedService = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => 2,
      detectPluginInstalled: () => true,
      installPlugin: async () => ({
        ok: true,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: '/tmp/mpv',
        message: 'ok',
      }),
      onStateChanged: () => undefined,
    });

    await completedService.ensureSetupStateInitialized();
    await completedService.markSetupCompleted();

    const service = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => 2,
      detectPluginInstalled: () => false,
      installPlugin: async () => ({
        ok: true,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: null,
        message: 'ok',
      }),
      onStateChanged: () => undefined,
    });

    const snapshot = await service.ensureSetupStateInitialized();
    assert.equal(snapshot.state.status, 'completed');
    assert.equal(snapshot.canFinish, true);
    assert.equal(snapshot.pluginStatus, 'required');
  });
});

test('setup service reopens completed setup as in-progress when legacy mpv plugin removal is needed', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');

    const service = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => 2,
      detectPluginInstalled: () => true,
      detectLegacyMpvPluginCandidates: () => [
        { path: '/tmp/mpv/scripts/subminer', kind: 'directory' },
      ],
      onStateChanged: () => undefined,
    });

    await service.ensureSetupStateInitialized();
    await service.markSetupCompleted();

    const inProgress = await service.markSetupInProgress();
    assert.equal(inProgress.state.status, 'in_progress');
    assert.equal(inProgress.state.completedAt, null);

    const completed = await service.markSetupCompleted();
    assert.equal(completed.state.status, 'completed');
    assert.notEqual(completed.state.completedAt, null);
  });
});

test('setup service keeps completed when external-yomitan completion later has internal dictionaries available', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');

    const service = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => 0,
      isExternalYomitanConfigured: () => true,
      detectPluginInstalled: () => true,
      installPlugin: async () => ({
        ok: true,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: null,
        message: 'ok',
      }),
      onStateChanged: () => undefined,
    });

    await service.ensureSetupStateInitialized();
    await service.markSetupCompleted();

    const relaunched = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => 2,
      isExternalYomitanConfigured: () => false,
      detectPluginInstalled: () => true,
      installPlugin: async () => ({
        ok: true,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: null,
        message: 'ok',
      }),
      onStateChanged: () => undefined,
    });

    const snapshot = await relaunched.ensureSetupStateInitialized();
    assert.equal(snapshot.state.status, 'completed');
    assert.equal(snapshot.canFinish, true);
  });
});

test('setup service marks cancelled when popup closes before completion', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');

    const service = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => 0,
      detectPluginInstalled: () => true,
      installPlugin: async () => ({
        ok: true,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: null,
        message: 'ok',
      }),
      onStateChanged: () => undefined,
    });

    await service.ensureSetupStateInitialized();
    await service.markSetupInProgress();
    const cancelled = await service.markSetupCancelled();
    assert.equal(cancelled.state.status, 'cancelled');
  });
});

test('setup service reflects detected Windows mpv shortcuts before preferences are persisted', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');

    const service = createFirstRunSetupService({
      platform: 'win32',
      configDir,
      getYomitanDictionaryCount: async () => 0,
      detectPluginInstalled: () => true,
      installPlugin: async () => ({
        ok: true,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: null,
        message: 'ok',
      }),
      detectWindowsMpvShortcuts: async () => ({
        startMenuInstalled: false,
        desktopInstalled: true,
      }),
      onStateChanged: () => undefined,
    });

    const snapshot = await service.ensureSetupStateInitialized();
    assert.equal(snapshot.windowsMpvShortcuts.startMenuEnabled, false);
    assert.equal(snapshot.windowsMpvShortcuts.desktopEnabled, true);
    assert.equal(snapshot.windowsMpvShortcuts.startMenuInstalled, false);
    assert.equal(snapshot.windowsMpvShortcuts.desktopInstalled, true);
  });
});

test('setup service persists Windows mpv shortcut preferences and status with one state write', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');
    const stateChanges: string[] = [];

    const service = createFirstRunSetupService({
      platform: 'win32',
      configDir,
      getYomitanDictionaryCount: async () => 0,
      detectPluginInstalled: () => true,
      installPlugin: async () => ({
        ok: true,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: null,
        message: 'ok',
      }),
      applyWindowsMpvShortcuts: async () => ({
        ok: true,
        status: 'installed',
        message: 'shortcuts updated',
      }),
      onStateChanged: (state) => {
        stateChanges.push(state.windowsMpvShortcutLastStatus);
      },
    });

    await service.ensureSetupStateInitialized();
    stateChanges.length = 0;

    const snapshot = await service.configureWindowsMpvShortcuts({
      startMenuEnabled: false,
      desktopEnabled: true,
    });

    assert.equal(snapshot.windowsMpvShortcuts.startMenuEnabled, false);
    assert.equal(snapshot.windowsMpvShortcuts.desktopEnabled, true);
    assert.equal(snapshot.state.windowsMpvShortcutLastStatus, 'installed');
    assert.equal(snapshot.message, 'shortcuts updated');
    assert.deepEqual(stateChanges, ['installed']);
  });
});

test('setup service snapshot includes command-line launcher status', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });

    const commandLineLauncher = createCommandLineLauncherSnapshot({
      bun: {
        status: 'ready',
        commandPath: '/usr/local/bin/bun',
        version: '1.3.5',
        installMethod: null,
        installCommand: null,
        message: null,
      },
    });

    const service = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => 0,
      detectPluginInstalled: () => true,
      detectCommandLineLauncher: async () => commandLineLauncher,
      onStateChanged: () => undefined,
    });

    const snapshot = await service.refreshStatus();
    assert.deepEqual(snapshot.commandLineLauncher, commandLineLauncher);
  });
});

test('setup service installBun persists installed and failed status', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');
    let installOk = true;

    const service = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => 1,
      detectPluginInstalled: () => true,
      detectCommandLineLauncher: async () => createCommandLineLauncherSnapshot(),
      installBun: async () => ({
        ok: installOk,
        message: installOk ? 'Bun installed. Open a new terminal.' : 'Bun install failed.',
      }),
      onStateChanged: () => undefined,
    });

    const installed = await service.installBun();
    assert.equal(installed.state.bunInstallStatus, 'installed');
    assert.equal(installed.canFinish, true);
    assert.equal(installed.message, 'Bun installed. Open a new terminal.');

    installOk = false;
    const failed = await service.installBun();
    assert.equal(failed.state.bunInstallStatus, 'failed');
    assert.equal(failed.canFinish, true);
    assert.equal(failed.message, 'Bun install failed.');
  });
});

test('setup service installCommandLineLauncher persists status and path', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');
    let installOk = true;

    const service = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => 1,
      detectPluginInstalled: () => true,
      detectCommandLineLauncher: async () => createCommandLineLauncherSnapshot(),
      installCommandLineLauncher: async () => ({
        ok: installOk,
        installPath: installOk ? '/home/tester/.local/bin/subminer' : null,
        message: installOk ? 'Launcher installed.' : 'Launcher install failed.',
      }),
      onStateChanged: () => undefined,
    });

    const installed = await service.installCommandLineLauncher();
    assert.equal(installed.state.launcherInstallStatus, 'installed');
    assert.equal(installed.state.launcherInstallPath, '/home/tester/.local/bin/subminer');
    assert.equal(installed.canFinish, true);

    installOk = false;
    const failed = await service.installCommandLineLauncher();
    assert.equal(failed.state.launcherInstallStatus, 'failed');
    assert.equal(failed.state.launcherInstallPath, null);
    assert.equal(failed.canFinish, true);
  });
});

test('setup completion is unaffected by missing or failed command-line launcher setup', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');

    const service = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => 1,
      detectPluginInstalled: () => true,
      detectCommandLineLauncher: async () =>
        createCommandLineLauncherSnapshot({
          bun: {
            status: 'failed',
            commandPath: null,
            version: null,
            installMethod: 'official-script',
            installCommand: ['bash', '-lc', 'curl -fsSL https://bun.com/install | bash'],
            message: 'Bun install failed.',
          },
          launcher: {
            status: 'failed',
            commandPath: null,
            installPath: '/home/tester/.local/bin/subminer',
            pathDir: '/home/tester/.local/bin',
            shadowedBy: null,
            message: 'Launcher install failed.',
          },
        }),
      onStateChanged: () => undefined,
    });

    const initial = await service.ensureSetupStateInitialized();
    assert.equal(initial.canFinish, true);

    const completed = await service.markSetupCompleted();
    assert.equal(completed.state.status, 'completed');
    assert.equal(completed.canFinish, true);
  });
});

test('setup service removes legacy mpv plugin candidates and refreshes detection', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');
    let legacyCandidates = [{ path: '/tmp/mpv/scripts/subminer', kind: 'directory' as const }];

    const service = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => 1,
      detectPluginInstalled: () => legacyCandidates.length > 0,
      detectLegacyMpvPluginCandidates: () => legacyCandidates,
      removeLegacyMpvPlugins: async (candidates) => {
        assert.deepEqual(candidates, legacyCandidates);
        legacyCandidates = [];
        return {
          ok: true,
          removedPaths: ['/tmp/mpv/scripts/subminer'],
          failedPaths: [],
        };
      },
      installPlugin: async () => ({
        ok: true,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: null,
        message: 'ok',
      }),
      onStateChanged: () => undefined,
    });

    const before = await service.refreshStatus();
    assert.deepEqual(before.legacyMpvPluginPaths, ['/tmp/mpv/scripts/subminer']);

    const removed = await service.removeLegacyMpvPlugin();
    assert.equal(
      removed.message,
      'Legacy mpv plugin removed. Regular mpv will no longer load SubMiner. SubMiner-managed playback will use the bundled runtime plugin.',
    );
    assert.deepEqual(removed.legacyMpvPluginPaths, []);
  });
});

test('setup service reports failed legacy mpv plugin trash paths', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');
    const legacyCandidates = [
      { path: '/tmp/mpv/scripts/subminer', kind: 'directory' as const },
      { path: '/tmp/mpv/scripts/subminer.lua', kind: 'file' as const },
    ];

    const service = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => 1,
      detectPluginInstalled: () => true,
      detectLegacyMpvPluginCandidates: () => legacyCandidates,
      removeLegacyMpvPlugins: async () => ({
        ok: false,
        removedPaths: ['/tmp/mpv/scripts/subminer'],
        failedPaths: [{ path: '/tmp/mpv/scripts/subminer.lua', message: 'permission denied' }],
      }),
      installPlugin: async () => ({
        ok: true,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: null,
        message: 'ok',
      }),
      onStateChanged: () => undefined,
    });

    const removed = await service.removeLegacyMpvPlugin();
    assert.equal(
      removed.message,
      'Removed 1 legacy mpv plugin path, but failed to remove: /tmp/mpv/scripts/subminer.lua (permission denied). Delete the failed paths manually from mpv scripts.',
    );
    assert.deepEqual(removed.legacyMpvPluginPaths, [
      '/tmp/mpv/scripts/subminer',
      '/tmp/mpv/scripts/subminer.lua',
    ]);
  });
});
