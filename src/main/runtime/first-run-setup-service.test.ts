import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createFirstRunSetupService, shouldAutoOpenFirstRunSetup } from './first-run-setup-service';
import type { CliArgs } from '../../cli/args';

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
    help: false,
    autoStartOverlay: false,
    generateConfig: false,
    backupOverwrite: false,
    debug: false,
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

test('setup service requires mpv plugin install before finish', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');
    let dictionaryCount = 0;
    let pluginInstalled = false;

    const service = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => dictionaryCount,
      detectPluginInstalled: () => pluginInstalled,
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

    const installed = await service.installMpvPlugin();
    assert.equal(installed.state.pluginInstallStatus, 'installed');
    assert.equal(installed.pluginInstallPathSummary, '/tmp/mpv');

    pluginInstalled = true;
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

test('setup service reopens when a completed setup no longer has the mpv plugin installed', async () => {
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
    assert.equal(snapshot.state.status, 'incomplete');
    assert.equal(snapshot.canFinish, false);
    assert.equal(snapshot.pluginStatus, 'required');
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
