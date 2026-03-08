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
    stop: false,
    toggle: false,
    toggleVisibleOverlay: false,
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
    openRuntimeOptions: false,
    anilistStatus: false,
    anilistLogout: false,
    anilistSetup: false,
    anilistRetryQueue: false,
    dictionary: false,
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

test('setup service auto-completes legacy installs with config and dictionaries', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');

    const service = createFirstRunSetupService({
      configDir,
      getYomitanDictionaryCount: async () => 2,
      detectPluginInstalled: () => false,
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

test('setup service requires explicit finish for incomplete installs and supports plugin skip/install', async () => {
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

    const skipped = await service.skipPluginInstall();
    assert.equal(skipped.state.pluginInstallStatus, 'skipped');

    const installed = await service.installMpvPlugin();
    assert.equal(installed.state.pluginInstallStatus, 'installed');
    assert.equal(installed.pluginInstallPathSummary, '/tmp/mpv');

    dictionaryCount = 1;
    const refreshed = await service.refreshStatus();
    assert.equal(refreshed.canFinish, true);

    const completed = await service.markSetupCompleted();
    assert.equal(completed.state.status, 'completed');
    assert.equal(completed.state.completionSource, 'user');
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
      detectPluginInstalled: () => false,
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
