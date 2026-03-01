import test from 'node:test';
import assert from 'node:assert/strict';
import { hasExplicitCommand, parseArgs, shouldRunSettingsOnlyStartup, shouldStartApp } from './args';

test('parseArgs parses booleans and value flags', () => {
  const args = parseArgs([
    '--background',
    '--start',
    '--socket',
    '/tmp/mpv.sock',
    '--backend=hyprland',
    '--port',
    '6000',
    '--log-level',
    'warn',
    '--debug',
    '--jellyfin-play',
    '--jellyfin-server',
    'http://jellyfin.local:8096',
    '--jellyfin-item-id',
    'item-123',
    '--jellyfin-audio-stream-index',
    '2',
  ]);

  assert.equal(args.background, true);
  assert.equal(args.start, true);
  assert.equal(args.socketPath, '/tmp/mpv.sock');
  assert.equal(args.backend, 'hyprland');
  assert.equal(args.texthookerPort, 6000);
  assert.equal(args.logLevel, 'warn');
  assert.equal(args.debug, true);
  assert.equal(args.jellyfinPlay, true);
  assert.equal(args.jellyfinServer, 'http://jellyfin.local:8096');
  assert.equal(args.jellyfinItemId, 'item-123');
  assert.equal(args.jellyfinAudioStreamIndex, 2);
});

test('parseArgs ignores missing value after --log-level', () => {
  const args = parseArgs(['--log-level', '--start']);
  assert.equal(args.logLevel, undefined);
  assert.equal(args.start, true);
});

test('parseArgs handles jellyfin item listing controls', () => {
  const args = parseArgs([
    '--jellyfin-items',
    '--jellyfin-recursive=false',
    '--jellyfin-include-item-types',
    'Series,Movie,Folder',
  ]);

  assert.equal(args.jellyfinItems, true);
  assert.equal(args.jellyfinRecursive, false);
  assert.equal(args.jellyfinIncludeItemTypes, 'Series,Movie,Folder');
});

test('parseArgs handles space-separated jellyfin recursive control', () => {
  const args = parseArgs(['--jellyfin-items', '--jellyfin-recursive', 'false']);
  assert.equal(args.jellyfinRecursive, false);
});

test('hasExplicitCommand and shouldStartApp preserve command intent', () => {
  const stopOnly = parseArgs(['--stop']);
  assert.equal(hasExplicitCommand(stopOnly), true);
  assert.equal(shouldStartApp(stopOnly), false);

  const toggle = parseArgs(['--toggle-visible-overlay']);
  assert.equal(hasExplicitCommand(toggle), true);
  assert.equal(shouldStartApp(toggle), true);

  const noCommand = parseArgs(['--log-level', 'warn']);
  assert.equal(hasExplicitCommand(noCommand), false);
  assert.equal(shouldStartApp(noCommand), false);

  const refreshKnownWords = parseArgs(['--refresh-known-words']);
  assert.equal(refreshKnownWords.help, false);
  assert.equal(hasExplicitCommand(refreshKnownWords), true);
  assert.equal(shouldStartApp(refreshKnownWords), false);

  const settings = parseArgs(['--settings']);
  assert.equal(settings.settings, true);
  assert.equal(hasExplicitCommand(settings), true);
  assert.equal(shouldStartApp(settings), true);
  assert.equal(shouldRunSettingsOnlyStartup(settings), true);

  const settingsWithOverlay = parseArgs(['--settings', '--toggle-visible-overlay']);
  assert.equal(settingsWithOverlay.settings, true);
  assert.equal(settingsWithOverlay.toggleVisibleOverlay, true);
  assert.equal(shouldRunSettingsOnlyStartup(settingsWithOverlay), false);

  const yomitanAlias = parseArgs(['--yomitan']);
  assert.equal(yomitanAlias.settings, true);
  assert.equal(hasExplicitCommand(yomitanAlias), true);
  assert.equal(shouldStartApp(yomitanAlias), true);

  const help = parseArgs(['--help']);
  assert.equal(help.help, true);
  assert.equal(hasExplicitCommand(help), true);
  assert.equal(shouldStartApp(help), false);
  assert.equal(shouldRunSettingsOnlyStartup(help), false);

  const anilistStatus = parseArgs(['--anilist-status']);
  assert.equal(anilistStatus.anilistStatus, true);
  assert.equal(hasExplicitCommand(anilistStatus), true);
  assert.equal(shouldStartApp(anilistStatus), false);

  const anilistRetryQueue = parseArgs(['--anilist-retry-queue']);
  assert.equal(anilistRetryQueue.anilistRetryQueue, true);
  assert.equal(hasExplicitCommand(anilistRetryQueue), true);
  assert.equal(shouldStartApp(anilistRetryQueue), false);

  const jellyfinLibraries = parseArgs(['--jellyfin-libraries']);
  assert.equal(jellyfinLibraries.jellyfinLibraries, true);
  assert.equal(hasExplicitCommand(jellyfinLibraries), true);
  assert.equal(shouldStartApp(jellyfinLibraries), false);

  const jellyfinSetup = parseArgs(['--jellyfin']);
  assert.equal(jellyfinSetup.jellyfin, true);
  assert.equal(hasExplicitCommand(jellyfinSetup), true);
  assert.equal(shouldStartApp(jellyfinSetup), true);

  const jellyfinPlay = parseArgs(['--jellyfin-play']);
  assert.equal(jellyfinPlay.jellyfinPlay, true);
  assert.equal(hasExplicitCommand(jellyfinPlay), true);
  assert.equal(shouldStartApp(jellyfinPlay), true);

  const jellyfinSubtitles = parseArgs(['--jellyfin-subtitles', '--jellyfin-subtitle-urls']);
  assert.equal(jellyfinSubtitles.jellyfinSubtitles, true);
  assert.equal(jellyfinSubtitles.jellyfinSubtitleUrlsOnly, true);
  assert.equal(hasExplicitCommand(jellyfinSubtitles), true);
  assert.equal(shouldStartApp(jellyfinSubtitles), false);

  const jellyfinRemoteAnnounce = parseArgs(['--jellyfin-remote-announce']);
  assert.equal(jellyfinRemoteAnnounce.jellyfinRemoteAnnounce, true);
  assert.equal(hasExplicitCommand(jellyfinRemoteAnnounce), true);
  assert.equal(shouldStartApp(jellyfinRemoteAnnounce), false);

  const jellyfinPreviewAuth = parseArgs([
    '--jellyfin-preview-auth',
    '--jellyfin-response-path',
    '/tmp/subminer-jf-response.json',
  ]);
  assert.equal(jellyfinPreviewAuth.jellyfinPreviewAuth, true);
  assert.equal(
    jellyfinPreviewAuth.jellyfinResponsePath,
    '/tmp/subminer-jf-response.json',
  );
  assert.equal(hasExplicitCommand(jellyfinPreviewAuth), true);
  assert.equal(shouldStartApp(jellyfinPreviewAuth), false);

  const background = parseArgs(['--background']);
  assert.equal(background.background, true);
  assert.equal(hasExplicitCommand(background), true);
  assert.equal(shouldStartApp(background), true);
});
