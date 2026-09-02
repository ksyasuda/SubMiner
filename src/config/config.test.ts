import test from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { ConfigService, ConfigStartupParseError } from './service';
import {
  DEFAULT_CONFIG,
  DEFAULT_KEYBINDINGS,
  RUNTIME_OPTION_REGISTRY,
  deepMergeRawConfig,
} from './definitions';
import { parseConfigContent } from './parse';
import { generateConfigTemplate } from './template';
import {
  buildSubtitleCssDeclarationObject,
  getSubtitleCssManagedConfigPaths,
  getSubtitleCssPath,
  type SubtitleCssScope,
} from '../settings/subtitle-style-css';

const DEFAULT_SUBTITLE_FONT_FAMILY =
  'Hiragino Sans, M PLUS 1, Source Han Sans JP, Noto Sans CJK JP';
const DEFAULT_SECONDARY_SUBTITLE_FONT_FAMILY = DEFAULT_SUBTITLE_FONT_FAMILY;
const DEFAULT_SUBTITLE_TEXT_SHADOW =
  '-1px -1px 2px rgba(0,0,0,0.95), 1px -1px 2px rgba(0,0,0,0.95), -1px 1px 2px rgba(0,0,0,0.95), 1px 1px 2px rgba(0,0,0,0.95), 0 0 8px rgba(0,0,0,0.5)';
const SUBTITLE_CSS_SCOPES: SubtitleCssScope[] = ['primary', 'secondary', 'sidebar'];

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-config-test-'));
}

function getValueAtPath(root: unknown, path: string): unknown {
  let current = root;
  for (const segment of path.split('.')) {
    if (current === null || typeof current !== 'object' || Array.isArray(current)) {
      return undefined;
    }
    current = (current as Record<string, unknown>)[segment];
  }
  return current;
}

function buildDefaultSubtitleCssDeclarations(scope: SubtitleCssScope): Record<string, string> {
  const values: Record<string, unknown> = {
    [getSubtitleCssPath(scope)]: getValueAtPath(DEFAULT_CONFIG, getSubtitleCssPath(scope)),
  };
  for (const path of getSubtitleCssManagedConfigPaths(scope)) {
    values[path] = getValueAtPath(DEFAULT_CONFIG, path);
  }
  return buildSubtitleCssDeclarationObject(scope, values);
}

test('loads defaults when config is missing', () => {
  const dir = makeTempDir();
  const service = new ConfigService(dir);
  const config = service.getConfig();
  assert.equal(config.websocket.enabled, false);
  assert.equal(config.websocket.port, DEFAULT_CONFIG.websocket.port);
  assert.equal(config.annotationWebsocket.enabled, false);
  assert.equal(config.annotationWebsocket.port, DEFAULT_CONFIG.annotationWebsocket.port);
  assert.equal(config.texthooker.launchAtStartup, false);
  assert.equal(config.ankiConnect.behavior.autoUpdateNewCards, true);
  assert.deepEqual(config.ankiConnect.tags, ['SubMiner']);
  assert.equal(config.ankiConnect.media.audioPadding, 0);
  assert.equal(config.anilist.enabled, false);
  assert.equal(config.subtitleStyle.nameMatchImagesEnabled, false);
  assert.equal(config.anilist.characterDictionary.refreshTtlHours, 168);
  assert.equal(config.anilist.characterDictionary.maxLoaded, 3);
  assert.equal(config.anilist.characterDictionary.evictionPolicy, 'delete');
  assert.equal(config.anilist.characterDictionary.profileScope, 'all');
  assert.equal(config.anilist.characterDictionary.collapsibleSections.description, false);
  assert.equal(config.anilist.characterDictionary.collapsibleSections.characterInformation, false);
  assert.equal(config.anilist.characterDictionary.collapsibleSections.voicedBy, false);
  assert.equal(config.yomitan.externalProfilePath, '');
  assert.equal(config.jellyfin.remoteControlEnabled, true);
  assert.equal(config.jellyfin.remoteControlAutoConnect, true);
  assert.equal(config.jellyfin.autoAnnounce, false);
  assert.equal('clientName' in config.jellyfin, false);
  assert.equal('remoteControlDeviceName' in config.jellyfin, false);
  assert.equal('deviceId' in config.jellyfin, false);
  assert.equal('clientVersion' in config.jellyfin, false);
  assert.equal(config.youtube.mediaCache.mode, 'direct');
  assert.equal(config.youtube.mediaCache.maxHeight, 720);
  assert.equal(config.ai.enabled, false);
  assert.equal(config.ai.apiKeyCommand, '');
  assert.equal(config.texthooker.openBrowser, false);
  assert.equal(config.controller.enabled, false);
  assert.equal(config.ankiConnect.enabled, true);
  assert.deepEqual(config.ankiConnect.ai, {
    enabled: false,
    model: '',
    systemPrompt: '',
  });
  assert.equal(config.ankiConnect.media.normalizeAudio, true);
  assert.equal(config.ankiConnect.media.mirrorMpvVolume, true);
  assert.equal(config.startupWarmups.lowPowerMode, false);
  assert.equal(config.startupWarmups.mecab, true);
  assert.equal(config.startupWarmups.yomitanExtension, true);
  assert.equal(config.startupWarmups.subtitleDictionaries, true);
  assert.equal(config.startupWarmups.jellyfinRemoteSession, false);
  assert.equal(config.shortcuts.markAudioCard, 'CommandOrControl+Shift+A');
  assert.equal(config.shortcuts.openCharacterDictionaryManager, 'CommandOrControl+D');
  assert.equal(config.shortcuts.toggleSubtitleSidebar, 'Backslash');
  assert.equal(config.shortcuts.toggleNotificationHistory, 'CommandOrControl+N');
  assert.equal(config.discordPresence.enabled, true);
  assert.equal(config.discordPresence.updateIntervalMs, 3_000);
  assert.equal(config.subtitleStyle.backgroundColor, 'transparent');
  assert.equal(config.subtitleStyle.primaryDefaultMode, 'visible');
  assert.equal(config.subtitleStyle.preserveLineBreaks, false);
  assert.equal(config.subtitleStyle.autoPauseVideoOnHover, true);
  assert.equal(config.subtitleStyle.autoPauseVideoOnYomitanPopup, true);
  assert.equal(config.subtitleStyle.primaryVisibleOnYomitanPopup, true);
  assert.equal(config.subtitleSidebar.enabled, true);
  assert.equal(config.subtitleSidebar.pauseVideoOnHover, true);
  assert.equal(config.subtitleStyle.hoverTokenColor, '#f4dbd6');
  assert.equal(config.subtitleStyle.hoverTokenBackgroundColor, 'transparent');
  assert.equal(config.subtitleStyle.fontFamily, DEFAULT_SUBTITLE_FONT_FAMILY);
  assert.equal(config.subtitleStyle.fontWeight, '600');
  assert.equal(config.subtitleStyle.lineHeight, 1.35);
  assert.equal(config.subtitleStyle.letterSpacing, '-0.01em');
  assert.equal(config.subtitleStyle.wordSpacing, 0);
  assert.equal(config.subtitleStyle.fontKerning, 'normal');
  assert.equal(config.subtitleStyle.textRendering, 'geometricPrecision');
  assert.equal(config.subtitleStyle.textShadow, DEFAULT_SUBTITLE_TEXT_SHADOW);
  assert.equal(config.subtitleStyle.paintOrder, '');
  assert.equal(config.subtitleStyle.WebkitTextStroke, '');
  assert.equal(config.subtitleStyle.backdropFilter, 'blur(6px)');
  assert.equal(config.subtitleStyle.jlptColors.N4, '#8bd5ca');
  assert.equal(config.subtitleStyle.secondary.fontFamily, DEFAULT_SECONDARY_SUBTITLE_FONT_FAMILY);
  assert.equal(config.subtitleStyle.secondary.fontColor, '#cad3f5');
  assert.equal(config.subtitleStyle.secondary.fontWeight, '600');
  assert.equal(config.subtitleStyle.secondary.textShadow, DEFAULT_SUBTITLE_TEXT_SHADOW);
  assert.equal(config.subtitleStyle.secondary.paintOrder, '');
  assert.equal(config.subtitleStyle.secondary.WebkitTextStroke, '');
  assert.equal(config.subtitleStyle.secondary.backgroundColor, 'transparent');
  assert.deepEqual(config.subtitleSidebar.css, {});
  assert.equal(config.subtitleSidebar.fontFamily, DEFAULT_SUBTITLE_FONT_FAMILY);
  assert.equal(config.immersionTracking.enabled, true);
  assert.equal(config.immersionTracking.dbPath, '');
  assert.equal(config.immersionTracking.batchSize, 25);
  assert.equal(config.immersionTracking.flushIntervalMs, 500);
  assert.equal(config.immersionTracking.queueCap, 1000);
  assert.equal(config.immersionTracking.payloadCapBytes, 256);
  assert.equal(config.immersionTracking.maintenanceIntervalMs, 86_400_000);
  assert.equal(config.immersionTracking.retention.eventsDays, 0);
  assert.equal(config.immersionTracking.retention.telemetryDays, 0);
  assert.equal(config.immersionTracking.retention.sessionsDays, 0);
  assert.equal(config.immersionTracking.retention.dailyRollupsDays, 0);
  assert.equal(config.immersionTracking.retention.monthlyRollupsDays, 0);
  assert.equal(config.immersionTracking.retention.vacuumIntervalDays, 0);
  assert.equal(config.immersionTracking.retentionMode, 'preset');
  assert.equal(config.immersionTracking.retentionPreset, 'balanced');
  assert.equal(config.immersionTracking.lifetimeSummaries?.global, true);
  assert.equal(config.immersionTracking.lifetimeSummaries?.anime, true);
  assert.equal(config.immersionTracking.lifetimeSummaries?.media, true);
  assert.equal(config.stats.autoOpenBrowser, false);
  assert.equal(config.updates.enabled, true);
  assert.equal(config.updates.checkIntervalHours, 24);
  assert.equal(config.updates.notificationType, 'overlay');
  assert.equal(config.updates.channel, 'stable');
  assert.equal(config.mpv.socketPath, DEFAULT_CONFIG.mpv.socketPath);
  assert.equal(config.mpv.backend, 'auto');
  assert.equal(config.mpv.profile, '');
  assert.equal(config.mpv.autoStartSubMiner, true);
  assert.equal(config.mpv.pauseUntilOverlayReady, true);
  assert.equal(config.mpv.subminerBinaryPath, '');
  assert.equal(config.mpv.aniskipEnabled, true);
  assert.equal(config.mpv.aniskipButtonKey, 'TAB');
});

test('rejects invalid mpv volume mirroring values', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "media": {
          "mirrorMpvVolume": "false"
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);

  assert.equal(
    service.getConfig().ankiConnect.media.mirrorMpvVolume,
    DEFAULT_CONFIG.ankiConnect.media.mirrorMpvVolume,
  );
  assert.ok(
    service.getWarnings().some((warning) => warning.path === 'ankiConnect.media.mirrorMpvVolume'),
  );
});

test('parses updates config and warns on invalid values', () => {
  const validDir = makeTempDir();
  fs.writeFileSync(
    path.join(validDir, 'config.jsonc'),
    `{
      "updates": {
        "enabled": false,
        "checkIntervalHours": 6,
        "notificationType": "osd-system",
        "channel": "prerelease"
      }
    }`,
    'utf-8',
  );

  const validService = new ConfigService(validDir);
  assert.equal(validService.getConfig().updates.enabled, false);
  assert.equal(validService.getConfig().updates.checkIntervalHours, 6);
  assert.equal(validService.getConfig().updates.notificationType, 'osd-system');
  assert.equal(validService.getConfig().updates.channel, 'prerelease');

  const invalidDir = makeTempDir();
  fs.writeFileSync(
    path.join(invalidDir, 'config.jsonc'),
    `{
      "updates": {
        "enabled": "yes",
        "checkIntervalHours": 0,
        "notificationType": "toast",
        "channel": "nightly"
      }
    }`,
    'utf-8',
  );

  const invalidService = new ConfigService(invalidDir);
  const config = invalidService.getConfig();
  const warnings = invalidService.getWarnings();
  assert.equal(config.updates.enabled, DEFAULT_CONFIG.updates.enabled);
  assert.equal(config.updates.checkIntervalHours, DEFAULT_CONFIG.updates.checkIntervalHours);
  assert.equal(config.updates.notificationType, DEFAULT_CONFIG.updates.notificationType);
  assert.equal(config.updates.channel, DEFAULT_CONFIG.updates.channel);
  assert.ok(warnings.some((warning) => warning.path === 'updates.enabled'));
  assert.ok(warnings.some((warning) => warning.path === 'updates.checkIntervalHours'));
  assert.ok(warnings.some((warning) => warning.path === 'updates.notificationType'));
  assert.ok(warnings.some((warning) => warning.path === 'updates.channel'));
});

test('accepts overlay notification config values', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "updates": {
        "notificationType": "overlay"
      },
      "ankiConnect": {
        "behavior": {
          "notificationType": "osd-system"
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);

  assert.equal(service.getConfig().updates.notificationType, 'overlay');
  assert.equal(service.getConfig().ankiConnect.behavior.notificationType, 'osd-system');
  assert.deepEqual(service.getWarnings(), []);
});

test('parses overlay notification position config and warns on invalid values', () => {
  const validDir = makeTempDir();
  fs.writeFileSync(
    path.join(validDir, 'config.jsonc'),
    `{
      "notifications": {
        "overlayPosition": "top-left"
      }
    }`,
    'utf-8',
  );

  const validService = new ConfigService(validDir);
  assert.equal(validService.getConfig().notifications.overlayPosition, 'top-left');
  assert.deepEqual(validService.getWarnings(), []);

  const invalidDir = makeTempDir();
  fs.writeFileSync(
    path.join(invalidDir, 'config.jsonc'),
    `{
      "notifications": {
        "overlayPosition": "bottom-right"
      }
    }`,
    'utf-8',
  );

  const invalidService = new ConfigService(invalidDir);
  assert.equal(
    invalidService.getConfig().notifications.overlayPosition,
    DEFAULT_CONFIG.notifications.overlayPosition,
  );
  assert.ok(
    invalidService
      .getWarnings()
      .some((warning) => warning.path === 'notifications.overlayPosition'),
  );
});

test('throws actionable startup parse error for malformed config at construction time', () => {
  const dir = makeTempDir();
  const configPath = path.join(dir, 'config.jsonc');
  fs.writeFileSync(configPath, '{"websocket":', 'utf-8');

  assert.throws(
    () => new ConfigService(dir),
    (error: unknown) => {
      assert.ok(error instanceof ConfigStartupParseError);
      assert.equal(error.path, configPath);
      assert.ok(error.parseError.length > 0);
      assert.ok(error.message.includes(configPath));
      assert.ok(error.message.includes(error.parseError));
      return true;
    },
  );
});

test('resolves legacy subtitle appearance options without rewriting config on load', () => {
  const dir = makeTempDir();
  const configPath = path.join(dir, 'config.jsonc');
  const originalContent = `{
      "subtitleStyle": {
        "fontSize": 42,
        "fontColor": "#ffffff",
        "hoverTokenColor": "#abcdef",
        "hoverTokenBackgroundColor": "transparent",
        "css": {
          "font-size": "44px",
          "text-wrap": "balance"
        },
        "secondary": {
          "fontSize": 28,
          "fontColor": "#bbbbbb"
        }
      },
      "subtitleSidebar": {
        "fontFamily": "M PLUS 1, sans-serif",
        "fontSize": 18,
        "textColor": "#dddddd",
        "timestampColor": "#aaaaaa",
        "css": {
          "font-size": "19px"
        }
      }
    }`;
  fs.writeFileSync(configPath, originalContent, 'utf-8');

  const service = new ConfigService(dir);
  assert.equal(fs.readFileSync(configPath, 'utf-8'), originalContent);

  assert.deepEqual(service.getConfig().subtitleStyle.css, {
    color: '#ffffff',
    'font-size': '44px',
    '--subtitle-hover-token-color': '#abcdef',
    '--subtitle-hover-token-background-color': 'transparent',
    'text-wrap': 'balance',
  });
  assert.deepEqual(service.getConfig().subtitleStyle.secondary.css, {
    color: '#bbbbbb',
    'font-size': '28px',
  });
  assert.deepEqual(service.getConfig().subtitleSidebar.css, {
    'font-family': 'M PLUS 1, sans-serif',
    color: '#dddddd',
    'font-size': '19px',
    '--subtitle-sidebar-timestamp-color': '#aaaaaa',
  });
});

test('parses subtitleStyle.preserveLineBreaks and warns on invalid values', () => {
  const validDir = makeTempDir();
  fs.writeFileSync(
    path.join(validDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "preserveLineBreaks": true
      }
    }`,
    'utf-8',
  );

  const validService = new ConfigService(validDir);
  assert.equal(validService.getConfig().subtitleStyle.preserveLineBreaks, true);

  const invalidDir = makeTempDir();
  fs.writeFileSync(
    path.join(invalidDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "preserveLineBreaks": "yes"
      }
    }`,
    'utf-8',
  );

  const invalidService = new ConfigService(invalidDir);
  assert.equal(
    invalidService.getConfig().subtitleStyle.preserveLineBreaks,
    DEFAULT_CONFIG.subtitleStyle.preserveLineBreaks,
  );
  assert.ok(
    invalidService
      .getWarnings()
      .some((warning) => warning.path === 'subtitleStyle.preserveLineBreaks'),
  );
});

test('parses texthooker.launchAtStartup and warns on invalid values', () => {
  const validDir = makeTempDir();
  fs.writeFileSync(
    path.join(validDir, 'config.jsonc'),
    `{
      "texthooker": {
        "launchAtStartup": false
      }
    }`,
    'utf-8',
  );

  const validService = new ConfigService(validDir);
  assert.equal(validService.getConfig().texthooker.launchAtStartup, false);

  const invalidDir = makeTempDir();
  fs.writeFileSync(
    path.join(invalidDir, 'config.jsonc'),
    `{
      "texthooker": {
        "launchAtStartup": "yes"
      }
    }`,
    'utf-8',
  );

  const invalidService = new ConfigService(invalidDir);
  assert.equal(
    invalidService.getConfig().texthooker.launchAtStartup,
    DEFAULT_CONFIG.texthooker.launchAtStartup,
  );
  assert.ok(
    invalidService.getWarnings().some((warning) => warning.path === 'texthooker.launchAtStartup'),
  );
});

test('parses managed mpv plugin runtime settings from config', () => {
  const validDir = makeTempDir();
  fs.writeFileSync(
    path.join(validDir, 'config.jsonc'),
    `{
      "mpv": {
        "socketPath": "/tmp/custom-subminer.sock",
        "backend": "x11",
        "profile": " anime ",
        "autoStartSubMiner": false,
        "pauseUntilOverlayReady": false,
        "subminerBinaryPath": "/opt/SubMiner/SubMiner.AppImage",
        "aniskipEnabled": false,
        "aniskipButtonKey": "F8"
      }
    }`,
    'utf-8',
  );

  const validService = new ConfigService(validDir);
  const config = validService.getConfig();
  assert.equal(config.mpv.socketPath, '/tmp/custom-subminer.sock');
  assert.equal(config.mpv.backend, 'x11');
  assert.equal(config.mpv.profile, 'anime');
  assert.equal(config.mpv.autoStartSubMiner, false);
  assert.equal(config.mpv.pauseUntilOverlayReady, false);
  assert.equal(config.mpv.subminerBinaryPath, '/opt/SubMiner/SubMiner.AppImage');
  assert.equal(config.mpv.aniskipEnabled, false);
  assert.equal(config.mpv.aniskipButtonKey, 'F8');

  const invalidDir = makeTempDir();
  fs.writeFileSync(
    path.join(invalidDir, 'config.jsonc'),
    `{
      "mpv": {
        "socketPath": "",
        "backend": "weston",
        "profile": 12,
        "autoStartSubMiner": "yes",
        "pauseUntilOverlayReady": "no",
        "subminerBinaryPath": 42,
        "aniskipEnabled": "disabled",
        "aniskipButtonKey": ""
      }
    }`,
    'utf-8',
  );

  const invalidService = new ConfigService(invalidDir);
  const invalidConfig = invalidService.getConfig();
  const warnings = invalidService.getWarnings();
  assert.equal(invalidConfig.mpv.socketPath, DEFAULT_CONFIG.mpv.socketPath);
  assert.equal(invalidConfig.mpv.backend, DEFAULT_CONFIG.mpv.backend);
  assert.equal(invalidConfig.mpv.profile, DEFAULT_CONFIG.mpv.profile);
  assert.equal(invalidConfig.mpv.autoStartSubMiner, DEFAULT_CONFIG.mpv.autoStartSubMiner);
  assert.equal(invalidConfig.mpv.pauseUntilOverlayReady, DEFAULT_CONFIG.mpv.pauseUntilOverlayReady);
  assert.equal(invalidConfig.mpv.subminerBinaryPath, DEFAULT_CONFIG.mpv.subminerBinaryPath);
  assert.equal(invalidConfig.mpv.aniskipEnabled, DEFAULT_CONFIG.mpv.aniskipEnabled);
  assert.equal(invalidConfig.mpv.aniskipButtonKey, DEFAULT_CONFIG.mpv.aniskipButtonKey);
  assert.ok(warnings.some((warning) => warning.path === 'mpv.socketPath'));
  assert.ok(warnings.some((warning) => warning.path === 'mpv.backend'));
  assert.ok(warnings.some((warning) => warning.path === 'mpv.profile'));
  assert.ok(warnings.some((warning) => warning.path === 'mpv.autoStartSubMiner'));
  assert.ok(warnings.some((warning) => warning.path === 'mpv.pauseUntilOverlayReady'));
  assert.ok(warnings.some((warning) => warning.path === 'mpv.subminerBinaryPath'));
  assert.ok(warnings.some((warning) => warning.path === 'mpv.aniskipEnabled'));
  assert.ok(warnings.some((warning) => warning.path === 'mpv.aniskipButtonKey'));
});

test('parses annotationWebsocket settings and warns on invalid values', () => {
  const validDir = makeTempDir();
  fs.writeFileSync(
    path.join(validDir, 'config.jsonc'),
    `{
      "annotationWebsocket": {
        "enabled": false,
        "port": 7788
      }
    }`,
    'utf-8',
  );

  const validService = new ConfigService(validDir);
  assert.equal(validService.getConfig().annotationWebsocket.enabled, false);
  assert.equal(validService.getConfig().annotationWebsocket.port, 7788);

  const invalidDir = makeTempDir();
  fs.writeFileSync(
    path.join(invalidDir, 'config.jsonc'),
    `{
      "annotationWebsocket": {
        "enabled": "yes",
        "port": "bad"
      }
    }`,
    'utf-8',
  );

  const invalidService = new ConfigService(invalidDir);
  assert.equal(
    invalidService.getConfig().annotationWebsocket.enabled,
    DEFAULT_CONFIG.annotationWebsocket.enabled,
  );
  assert.equal(
    invalidService.getConfig().annotationWebsocket.port,
    DEFAULT_CONFIG.annotationWebsocket.port,
  );
  assert.ok(
    invalidService.getWarnings().some((warning) => warning.path === 'annotationWebsocket.enabled'),
  );
  assert.ok(
    invalidService.getWarnings().some((warning) => warning.path === 'annotationWebsocket.port'),
  );
});

test('parses subtitleStyle.autoPauseVideoOnHover and warns on invalid values', () => {
  const validDir = makeTempDir();
  fs.writeFileSync(
    path.join(validDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "autoPauseVideoOnHover": true
      }
    }`,
    'utf-8',
  );

  const validService = new ConfigService(validDir);
  assert.equal(validService.getConfig().subtitleStyle.autoPauseVideoOnHover, true);

  const invalidDir = makeTempDir();
  fs.writeFileSync(
    path.join(invalidDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "autoPauseVideoOnHover": "yes"
      }
    }`,
    'utf-8',
  );

  const invalidService = new ConfigService(invalidDir);
  assert.equal(
    invalidService.getConfig().subtitleStyle.autoPauseVideoOnHover,
    DEFAULT_CONFIG.subtitleStyle.autoPauseVideoOnHover,
  );
  assert.ok(
    invalidService
      .getWarnings()
      .some((warning) => warning.path === 'subtitleStyle.autoPauseVideoOnHover'),
  );
});

test('parses subtitleStyle.autoPauseVideoOnYomitanPopup and warns on invalid values', () => {
  const validDir = makeTempDir();
  fs.writeFileSync(
    path.join(validDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "autoPauseVideoOnYomitanPopup": true
      }
    }`,
    'utf-8',
  );

  const validService = new ConfigService(validDir);
  assert.equal(validService.getConfig().subtitleStyle.autoPauseVideoOnYomitanPopup, true);

  const invalidDir = makeTempDir();
  fs.writeFileSync(
    path.join(invalidDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "autoPauseVideoOnYomitanPopup": "yes"
      }
    }`,
    'utf-8',
  );

  const invalidService = new ConfigService(invalidDir);
  assert.equal(
    invalidService.getConfig().subtitleStyle.autoPauseVideoOnYomitanPopup,
    DEFAULT_CONFIG.subtitleStyle.autoPauseVideoOnYomitanPopup,
  );
  assert.ok(
    invalidService
      .getWarnings()
      .some((warning) => warning.path === 'subtitleStyle.autoPauseVideoOnYomitanPopup'),
  );
});

test('parses subtitleStyle.primaryVisibleOnYomitanPopup and warns on invalid values', () => {
  const validDir = makeTempDir();
  fs.writeFileSync(
    path.join(validDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "primaryVisibleOnYomitanPopup": false
      }
    }`,
    'utf-8',
  );

  const validService = new ConfigService(validDir);
  assert.equal(validService.getConfig().subtitleStyle.primaryVisibleOnYomitanPopup, false);

  const invalidDir = makeTempDir();
  fs.writeFileSync(
    path.join(invalidDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "primaryVisibleOnYomitanPopup": "yes"
      }
    }`,
    'utf-8',
  );

  const invalidService = new ConfigService(invalidDir);
  assert.equal(
    invalidService.getConfig().subtitleStyle.primaryVisibleOnYomitanPopup,
    DEFAULT_CONFIG.subtitleStyle.primaryVisibleOnYomitanPopup,
  );
  assert.ok(
    invalidService
      .getWarnings()
      .some((warning) => warning.path === 'subtitleStyle.primaryVisibleOnYomitanPopup'),
  );
});

test('parses subtitleStyle.hoverTokenColor and warns on invalid values', () => {
  const validDir = makeTempDir();
  fs.writeFileSync(
    path.join(validDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "hoverTokenColor": "#c6a0f6"
      }
    }`,
    'utf-8',
  );

  const validService = new ConfigService(validDir);
  assert.equal(validService.getConfig().subtitleStyle.hoverTokenColor, '#c6a0f6');

  const invalidDir = makeTempDir();
  fs.writeFileSync(
    path.join(invalidDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "hoverTokenColor": "purple"
      }
    }`,
    'utf-8',
  );

  const invalidService = new ConfigService(invalidDir);
  assert.equal(
    invalidService.getConfig().subtitleStyle.hoverTokenColor,
    DEFAULT_CONFIG.subtitleStyle.hoverTokenColor,
  );
  assert.ok(
    invalidService
      .getWarnings()
      .some((warning) => warning.path === 'subtitleStyle.hoverTokenColor'),
  );
});

test('parses subtitleStyle.nameMatchColor and warns on invalid values', () => {
  const validDir = makeTempDir();
  fs.writeFileSync(
    path.join(validDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "nameMatchColor": "#eed49f"
      }
    }`,
    'utf-8',
  );

  const validService = new ConfigService(validDir);
  assert.equal(
    ((validService.getConfig().subtitleStyle as unknown as Record<string, unknown>)
      .nameMatchColor ?? null) as string | null,
    '#eed49f',
  );

  const invalidDir = makeTempDir();
  fs.writeFileSync(
    path.join(invalidDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "nameMatchColor": "pink"
      }
    }`,
    'utf-8',
  );

  const invalidService = new ConfigService(invalidDir);
  assert.equal(
    ((invalidService.getConfig().subtitleStyle as unknown as Record<string, unknown>)
      .nameMatchColor ?? null) as string | null,
    '#f5bde6',
  );
  assert.ok(
    invalidService.getWarnings().some((warning) => warning.path === 'subtitleStyle.nameMatchColor'),
  );
});

test('parses subtitleStyle.hoverTokenBackgroundColor and warns on invalid values', () => {
  const validDir = makeTempDir();
  fs.writeFileSync(
    path.join(validDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "hoverTokenBackgroundColor": "#363a4fd6"
      }
    }`,
    'utf-8',
  );

  const validService = new ConfigService(validDir);
  assert.equal(validService.getConfig().subtitleStyle.hoverTokenBackgroundColor, '#363a4fd6');

  const invalidDir = makeTempDir();
  fs.writeFileSync(
    path.join(invalidDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "hoverTokenBackgroundColor": true
      }
    }`,
    'utf-8',
  );

  const invalidService = new ConfigService(invalidDir);
  assert.equal(
    invalidService.getConfig().subtitleStyle.hoverTokenBackgroundColor,
    DEFAULT_CONFIG.subtitleStyle.hoverTokenBackgroundColor,
  );
  assert.ok(
    invalidService
      .getWarnings()
      .some((warning) => warning.path === 'subtitleStyle.hoverTokenBackgroundColor'),
  );
});

test('parses subtitleStyle.hoverBackground as a hoverTokenBackgroundColor alias', () => {
  const validDir = makeTempDir();
  fs.writeFileSync(
    path.join(validDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "hoverBackground": "transparent"
      }
    }`,
    'utf-8',
  );

  const validService = new ConfigService(validDir);
  assert.equal(validService.getConfig().subtitleStyle.hoverTokenBackgroundColor, 'transparent');
});

test('parses subtitleStyle.hoverTokenBackgroundColor null as invalid instead of missing', () => {
  const invalidDir = makeTempDir();
  fs.writeFileSync(
    path.join(invalidDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "hoverTokenBackgroundColor": null
      }
    }`,
    'utf-8',
  );

  const invalidService = new ConfigService(invalidDir);
  assert.equal(
    invalidService.getConfig().subtitleStyle.hoverTokenBackgroundColor,
    DEFAULT_CONFIG.subtitleStyle.hoverTokenBackgroundColor,
  );
  assert.ok(
    invalidService
      .getWarnings()
      .some((warning) => warning.path === 'subtitleStyle.hoverTokenBackgroundColor'),
  );
});

test('parses subtitleStyle.nameMatchEnabled and warns on invalid values', () => {
  const validDir = makeTempDir();
  fs.writeFileSync(
    path.join(validDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "nameMatchEnabled": false
      }
    }`,
    'utf-8',
  );

  const validService = new ConfigService(validDir);
  assert.equal(validService.getConfig().subtitleStyle.nameMatchEnabled, false);

  const invalidDir = makeTempDir();
  fs.writeFileSync(
    path.join(invalidDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "nameMatchEnabled": "no"
      }
    }`,
    'utf-8',
  );

  const invalidService = new ConfigService(invalidDir);
  assert.equal(
    invalidService.getConfig().subtitleStyle.nameMatchEnabled,
    DEFAULT_CONFIG.subtitleStyle.nameMatchEnabled,
  );
  assert.ok(
    invalidService
      .getWarnings()
      .some((warning) => warning.path === 'subtitleStyle.nameMatchEnabled'),
  );
});

test('parses subtitleStyle.nameMatchImagesEnabled and warns on invalid values', () => {
  const validDir = makeTempDir();
  fs.writeFileSync(
    path.join(validDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "nameMatchImagesEnabled": true
      }
    }`,
    'utf-8',
  );

  const validService = new ConfigService(validDir);
  assert.equal(validService.getConfig().subtitleStyle.nameMatchImagesEnabled, true);

  const invalidDir = makeTempDir();
  fs.writeFileSync(
    path.join(invalidDir, 'config.jsonc'),
    `{
      "subtitleStyle": {
        "nameMatchImagesEnabled": "yes"
      }
    }`,
    'utf-8',
  );

  const invalidService = new ConfigService(invalidDir);
  assert.equal(
    invalidService.getConfig().subtitleStyle.nameMatchImagesEnabled,
    DEFAULT_CONFIG.subtitleStyle.nameMatchImagesEnabled,
  );
  assert.ok(
    invalidService
      .getWarnings()
      .some((warning) => warning.path === 'subtitleStyle.nameMatchImagesEnabled'),
  );
});

test('parses anilist.enabled and warns for invalid value', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "anilist": {
        "enabled": "yes"
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(config.anilist.enabled, DEFAULT_CONFIG.anilist.enabled);
  assert.ok(warnings.some((warning) => warning.path === 'anilist.enabled'));

  service.patchRawConfig({ anilist: { enabled: true } });
  assert.equal(service.getConfig().anilist.enabled, true);
});

test('parses anilist.characterDictionary config with clamping and enum validation', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "anilist": {
        "characterDictionary": {
          "enabled": true,
          "refreshTtlHours": 0,
          "maxLoaded": 1000,
          "evictionPolicy": "remove",
          "profileScope": "everywhere"
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(config.anilist.characterDictionary.refreshTtlHours, 1);
  assert.equal(config.anilist.characterDictionary.maxLoaded, 20);
  assert.equal(config.anilist.characterDictionary.evictionPolicy, 'delete');
  assert.equal(config.anilist.characterDictionary.profileScope, 'all');
  assert.ok(
    warnings.some((warning) => warning.path === 'anilist.characterDictionary.refreshTtlHours'),
  );
  assert.ok(warnings.some((warning) => warning.path === 'anilist.characterDictionary.maxLoaded'));
  assert.ok(
    warnings.some((warning) => warning.path === 'anilist.characterDictionary.evictionPolicy'),
  );
  assert.ok(
    warnings.some((warning) => warning.path === 'anilist.characterDictionary.profileScope'),
  );
});

test('parses anilist.characterDictionary.collapsibleSections booleans and warns on invalid values', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "anilist": {
        "characterDictionary": {
          "collapsibleSections": {
            "description": true,
            "characterInformation": "yes",
            "voicedBy": true
          }
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(config.anilist.characterDictionary.collapsibleSections.description, true);
  assert.equal(config.anilist.characterDictionary.collapsibleSections.characterInformation, false);
  assert.equal(config.anilist.characterDictionary.collapsibleSections.voicedBy, true);
  assert.ok(
    warnings.some(
      (warning) =>
        warning.path === 'anilist.characterDictionary.collapsibleSections.characterInformation',
    ),
  );
});

test('parses jellyfin remote control fields and ignores legacy identity fields', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "jellyfin": {
        "enabled": true,
        "serverUrl": "http://127.0.0.1:8096",
        "remoteControlEnabled": true,
        "remoteControlAutoConnect": true,
        "autoAnnounce": true,
        "clientName": "Custom Client",
        "remoteControlDeviceName": "SubMiner"
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.jellyfin.enabled, true);
  assert.equal(config.jellyfin.serverUrl, 'http://127.0.0.1:8096');
  assert.equal(config.jellyfin.remoteControlEnabled, true);
  assert.equal(config.jellyfin.remoteControlAutoConnect, true);
  assert.equal(config.jellyfin.autoAnnounce, true);
  assert.equal('clientName' in config.jellyfin, false);
  assert.equal('remoteControlDeviceName' in config.jellyfin, false);
});

test('parses jellyfin.enabled and remoteControlEnabled disabled combinations', () => {
  const disabledDir = makeTempDir();
  fs.writeFileSync(
    path.join(disabledDir, 'config.jsonc'),
    `{
      "jellyfin": {
        "enabled": false,
        "remoteControlEnabled": false
      }
    }`,
    'utf-8',
  );

  const disabledService = new ConfigService(disabledDir);
  const disabledConfig = disabledService.getConfig();
  assert.equal(disabledConfig.jellyfin.enabled, false);
  assert.equal(disabledConfig.jellyfin.remoteControlEnabled, false);
  assert.equal(
    disabledService
      .getWarnings()
      .some(
        (warning) =>
          warning.path === 'jellyfin.enabled' || warning.path === 'jellyfin.remoteControlEnabled',
      ),
    false,
  );

  const mixedDir = makeTempDir();
  fs.writeFileSync(
    path.join(mixedDir, 'config.jsonc'),
    `{
      "jellyfin": {
        "enabled": true,
        "remoteControlEnabled": false
      }
    }`,
    'utf-8',
  );

  const mixedService = new ConfigService(mixedDir);
  const mixedConfig = mixedService.getConfig();
  assert.equal(mixedConfig.jellyfin.enabled, true);
  assert.equal(mixedConfig.jellyfin.remoteControlEnabled, false);
  assert.equal(
    mixedService
      .getWarnings()
      .some(
        (warning) =>
          warning.path === 'jellyfin.enabled' || warning.path === 'jellyfin.remoteControlEnabled',
      ),
    false,
  );
});

test('parses startup warmup toggles and low-power mode', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "startupWarmups": {
        "lowPowerMode": true,
        "mecab": false,
        "yomitanExtension": true,
        "subtitleDictionaries": false,
        "jellyfinRemoteSession": false
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  assert.equal(config.startupWarmups.lowPowerMode, true);
  assert.equal(config.startupWarmups.mecab, false);
  assert.equal(config.startupWarmups.yomitanExtension, true);
  assert.equal(config.startupWarmups.subtitleDictionaries, false);
  assert.equal(config.startupWarmups.jellyfinRemoteSession, false);
});

test('invalid startup warmup values warn and keep defaults', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "startupWarmups": {
        "lowPowerMode": "yes",
        "mecab": 1,
        "yomitanExtension": null,
        "subtitleDictionaries": "no",
        "jellyfinRemoteSession": []
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(config.startupWarmups.lowPowerMode, DEFAULT_CONFIG.startupWarmups.lowPowerMode);
  assert.equal(config.startupWarmups.mecab, DEFAULT_CONFIG.startupWarmups.mecab);
  assert.equal(
    config.startupWarmups.yomitanExtension,
    DEFAULT_CONFIG.startupWarmups.yomitanExtension,
  );
  assert.equal(
    config.startupWarmups.subtitleDictionaries,
    DEFAULT_CONFIG.startupWarmups.subtitleDictionaries,
  );
  assert.equal(
    config.startupWarmups.jellyfinRemoteSession,
    DEFAULT_CONFIG.startupWarmups.jellyfinRemoteSession,
  );
  assert.ok(warnings.some((warning) => warning.path === 'startupWarmups.lowPowerMode'));
  assert.ok(warnings.some((warning) => warning.path === 'startupWarmups.mecab'));
  assert.ok(warnings.some((warning) => warning.path === 'startupWarmups.yomitanExtension'));
  assert.ok(warnings.some((warning) => warning.path === 'startupWarmups.subtitleDictionaries'));
  assert.ok(warnings.some((warning) => warning.path === 'startupWarmups.jellyfinRemoteSession'));
});

test('parses discordPresence fields and warns for invalid types', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "discordPresence": {
        "enabled": true,
        "updateIntervalMs": 3000,
        "debounceMs": 250
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  assert.equal(config.discordPresence.enabled, true);
  assert.equal(config.discordPresence.updateIntervalMs, 3000);
  assert.equal(config.discordPresence.debounceMs, 250);

  service.patchRawConfig({ discordPresence: { enabled: 'yes' as never } });
  assert.equal(service.getConfig().discordPresence.enabled, DEFAULT_CONFIG.discordPresence.enabled);
  assert.ok(service.getWarnings().some((warning) => warning.path === 'discordPresence.enabled'));
});

test('accepts immersion tracking config values', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "immersionTracking": {
        "enabled": false,
        "dbPath": "/tmp/immersions/custom.sqlite",
        "batchSize": 50,
        "flushIntervalMs": 750,
        "queueCap": 2000,
        "payloadCapBytes": 512,
        "maintenanceIntervalMs": 3600000,
        "retentionMode": "preset",
        "retentionPreset": "minimal",
        "retention": {
          "eventsDays": 14,
          "telemetryDays": 45,
          "sessionsDays": 60,
          "dailyRollupsDays": 730,
          "monthlyRollupsDays": 3650,
          "vacuumIntervalDays": 14
        },
        "lifetimeSummaries": {
          "global": false,
          "anime": true,
          "media": false
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.immersionTracking.enabled, false);
  assert.equal(config.immersionTracking.dbPath, '/tmp/immersions/custom.sqlite');
  assert.equal(config.immersionTracking.batchSize, 50);
  assert.equal(config.immersionTracking.flushIntervalMs, 750);
  assert.equal(config.immersionTracking.queueCap, 2000);
  assert.equal(config.immersionTracking.payloadCapBytes, 512);
  assert.equal(config.immersionTracking.maintenanceIntervalMs, 3_600_000);
  assert.equal(config.immersionTracking.retention.eventsDays, 14);
  assert.equal(config.immersionTracking.retention.telemetryDays, 45);
  assert.equal(config.immersionTracking.retention.sessionsDays, 60);
  assert.equal(config.immersionTracking.retention.dailyRollupsDays, 730);
  assert.equal(config.immersionTracking.retention.monthlyRollupsDays, 3650);
  assert.equal(config.immersionTracking.retention.vacuumIntervalDays, 14);
  assert.equal(config.immersionTracking.retentionMode, 'preset');
  assert.equal(config.immersionTracking.retentionPreset, 'minimal');
  assert.equal(config.immersionTracking.lifetimeSummaries?.global, false);
  assert.equal(config.immersionTracking.lifetimeSummaries?.anime, true);
  assert.equal(config.immersionTracking.lifetimeSummaries?.media, false);
});

test('falls back for invalid immersion tracking tuning values', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "immersionTracking": {
        "retentionMode": "bad",
        "retentionPreset": "bad",
        "batchSize": 0,
        "flushIntervalMs": 1,
        "queueCap": 5,
        "payloadCapBytes": 16,
        "maintenanceIntervalMs": 1000,
        "retention": {
          "eventsDays": -1,
          "telemetryDays": 99999,
          "sessionsDays": -1,
          "dailyRollupsDays": -1,
          "monthlyRollupsDays": 999999,
          "vacuumIntervalDays": -1
        },
        "lifetimeSummaries": "bad"
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(config.immersionTracking.batchSize, 25);
  assert.equal(config.immersionTracking.flushIntervalMs, 500);
  assert.equal(config.immersionTracking.queueCap, 1000);
  assert.equal(config.immersionTracking.payloadCapBytes, 256);
  assert.equal(config.immersionTracking.maintenanceIntervalMs, 86_400_000);
  assert.equal(config.immersionTracking.retention.eventsDays, 0);
  assert.equal(config.immersionTracking.retention.telemetryDays, 0);
  assert.equal(config.immersionTracking.retention.sessionsDays, 0);
  assert.equal(config.immersionTracking.retention.dailyRollupsDays, 0);
  assert.equal(config.immersionTracking.retention.monthlyRollupsDays, 0);
  assert.equal(config.immersionTracking.retention.vacuumIntervalDays, 0);
  assert.equal(config.immersionTracking.retentionMode, 'preset');
  assert.equal(config.immersionTracking.retentionPreset, 'balanced');
  assert.equal(config.immersionTracking.lifetimeSummaries?.global, true);
  assert.equal(config.immersionTracking.lifetimeSummaries?.anime, true);
  assert.equal(config.immersionTracking.lifetimeSummaries?.media, true);

  assert.ok(warnings.some((warning) => warning.path === 'immersionTracking.batchSize'));
  assert.ok(warnings.some((warning) => warning.path === 'immersionTracking.flushIntervalMs'));
  assert.ok(warnings.some((warning) => warning.path === 'immersionTracking.queueCap'));
  assert.ok(warnings.some((warning) => warning.path === 'immersionTracking.payloadCapBytes'));
  assert.ok(warnings.some((warning) => warning.path === 'immersionTracking.maintenanceIntervalMs'));
  assert.ok(warnings.some((warning) => warning.path === 'immersionTracking.retention.eventsDays'));
  assert.ok(
    warnings.some((warning) => warning.path === 'immersionTracking.retention.telemetryDays'),
  );
  assert.ok(
    warnings.some((warning) => warning.path === 'immersionTracking.retention.sessionsDays'),
  );
  assert.ok(
    warnings.some((warning) => warning.path === 'immersionTracking.retention.dailyRollupsDays'),
  );
  assert.ok(
    warnings.some((warning) => warning.path === 'immersionTracking.retention.monthlyRollupsDays'),
  );
  assert.ok(
    warnings.some((warning) => warning.path === 'immersionTracking.retention.vacuumIntervalDays'),
  );
  assert.ok(warnings.some((warning) => warning.path === 'immersionTracking.retentionMode'));
  assert.ok(warnings.some((warning) => warning.path === 'immersionTracking.retentionPreset'));
  assert.ok(warnings.some((warning) => warning.path === 'immersionTracking.lifetimeSummaries'));
});

test('applies retention presets and explicit overrides', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "immersionTracking": {
        "retentionMode": "preset",
        "retentionPreset": "minimal",
        "retention": {
          "eventsDays": 11,
          "sessionsDays": 8
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.immersionTracking.retentionMode, 'preset');
  assert.equal(config.immersionTracking.retentionPreset, 'minimal');
  assert.equal(config.immersionTracking.retention.eventsDays, 11);
  assert.equal(config.immersionTracking.retention.sessionsDays, 8);
  assert.equal(config.immersionTracking.retention.telemetryDays, 14);
  assert.equal(config.immersionTracking.retention.dailyRollupsDays, 30);
});

test('parses jsonc and warns/falls back on invalid value', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      // invalid websocket port
      "websocket": { "port": "bad" }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  assert.equal(config.websocket.port, DEFAULT_CONFIG.websocket.port);
  assert.ok(service.getWarnings().some((w) => w.path === 'websocket.port'));
});

test('accepts trailing commas in jsonc', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "websocket": {
        "enabled": "auto",
        "port": 7788,
      },
      "youtube": {
        "primarySubLanguages": ["ja", "en",],
      },
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  assert.equal(config.websocket.port, 7788);
  assert.deepEqual(config.youtube.primarySubLanguages, ['ja', 'en']);
});

test('reloadConfigStrict rejects invalid jsonc and preserves previous config', () => {
  const dir = makeTempDir();
  const configPath = path.join(dir, 'config.jsonc');
  fs.writeFileSync(
    configPath,
    `{
    "logging": {
      "level": "warn"
    }
  }`,
  );

  const service = new ConfigService(dir);
  assert.equal(service.getConfig().logging.level, 'warn');

  fs.writeFileSync(
    configPath,
    `{
    "logging":`,
  );

  const result = service.reloadConfigStrict();
  assert.equal(result.ok, false);
  if (result.ok) {
    throw new Error('Expected strict reload to fail on invalid JSONC.');
  }
  assert.equal(result.path, configPath);
  assert.equal(service.getConfig().logging.level, 'warn');
});

test('reloadConfigStrict rejects invalid json and preserves previous config', () => {
  const dir = makeTempDir();
  const configPath = path.join(dir, 'config.json');
  fs.writeFileSync(configPath, JSON.stringify({ logging: { level: 'error' } }, null, 2));

  const service = new ConfigService(dir);
  assert.equal(service.getConfig().logging.level, 'error');

  fs.writeFileSync(configPath, '{"logging":');

  const result = service.reloadConfigStrict();
  assert.equal(result.ok, false);
  if (result.ok) {
    throw new Error('Expected strict reload to fail on invalid JSON.');
  }
  assert.equal(result.path, configPath);
  assert.equal(service.getConfig().logging.level, 'error');
});

test('prefers config.jsonc over config.json when both exist', () => {
  const dir = makeTempDir();
  const jsonPath = path.join(dir, 'config.json');
  const jsoncPath = path.join(dir, 'config.jsonc');
  fs.writeFileSync(jsonPath, JSON.stringify({ logging: { level: 'error' } }, null, 2));
  fs.writeFileSync(
    jsoncPath,
    `{
      "logging": {
        "level": "warn"
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  assert.equal(service.getConfig().logging.level, 'warn');
  assert.equal(service.getConfigPath(), jsoncPath);
});

test('reloadConfigStrict parse failure does not mutate raw config or warnings', () => {
  const dir = makeTempDir();
  const configPath = path.join(dir, 'config.jsonc');
  fs.writeFileSync(
    configPath,
    `{
      "logging": {
        "level": "warn"
      },
      "websocket": {
        "port": "bad"
      }
    }`,
  );

  const service = new ConfigService(dir);
  const beforePath = service.getConfigPath();
  const beforeConfig = service.getConfig();
  const beforeRaw = service.getRawConfig();
  const beforeWarnings = service.getWarnings();

  fs.writeFileSync(configPath, '{"logging":');

  const result = service.reloadConfigStrict();
  assert.equal(result.ok, false);
  assert.equal(service.getConfigPath(), beforePath);
  assert.deepEqual(service.getConfig(), beforeConfig);
  assert.deepEqual(service.getRawConfig(), beforeRaw);
  assert.deepEqual(service.getWarnings(), beforeWarnings);
});

test('SM-012 config paths do not use JSON serialize-clone helpers', () => {
  const definitionsSource = fs.readFileSync(
    path.join(process.cwd(), 'src/config/definitions.ts'),
    'utf-8',
  );
  const serviceSource = fs.readFileSync(path.join(process.cwd(), 'src/config/service.ts'), 'utf-8');

  assert.equal(definitionsSource.includes('JSON.parse(JSON.stringify('), false);
  assert.equal(serviceSource.includes('JSON.parse(JSON.stringify('), false);
});

test('getRawConfig returns a detached clone', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "tags": ["SubMiner"]
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const raw = service.getRawConfig();
  raw.ankiConnect!.tags!.push('mutated');

  assert.deepEqual(service.getRawConfig().ankiConnect?.tags, ['SubMiner']);
});

test('deepMergeRawConfig returns a detached merged clone', () => {
  const base = {
    ankiConnect: {
      tags: ['SubMiner'],
      behavior: {
        autoUpdateNewCards: true,
      },
    },
  };

  const merged = deepMergeRawConfig(base, {
    ankiConnect: {
      behavior: {
        autoUpdateNewCards: false,
      },
    },
  });

  merged.ankiConnect!.tags!.push('mutated');
  merged.ankiConnect!.behavior!.autoUpdateNewCards = true;

  assert.deepEqual(base.ankiConnect?.tags, ['SubMiner']);
  assert.equal(base.ankiConnect?.behavior?.autoUpdateNewCards, true);
});

test('warning emission order is deterministic across reloads', () => {
  const dir = makeTempDir();
  const configPath = path.join(dir, 'config.jsonc');
  fs.writeFileSync(
    configPath,
    `{
      "unknownFeature": true,
      "websocket": {
        "enabled": "sometimes",
        "port": -1
      },
      "annotationWebsocket": {
        "enabled": "sometimes",
        "port": -1
      },
      "logging": {
        "level": "trace"
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const firstWarnings = service.getWarnings();

  service.reloadConfig();
  const secondWarnings = service.getWarnings();

  assert.deepEqual(secondWarnings, firstWarnings);
  assert.deepEqual(
    firstWarnings.map((warning) => warning.path),
    [
      'unknownFeature',
      'websocket.enabled',
      'websocket.port',
      'annotationWebsocket.enabled',
      'annotationWebsocket.port',
      'logging.level',
    ],
  );
});

test('accepts valid logging.level', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "logging": {
        "level": "warn"
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.logging.level, 'warn');
});

test('accepts valid logging.rotation', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "logging": {
        "rotation": 14
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.logging.rotation, 14);
});

test('accepts valid logging file toggles', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "logging": {
        "files": {
          "app": false,
          "launcher": true,
          "mpv": true
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.deepEqual(config.logging.files, {
    app: false,
    launcher: true,
    mpv: true,
  });
});

test('falls back for invalid logging.level and reports warning', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "logging": {
        "level": "trace"
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(config.logging.level, DEFAULT_CONFIG.logging.level);
  assert.ok(warnings.some((warning) => warning.path === 'logging.level'));
});

test('falls back for invalid logging.rotation and reports warning', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "logging": {
        "rotation": 0
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(config.logging.rotation, DEFAULT_CONFIG.logging.rotation);
  assert.ok(warnings.some((warning) => warning.path === 'logging.rotation'));
});

test('falls back for invalid logging file toggles and reports warning', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "logging": {
        "files": {
          "mpv": "yes"
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(config.logging.files.mpv, DEFAULT_CONFIG.logging.files.mpv);
  assert.ok(warnings.some((warning) => warning.path === 'logging.files.mpv'));
});

test('falls back for invalid logging files object and reports warning', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "logging": {
        "files": false
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.deepEqual(config.logging.files, DEFAULT_CONFIG.logging.files);
  assert.ok(warnings.some((warning) => warning.path === 'logging.files'));
});

test('warns and ignores unknown top-level config keys', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "websocket": {
        "port": 7788
      },
      "unknownFeatureFlag": {
        "enabled": true
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(config.websocket.port, 7788);
  assert.ok(warnings.some((warning) => warning.path === 'unknownFeatureFlag'));
});

test('parses global shortcuts and startup settings', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ai": {
        "enabled": true,
        "apiKeyCommand": "pass show subminer/ai",
        "model": "openai/gpt-4o-mini"
      },
      "shortcuts": {
        "toggleVisibleOverlayGlobal": "Alt+Shift+U",
        "openJimaku": "Ctrl+Alt+J"
      },
      "youtube": {
        "primarySubLanguages": ["ja", "jpn", "jp"]
      },
      "youtubeSubgen": {
        "whisperVadModel": "/models/vad.bin",
        "whisperThreads": 12,
        "fixWithAi": true
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  assert.equal(config.ai.enabled, true);
  assert.equal(config.ai.apiKeyCommand, 'pass show subminer/ai');
  assert.equal(config.shortcuts.toggleVisibleOverlayGlobal, 'Alt+Shift+U');
  assert.equal(config.shortcuts.openJimaku, 'Ctrl+Alt+J');
  assert.deepEqual(config.youtube.primarySubLanguages, ['ja', 'jpn', 'jp']);
  assert.equal(config.youtubeSubgen.whisperVadModel, '/models/vad.bin');
  assert.equal(config.youtubeSubgen.whisperThreads, 12);
  assert.equal(config.youtubeSubgen.fixWithAi, true);
});

test('parses YouTube media cache config and warns on invalid values', () => {
  const validDir = makeTempDir();
  fs.writeFileSync(
    path.join(validDir, 'config.jsonc'),
    `{
      "youtube": {
        "mediaCache": {
          "mode": "background",
          "maxHeight": 480
        }
      }
    }`,
    'utf-8',
  );

  const validService = new ConfigService(validDir);
  assert.equal(validService.getConfig().youtube.mediaCache.mode, 'background');
  assert.equal(validService.getConfig().youtube.mediaCache.maxHeight, 480);

  const invalidDir = makeTempDir();
  fs.writeFileSync(
    path.join(invalidDir, 'config.jsonc'),
    `{
      "youtube": {
        "mediaCache": {
          "mode": "always",
          "maxHeight": -1
        }
      }
    }`,
    'utf-8',
  );

  const invalidService = new ConfigService(invalidDir);
  assert.equal(
    invalidService.getConfig().youtube.mediaCache.mode,
    DEFAULT_CONFIG.youtube.mediaCache.mode,
  );
  assert.equal(
    invalidService.getConfig().youtube.mediaCache.maxHeight,
    DEFAULT_CONFIG.youtube.mediaCache.maxHeight,
  );
  assert.ok(
    invalidService.getWarnings().some((warning) => warning.path === 'youtube.mediaCache.mode'),
  );
  assert.ok(
    invalidService.getWarnings().some((warning) => warning.path === 'youtube.mediaCache.maxHeight'),
  );
});

test('parses controller settings with logical bindings and tuning knobs', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "controller": {
        "enabled": true,
        "preferredGamepadId": "Xbox Wireless Controller (STANDARD GAMEPAD Vendor: 045e Product: 0b13)",
        "preferredGamepadLabel": "Xbox Wireless Controller",
        "smoothScroll": false,
        "scrollPixelsPerSecond": 1440,
        "horizontalJumpPixels": 180,
        "stickDeadzone": 0.3,
        "triggerInputMode": "analog",
        "triggerDeadzone": 0.4,
        "repeatDelayMs": 220,
        "repeatIntervalMs": 70,
        "buttonIndices": {
          "select": 6,
          "leftStickPress": 9,
          "rightStickPress": 10
        },
        "bindings": {
          "toggleLookup": "buttonWest",
          "closeLookup": "buttonEast",
          "toggleKeyboardOnlyMode": "buttonNorth",
          "mineCard": "buttonSouth",
          "quitMpv": "select",
          "previousAudio": "leftShoulder",
          "nextAudio": "rightShoulder",
          "playCurrentAudio": "none",
          "toggleMpvPause": "leftStickPress",
          "leftStickHorizontal": "rightStickX",
          "leftStickVertical": "rightStickY",
          "rightStickHorizontal": "leftStickX",
          "rightStickVertical": "leftStickY"
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.controller.enabled, true);
  assert.equal(
    config.controller.preferredGamepadId,
    'Xbox Wireless Controller (STANDARD GAMEPAD Vendor: 045e Product: 0b13)',
  );
  assert.equal(config.controller.preferredGamepadLabel, 'Xbox Wireless Controller');
  assert.equal(config.controller.smoothScroll, false);
  assert.equal(config.controller.scrollPixelsPerSecond, 1440);
  assert.equal(config.controller.horizontalJumpPixels, 180);
  assert.equal(config.controller.stickDeadzone, 0.3);
  assert.equal(config.controller.triggerInputMode, 'analog');
  assert.equal(config.controller.triggerDeadzone, 0.4);
  assert.equal(config.controller.repeatDelayMs, 220);
  assert.equal(config.controller.repeatIntervalMs, 70);
  assert.equal(config.controller.buttonIndices.select, 6);
  assert.equal(config.controller.buttonIndices.leftStickPress, 9);
  assert.deepEqual(config.controller.bindings.toggleLookup, { kind: 'button', buttonIndex: 2 });
  assert.deepEqual(config.controller.bindings.quitMpv, { kind: 'button', buttonIndex: 6 });
  assert.deepEqual(config.controller.bindings.playCurrentAudio, { kind: 'none' });
  assert.deepEqual(config.controller.bindings.toggleMpvPause, { kind: 'button', buttonIndex: 9 });
  assert.deepEqual(config.controller.bindings.leftStickHorizontal, {
    kind: 'axis',
    axisIndex: 3,
    dpadFallback: 'horizontal',
  });
  assert.deepEqual(config.controller.bindings.rightStickVertical, {
    kind: 'axis',
    axisIndex: 1,
    dpadFallback: 'none',
  });
});

test('parses descriptor-based controller bindings', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "controller": {
        "bindings": {
          "toggleLookup": { "kind": "button", "buttonIndex": 11 },
          "closeLookup": { "kind": "axis", "axisIndex": 4, "direction": "negative" },
          "playCurrentAudio": { "kind": "none" },
          "leftStickHorizontal": { "kind": "axis", "axisIndex": 7, "dpadFallback": "none" },
          "leftStickVertical": { "kind": "axis", "axisIndex": 2, "dpadFallback": "vertical" }
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.deepEqual(config.controller.bindings.toggleLookup, {
    kind: 'button',
    buttonIndex: 11,
  });
  assert.deepEqual(config.controller.bindings.closeLookup, {
    kind: 'axis',
    axisIndex: 4,
    direction: 'negative',
  });
  assert.deepEqual(config.controller.bindings.playCurrentAudio, { kind: 'none' });
  assert.deepEqual(config.controller.bindings.leftStickHorizontal, {
    kind: 'axis',
    axisIndex: 7,
    dpadFallback: 'none',
  });
  assert.deepEqual(config.controller.bindings.leftStickVertical, {
    kind: 'axis',
    axisIndex: 2,
    dpadFallback: 'vertical',
  });
});

test('parses controller profiles as per-gamepad binding overrides', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "controller": {
        "buttonIndices": {
          "buttonSouth": 0,
          "leftTrigger": 6
        },
        "bindings": {
          "toggleLookup": { "kind": "button", "buttonIndex": 0 },
          "quitMpv": "leftTrigger"
        },
        "profiles": {
          "8BitDo SN30": {
            "label": "8BitDo SN30",
            "bindings": {
              "toggleLookup": { "kind": "button", "buttonIndex": 11 },
              "leftStickVertical": { "kind": "axis", "axisIndex": 7, "dpadFallback": "none" }
            }
          },
          "Xbox Wireless Controller": {
            "buttonIndices": {
              "leftTrigger": 8
            },
            "bindings": {
              "quitMpv": "leftTrigger"
            }
          }
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.deepEqual(config.controller.profiles['8BitDo SN30']?.bindings.toggleLookup, {
    kind: 'button',
    buttonIndex: 11,
  });
  assert.deepEqual(config.controller.profiles['8BitDo SN30']?.bindings.closeLookup, {
    kind: 'button',
    buttonIndex: 1,
  });
  assert.deepEqual(config.controller.profiles['8BitDo SN30']?.bindings.leftStickVertical, {
    kind: 'axis',
    axisIndex: 7,
    dpadFallback: 'none',
  });
  assert.deepEqual(config.controller.profiles['Xbox Wireless Controller']?.bindings.quitMpv, {
    kind: 'button',
    buttonIndex: 8,
  });
  assert.equal(
    config.controller.profiles['Xbox Wireless Controller']?.buttonIndices.leftTrigger,
    8,
  );
  assert.deepEqual(config.controller.bindings.quitMpv, { kind: 'button', buttonIndex: 6 });
});

test('rejects reserved controller profile ids from config', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "controller": {
        "profiles": {
          "__proto__": { "label": "polluted" },
          "constructor": { "label": "ctor" },
          "prototype": { "label": "proto" },
          "pad-1": { "label": "Pad 1" }
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(Object.hasOwn(config.controller.profiles, '__proto__'), false);
  assert.equal(Object.hasOwn(config.controller.profiles, 'constructor'), false);
  assert.equal(Object.hasOwn(config.controller.profiles, 'prototype'), false);
  assert.equal(config.controller.profiles['pad-1']?.label, 'Pad 1');
  assert.equal(
    warnings.some((warning) => warning.path === 'controller.profiles.constructor'),
    true,
  );
  assert.equal(
    warnings.some((warning) => warning.path === 'controller.profiles.prototype'),
    true,
  );
});

test('controller descriptor config rejects malformed binding objects', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "controller": {
        "bindings": {
          "toggleLookup": { "kind": "button", "buttonIndex": -1 },
          "closeLookup": { "kind": "axis", "axisIndex": 1, "direction": "sideways" },
          "leftStickHorizontal": { "kind": "axis", "axisIndex": 0, "dpadFallback": "diagonal" }
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.deepEqual(
    config.controller.bindings.toggleLookup,
    DEFAULT_CONFIG.controller.bindings.toggleLookup,
  );
  assert.deepEqual(
    config.controller.bindings.closeLookup,
    DEFAULT_CONFIG.controller.bindings.closeLookup,
  );
  assert.deepEqual(
    config.controller.bindings.leftStickHorizontal,
    DEFAULT_CONFIG.controller.bindings.leftStickHorizontal,
  );
  assert.equal(
    warnings.some((warning) => warning.path === 'controller.bindings.toggleLookup'),
    true,
  );
  assert.equal(
    warnings.some((warning) => warning.path === 'controller.bindings.closeLookup'),
    true,
  );
  assert.equal(
    warnings.some((warning) => warning.path === 'controller.bindings.leftStickHorizontal'),
    true,
  );
});

test('controller positive-number tuning rejects sub-unit values that floor to zero', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "controller": {
        "scrollPixelsPerSecond": 0.5,
        "horizontalJumpPixels": 0.2,
        "repeatDelayMs": 0.9,
        "repeatIntervalMs": 0.1
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(
    config.controller.scrollPixelsPerSecond,
    DEFAULT_CONFIG.controller.scrollPixelsPerSecond,
  );
  assert.equal(
    config.controller.horizontalJumpPixels,
    DEFAULT_CONFIG.controller.horizontalJumpPixels,
  );
  assert.equal(config.controller.repeatDelayMs, DEFAULT_CONFIG.controller.repeatDelayMs);
  assert.equal(config.controller.repeatIntervalMs, DEFAULT_CONFIG.controller.repeatIntervalMs);
  assert.equal(
    warnings.some((warning) => warning.path === 'controller.scrollPixelsPerSecond'),
    true,
  );
  assert.equal(
    warnings.some((warning) => warning.path === 'controller.horizontalJumpPixels'),
    true,
  );
  assert.equal(
    warnings.some((warning) => warning.path === 'controller.repeatDelayMs'),
    true,
  );
  assert.equal(
    warnings.some((warning) => warning.path === 'controller.repeatIntervalMs'),
    true,
  );
});

test('controller button index config rejects fractional values', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "controller": {
        "buttonIndices": {
          "select": 6.5,
          "leftStickPress": 9.1
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(
    config.controller.buttonIndices.select,
    DEFAULT_CONFIG.controller.buttonIndices.select,
  );
  assert.equal(
    config.controller.buttonIndices.leftStickPress,
    DEFAULT_CONFIG.controller.buttonIndices.leftStickPress,
  );
  assert.equal(
    warnings.some((warning) => warning.path === 'controller.buttonIndices.select'),
    true,
  );
  assert.equal(
    warnings.some((warning) => warning.path === 'controller.buttonIndices.leftStickPress'),
    true,
  );
});

test('runtime options registry is centralized', () => {
  const ids = RUNTIME_OPTION_REGISTRY.map((entry) => entry.id);
  assert.deepEqual(ids, [
    'anki.autoUpdateNewCards',
    'anki.mediaReviewTiming',
    'subtitle.annotation.knownWords.highlightEnabled',
    'subtitle.annotation.knownWords.maturityEnabled',
    'subtitle.annotation.nPlusOne',
    'subtitle.annotation.jlpt',
    'subtitle.annotation.frequency',
    'anki.nPlusOneMatchMode',
    'anki.kikuFieldGrouping',
    'anki.senrenFieldGrouping',
  ]);
});

test('validates ankiConnect knownWords behavior values', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "knownWords": {
          "highlightEnabled": "yes",
          "refreshMinutes": -5,
          "addMinedWordsImmediately": "no"
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(
    config.ankiConnect.knownWords.highlightEnabled,
    DEFAULT_CONFIG.ankiConnect.knownWords.highlightEnabled,
  );
  assert.equal(
    config.ankiConnect.knownWords.refreshMinutes,
    DEFAULT_CONFIG.ankiConnect.knownWords.refreshMinutes,
  );
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.knownWords.highlightEnabled'));
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.knownWords.refreshMinutes'));
  assert.equal(
    config.ankiConnect.knownWords.addMinedWordsImmediately,
    DEFAULT_CONFIG.ankiConnect.knownWords.addMinedWordsImmediately,
  );
  assert.ok(
    warnings.some((warning) => warning.path === 'ankiConnect.knownWords.addMinedWordsImmediately'),
  );
});

test('accepts valid ankiConnect knownWords behavior values', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "knownWords": {
          "highlightEnabled": true,
          "refreshMinutes": 120,
          "addMinedWordsImmediately": false
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.ankiConnect.knownWords.highlightEnabled, true);
  assert.equal(config.ankiConnect.knownWords.refreshMinutes, 120);
  assert.equal(config.ankiConnect.knownWords.addMinedWordsImmediately, false);
});

test('validates ankiConnect n+1 minimum sentence word count', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "minSentenceWords": 0
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(
    config.ankiConnect.nPlusOne.minSentenceWords,
    DEFAULT_CONFIG.ankiConnect.nPlusOne.minSentenceWords,
  );
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.nPlusOne.minSentenceWords'));
});

test('accepts valid ankiConnect n+1 minimum sentence word count', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "minSentenceWords": 4
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.ankiConnect.nPlusOne.minSentenceWords, 4);
});

test('validates ankiConnect knownWords match mode values', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "knownWords": {
          "matchMode": "bad-mode"
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(
    config.ankiConnect.knownWords.matchMode,
    DEFAULT_CONFIG.ankiConnect.knownWords.matchMode,
  );
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.knownWords.matchMode'));
});

test('accepts valid ankiConnect knownWords match mode values', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "knownWords": {
          "matchMode": "surface"
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.ankiConnect.knownWords.matchMode, 'surface');
});

test('ignores invalid legacy ankiConnect n+1 color value after migration attempt', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "nPlusOne": "not-a-color"
        },
        "knownWords": {
          "color": 123
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(config.subtitleStyle.nPlusOneColor, DEFAULT_CONFIG.subtitleStyle.nPlusOneColor);
  assert.equal(config.subtitleStyle.knownWordColor, DEFAULT_CONFIG.subtitleStyle.knownWordColor);
  assert.ok(warnings.every((warning) => warning.path !== 'ankiConnect.nPlusOne.nPlusOne'));
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.knownWords.color'));
});

test('resolves legacy ankiConnect n+1 color value without rewriting config', () => {
  const dir = makeTempDir();
  const configPath = path.join(dir, 'config.jsonc');
  const originalContent = `{
      "ankiConnect": {
        "nPlusOne": {
          "nPlusOne": "#c6a0f6"
        },
        "knownWords": {
          "color": "#a6da95"
        }
      }
    }`;
  fs.writeFileSync(configPath, originalContent, 'utf-8');

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.subtitleStyle.nPlusOneColor, '#c6a0f6');
  assert.equal(config.subtitleStyle.knownWordColor, '#a6da95');
  assert.equal(fs.readFileSync(configPath, 'utf-8'), originalContent);
});

test('legacy migration failures are logged and rethrown', () => {
  const source = fs.readFileSync(path.join(process.cwd(), 'src/config/service.ts'), 'utf-8');
  const catchBlock = source.match(/catch\s*\(error\)\s*\{(?<body>[\s\S]*?)\n    \}/)?.groups?.body;

  assert.ok(catchBlock);
  assert.match(catchBlock, /legacy config migration failed/);
  assert.match(catchBlock, /console\.error/);
  assert.match(catchBlock, /throw error;/);
});

test('resolves legacy ankiConnect nPlusOne known-word settings without rewriting config', () => {
  const dir = makeTempDir();
  const configPath = path.join(dir, 'config.jsonc');
  const originalContent = `{
      "ankiConnect": {
        "nPlusOne": {
          "highlightEnabled": true,
          "refreshMinutes": 90,
          "matchMode": "surface",
          "decks": ["Mining", "Kaishi 1.5k"],
          "knownWord": "#a6da95"
        }
      }
    }`;
  fs.writeFileSync(configPath, originalContent, 'utf-8');

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(config.ankiConnect.knownWords.highlightEnabled, true);
  assert.equal(config.ankiConnect.nPlusOne.enabled, DEFAULT_CONFIG.ankiConnect.nPlusOne.enabled);
  assert.equal(config.ankiConnect.knownWords.refreshMinutes, 90);
  assert.equal(config.ankiConnect.knownWords.matchMode, 'surface');
  assert.deepEqual(config.ankiConnect.knownWords.decks, {
    Mining: ['Expression', 'Word', 'Reading', 'Word Reading'],
    'Kaishi 1.5k': ['Expression', 'Word', 'Reading', 'Word Reading'],
  });
  assert.equal(config.subtitleStyle.knownWordColor, '#a6da95');
  assert.equal(fs.readFileSync(configPath, 'utf-8'), originalContent);
  assert.ok(warnings.every((warning) => !warning.path.startsWith('ankiConnect.nPlusOne.')));
});

test('resolves duplicate ankiConnect nPlusOne objects without rewriting config', () => {
  const dir = makeTempDir();
  const configPath = path.join(dir, 'config.jsonc');
  const originalContent = `{
      "ankiConnect": {
        "nPlusOne": {
          "enabled": true,
          "minSentenceWords": 3
        },
        "knownWords": {
          "highlightEnabled": true
        },
        "nPlusOne": {
          "minSentenceWords": "3"
        }
      }
    }`;
  fs.writeFileSync(configPath, originalContent, 'utf-8');

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.ankiConnect.nPlusOne.enabled, true);
  assert.equal(fs.readFileSync(configPath, 'utf-8'), originalContent);
});

test('later invalid duplicate nPlusOne values supersede earlier valid values', () => {
  const dir = makeTempDir();
  const configPath = path.join(dir, 'config.jsonc');
  const originalContent = `{
      "ankiConnect": {
        "nPlusOne": {
          "enabled": true,
          "minSentenceWords": 4
        },
        "nPlusOne": {
          "enabled": "yes",
          "minSentenceWords": "4"
        }
      }
    }`;
  fs.writeFileSync(configPath, originalContent, 'utf-8');

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(config.ankiConnect.nPlusOne.enabled, DEFAULT_CONFIG.ankiConnect.nPlusOne.enabled);
  assert.equal(
    config.ankiConnect.nPlusOne.minSentenceWords,
    DEFAULT_CONFIG.ankiConnect.nPlusOne.minSentenceWords,
  );
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.nPlusOne.enabled'));
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.nPlusOne.minSentenceWords'));
  assert.equal(fs.readFileSync(configPath, 'utf-8'), originalContent);
});

test('supports legacy ankiConnect.behavior N+1 settings as fallback', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "behavior": {
          "nPlusOneHighlightEnabled": true,
          "nPlusOneRefreshMinutes": 90,
          "nPlusOneMatchMode": "surface"
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(config.ankiConnect.knownWords.highlightEnabled, true);
  assert.equal(config.ankiConnect.nPlusOne.enabled, DEFAULT_CONFIG.ankiConnect.nPlusOne.enabled);
  assert.equal(config.ankiConnect.knownWords.refreshMinutes, 90);
  assert.equal(config.ankiConnect.knownWords.matchMode, 'surface');
  assert.ok(
    warnings.some(
      (warning) =>
        warning.path === 'ankiConnect.behavior.nPlusOneHighlightEnabled' ||
        warning.path === 'ankiConnect.behavior.nPlusOneRefreshMinutes' ||
        warning.path === 'ankiConnect.behavior.nPlusOneMatchMode',
    ),
  );
});

test('accepts top-level ai config', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ai": {
        "enabled": true,
        "apiKey": "abc123",
        "apiKeyCommand": "pass show subminer/ai",
        "baseUrl": "https://openrouter.ai/api",
        "model": "openrouter/test-model",
        "systemPrompt": "Return only fixed subtitles.",
        "requestTimeoutMs": 20000
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  assert.equal(config.ai.enabled, true);
  assert.equal(config.ai.apiKey, 'abc123');
  assert.equal(config.ai.apiKeyCommand, 'pass show subminer/ai');
  assert.equal(config.ai.baseUrl, 'https://openrouter.ai/api');
  assert.equal(config.ai.model, 'openrouter/test-model');
  assert.equal(config.ai.systemPrompt, 'Return only fixed subtitles.');
  assert.equal(config.ai.requestTimeoutMs, 20000);
});

test('accepts per-feature ai overrides for anki and YouTube subtitles', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ai": {
        "enabled": true,
        "apiKeyCommand": "pass show subminer/ai",
        "baseUrl": "https://openrouter.ai/api",
        "model": "openrouter/shared-model",
        "systemPrompt": "Legacy shared prompt."
      },
      "ankiConnect": {
        "ai": {
          "enabled": true,
          "model": "openrouter/anki-model",
          "systemPrompt": "Translate mined sentence text."
        }
      },
      "youtubeSubgen": {
        "ai": {
          "model": "openrouter/subgen-model",
          "systemPrompt": "Fix subtitle mistakes only."
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.ai.enabled, true);
  assert.equal(config.ai.model, 'openrouter/shared-model');
  assert.equal(config.ankiConnect.ai.enabled, true);
  assert.equal(config.ankiConnect.ai.model, 'openrouter/anki-model');
  assert.equal(config.ankiConnect.ai.systemPrompt, 'Translate mined sentence text.');
  assert.equal(config.youtubeSubgen.ai.model, 'openrouter/subgen-model');
  assert.equal(config.youtubeSubgen.ai.systemPrompt, 'Fix subtitle mistakes only.');
});

test('warns and falls back when ankiConnect.ai override values are invalid', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "ai": {
          "enabled": "yes",
          "model": 123,
          "systemPrompt": true
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.deepEqual(config.ankiConnect.ai, DEFAULT_CONFIG.ankiConnect.ai);
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.ai.enabled'));
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.ai.model'));
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.ai.systemPrompt'));
});

test('falls back and warns when legacy ankiConnect migration values are invalid', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "audioField": 123,
        "generateAudio": "yes",
        "imageType": "gif",
        "imageQuality": -1,
        "mediaInsertMode": "middle",
        "notificationType": "toast"
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(config.ankiConnect.fields.audio, DEFAULT_CONFIG.ankiConnect.fields.audio);
  assert.equal(
    config.ankiConnect.media.generateAudio,
    DEFAULT_CONFIG.ankiConnect.media.generateAudio,
  );
  assert.equal(config.ankiConnect.media.imageType, DEFAULT_CONFIG.ankiConnect.media.imageType);
  assert.equal(
    config.ankiConnect.media.imageQuality,
    DEFAULT_CONFIG.ankiConnect.media.imageQuality,
  );
  assert.equal(
    config.ankiConnect.behavior.mediaInsertMode,
    DEFAULT_CONFIG.ankiConnect.behavior.mediaInsertMode,
  );
  assert.equal(
    config.ankiConnect.behavior.notificationType,
    DEFAULT_CONFIG.ankiConnect.behavior.notificationType,
  );

  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.audioField'));
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.generateAudio'));
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.imageType'));
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.imageQuality'));
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.mediaInsertMode'));
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.notificationType'));
});

test('maps valid legacy ankiConnect values to equivalent modern config', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "audioField": "AudioLegacy",
        "imageField": "ImageLegacy",
        "generateAudio": false,
        "imageType": "avif",
        "imageFormat": "webp",
        "imageQuality": 88,
        "mediaInsertMode": "prepend",
        "notificationType": "both",
        "autoUpdateNewCards": false
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.ankiConnect.fields.audio, 'AudioLegacy');
  assert.equal(config.ankiConnect.fields.image, 'ImageLegacy');
  assert.equal(config.ankiConnect.media.generateAudio, false);
  assert.equal(config.ankiConnect.media.imageType, 'avif');
  assert.equal(config.ankiConnect.media.imageFormat, 'webp');
  assert.equal(config.ankiConnect.media.imageQuality, 88);
  assert.equal(config.ankiConnect.behavior.mediaInsertMode, 'prepend');
  assert.equal(config.ankiConnect.behavior.notificationType, 'both');
  assert.equal(config.ankiConnect.behavior.autoUpdateNewCards, false);
});

test('ignores deprecated isLapis sentence-card field overrides', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "isLapis": {
          "enabled": true,
          "sentenceCardModel": "Japanese sentences",
          "sentenceCardSentenceField": "CustomSentence",
          "sentenceCardAudioField": "CustomAudio"
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  const lapisConfig = config.ankiConnect.isLapis as Record<string, unknown>;
  assert.equal(lapisConfig.sentenceCardSentenceField, undefined);
  assert.equal(lapisConfig.sentenceCardAudioField, undefined);
  assert.ok(
    warnings.some((warning) => warning.path === 'ankiConnect.isLapis.sentenceCardSentenceField'),
  );
  assert.ok(
    warnings.some((warning) => warning.path === 'ankiConnect.isLapis.sentenceCardAudioField'),
  );
});

test('accepts a Kiku/Lapis word card kind and warns on an unknown one', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "isKiku": { "enabled": true },
        "lapisKiku": { "wordCardKind": "click" }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  assert.equal(service.getConfig().ankiConnect.lapisKiku.wordCardKind, 'click');
  assert.equal(service.getWarnings().length, 0);

  const invalidDir = makeTempDir();
  fs.writeFileSync(
    path.join(invalidDir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "lapisKiku": { "wordCardKind": "isClickCard" }
      }
    }`,
    'utf-8',
  );

  const invalidService = new ConfigService(invalidDir);
  assert.equal(invalidService.getConfig().ankiConnect.lapisKiku.wordCardKind, 'word-and-sentence');
  assert.ok(
    invalidService
      .getWarnings()
      .some((warning) => warning.path === 'ankiConnect.lapisKiku.wordCardKind'),
  );
});

test('forces Senren off when Kiku is also enabled and validates Senren fieldGrouping', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "isKiku": { "enabled": true },
        "isSenren": { "enabled": true }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  assert.equal(service.getConfig().ankiConnect.isKiku.enabled, true);
  assert.equal(service.getConfig().ankiConnect.isSenren.enabled, false);
  assert.ok(
    service.getWarnings().some((warning) => warning.path === 'ankiConnect.isSenren.enabled'),
  );

  const senrenOnlyDir = makeTempDir();
  fs.writeFileSync(
    path.join(senrenOnlyDir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "isSenren": { "enabled": true, "fieldGrouping": "sometimes" }
      }
    }`,
    'utf-8',
  );

  const senrenOnlyService = new ConfigService(senrenOnlyDir);
  assert.equal(senrenOnlyService.getConfig().ankiConnect.isSenren.enabled, true);
  assert.equal(senrenOnlyService.getConfig().ankiConnect.isSenren.fieldGrouping, 'auto');
  assert.ok(
    senrenOnlyService
      .getWarnings()
      .some((warning) => warning.path === 'ankiConnect.isSenren.fieldGrouping'),
  );
});

test('accepts valid ankiConnect knownWords deck object', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "knownWords": {
          "decks": { "Deck One": ["Word", "Reading"], "Deck Two": ["Expression"] }
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.deepEqual(config.ankiConnect.knownWords.decks, {
    'Deck One': ['Word', 'Reading'],
    'Deck Two': ['Expression'],
  });
});

test('accepts valid ankiConnect tags list', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "tags": ["SubMiner", "Mining"]
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.deepEqual(config.ankiConnect.tags, ['SubMiner', 'Mining']);
});

test('falls back to default when ankiConnect tags list is invalid', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "tags": ["SubMiner", 123]
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.deepEqual(config.ankiConnect.tags, ['SubMiner']);
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.tags'));
});

test('falls back to default when ankiConnect knownWords deck list is invalid', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "knownWords": {
          "decks": "not-an-array"
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.deepEqual(config.ankiConnect.knownWords.decks, {});
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.knownWords.decks'));
});

test('template generator includes known keys', () => {
  const output = generateConfigTemplate(DEFAULT_CONFIG);
  assert.match(output, /"ai":/);
  assert.match(output, /"ankiConnect":/);
  assert.match(output, /"controller":/);
  assert.match(output, /"logging":/);
  assert.match(output, /"websocket":/);
  assert.match(output, /"discordPresence":/);
  assert.match(output, /"startupWarmups":/);
  assert.match(output, /"updates":/);
  assert.match(output, /"youtube":/);
  assert.doesNotMatch(output, /"deviceId":/);
  assert.doesNotMatch(output, /"clientVersion":/);
  assert.doesNotMatch(output, /"youtubeSubgen":/);
  assert.match(output, /"characterDictionary":\s*\{/);
  assert.doesNotMatch(output, /"characterDictionary":\s*\{\s*"enabled":/);
  assert.match(output, /"preserveLineBreaks": false/);
  assert.match(output, /"knownWords"\s*:\s*\{/);
  assert.match(output, /"knownWordColor": "#a6da95"/);
  assert.match(output, /"nPlusOneColor": "#c6a0f6"/);
  assert.match(output, /"nPlusOne"\s*:\s*\{/);
  assert.match(output, /"minSentenceWords": 3/);
  assert.match(output, /auto-generated from src\/config\/definitions.ts/);
  assert.match(
    output,
    /"level": "warn",? \/\/ Minimum log level for runtime logging\. Values: debug \| info \| warn \| error/,
  );
  assert.match(
    output,
    /"enabled": false,? \/\/ Built-in subtitle websocket server mode\. Values: auto \| true \| false/,
  );
  assert.match(
    output,
    /"enabled": false,? \/\/ Annotated subtitle websocket server enabled state\. Values: true \| false/,
  );
  assert.match(
    output,
    /"scrollPixelsPerSecond": 900,? \/\/ Base popup scroll speed for controller stick input\./,
  );
  assert.match(
    output,
    /"triggerInputMode": "auto",? \/\/ How controller triggers are interpreted: auto, pressed-only, or thresholded analog\. Values: auto \| digital \| analog/,
  );
  assert.match(
    output,
    /"preferredGamepadId": "",? \/\/ Preferred controller id saved from the controller config modal\./,
  );
  assert.match(
    output,
    /"toggleLookup": \{\s*"kind": "button"[\s\S]*\},? \/\/ Controller binding descriptor for toggling lookup\. Use Alt\+C learn mode or set a raw button\/axis descriptor manually\./,
  );
  assert.match(
    output,
    /"kind": "button",? \/\/ Discrete binding input source kind\. When kind is "axis", set both axisIndex and direction\. Values: none \| button \| axis/,
  );
  assert.match(output, /"toggleLookup": \{\s*"kind": "button"/);
  assert.match(output, /"leftStickHorizontal": \{\s*"kind": "axis"/);
  assert.match(
    output,
    /"dpadFallback": "horizontal",? \/\/ Optional D-pad fallback used when this analog controller action should also read D-pad input\. Values: none \| horizontal \| vertical/,
  );
  assert.match(output, /"port": 6678,? \/\/ Annotated subtitle websocket server port\./);
  assert.match(
    output,
    /"openBrowser": false,? \/\/ Open the texthooker page in the default browser when the server starts\. Values: true \| false/,
  );
  assert.match(
    output,
    /"enabled": false,? \/\/ Enable overlay controller support through the Chrome Gamepad API\. Values: true \| false/,
  );
  assert.match(
    output,
    /"autoPauseVideoOnYomitanPopup": true,? \/\/ Automatically pause mpv playback while Yomitan popup is open, then resume when popup closes\. Values: true \| false/,
  );
  assert.match(
    output,
    /"enabled": true,? \/\/ Enable the subtitle sidebar feature for parsed subtitle sources\. Values: true \| false/,
  );
  assert.match(
    output,
    /"enabled": true,? \/\/ Enable AnkiConnect integration\. Values: true \| false/,
  );
  assert.match(
    output,
    /"enabled": false,? \/\/ Enable AI provider usage for Anki translation\/enrichment flows\. Values: true \| false/,
  );
  assert.match(
    output,
    /"model": "",? \/\/ Optional model override for Anki AI translation\/enrichment flows\./,
  );
  assert.match(
    output,
    /"enabled": false,? \/\/ Enable shared OpenAI-compatible AI provider features\. Values: true \| false/,
  );
  assert.match(
    output,
    /"enabled": true,? \/\/ Enable optional Discord Rich Presence updates\. Values: true \| false/,
  );
  assert.match(
    output,
    /"autoOpenBrowser": false,? \/\/ Automatically open the stats dashboard in a browser when the server starts\. Values: true \| false/,
  );
  assert.match(
    output,
    /"notificationType": "overlay",? \/\/ How SubMiner announces available updates\..*Values: overlay \| system \| both \| none \| osd \| osd-system/,
  );
  assert.match(
    output,
    /"channel": "stable",? \/\/ Release channel used for update checks\. Values: stable \| prerelease/,
  );
  assert.match(
    output,
    /"primarySubLanguages": \[\s*"ja",\s*"jpn"\s*\],? \/\/ Comma-separated primary subtitle language priority for managed subtitle auto-selection\./,
  );
  assert.doesNotMatch(output, /"mode": "automatic"/);
  assert.doesNotMatch(output, /"fixWithAi": false/);
  assert.doesNotMatch(output, /"whisperThreads": 4/);
  assert.match(
    output,
    /"launchAtStartup": false,? \/\/ Launch texthooker server automatically when SubMiner starts\. Values: true \| false/,
  );
});

test('template generator uses settings CSS declaration paths for appearance fields', () => {
  const output = generateConfigTemplate(DEFAULT_CONFIG);
  const parsed = parseConfigContent('config.example.jsonc', output);

  assert.deepEqual(
    getValueAtPath(parsed, 'subtitleStyle.css'),
    buildDefaultSubtitleCssDeclarations('primary'),
  );
  assert.deepEqual(
    getValueAtPath(parsed, 'subtitleStyle.secondary.css'),
    buildDefaultSubtitleCssDeclarations('secondary'),
  );
  assert.deepEqual(
    getValueAtPath(parsed, 'subtitleSidebar.css'),
    buildDefaultSubtitleCssDeclarations('sidebar'),
  );

  for (const scope of SUBTITLE_CSS_SCOPES) {
    for (const path of getSubtitleCssManagedConfigPaths(scope)) {
      assert.equal(
        getValueAtPath(parsed, path),
        undefined,
        `${path} should be represented by ${getSubtitleCssPath(scope)} in the generated template`,
      );
    }
  }
});

test('template generator shows built-in default keybindings in the keybindings array', () => {
  const output = generateConfigTemplate(DEFAULT_CONFIG);
  const parsed = parseConfigContent('config.example.jsonc', output) as {
    keybindings?: unknown;
  };

  assert.deepEqual(parsed.keybindings, DEFAULT_KEYBINDINGS);
});
