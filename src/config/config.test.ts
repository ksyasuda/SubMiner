import test from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { ConfigService, ConfigStartupParseError } from './service';
import { DEFAULT_CONFIG, RUNTIME_OPTION_REGISTRY } from './definitions';
import { generateConfigTemplate } from './template';

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-config-test-'));
}

test('loads defaults when config is missing', () => {
  const dir = makeTempDir();
  const service = new ConfigService(dir);
  const config = service.getConfig();
  assert.equal(config.websocket.port, DEFAULT_CONFIG.websocket.port);
  assert.equal(config.ankiConnect.behavior.autoUpdateNewCards, true);
  assert.deepEqual(config.ankiConnect.tags, ['SubMiner']);
  assert.equal(config.anilist.enabled, false);
  assert.equal(config.jellyfin.remoteControlEnabled, true);
  assert.equal(config.jellyfin.remoteControlAutoConnect, true);
  assert.equal(config.jellyfin.autoAnnounce, false);
  assert.equal(config.jellyfin.remoteControlDeviceName, 'SubMiner');
  assert.equal(config.startupWarmups.lowPowerMode, false);
  assert.equal(config.startupWarmups.mecab, true);
  assert.equal(config.startupWarmups.yomitanExtension, true);
  assert.equal(config.startupWarmups.subtitleDictionaries, true);
  assert.equal(config.startupWarmups.jellyfinRemoteSession, true);
  assert.equal(config.discordPresence.enabled, false);
  assert.equal(config.discordPresence.updateIntervalMs, 3_000);
  assert.equal(config.subtitleStyle.backgroundColor, 'rgb(30, 32, 48, 0.88)');
  assert.equal(config.subtitleStyle.preserveLineBreaks, false);
  assert.equal(config.subtitleStyle.autoPauseVideoOnHover, true);
  assert.equal(config.subtitleStyle.hoverTokenColor, '#f4dbd6');
  assert.equal(config.subtitleStyle.hoverTokenBackgroundColor, 'rgba(54, 58, 79, 0.84)');
  assert.equal(
    config.subtitleStyle.fontFamily,
    'M PLUS 1 Medium, Source Han Sans JP, Noto Sans CJK JP',
  );
  assert.equal(config.subtitleStyle.fontWeight, '600');
  assert.equal(config.subtitleStyle.lineHeight, 1.35);
  assert.equal(config.subtitleStyle.letterSpacing, '-0.01em');
  assert.equal(config.subtitleStyle.wordSpacing, 0);
  assert.equal(config.subtitleStyle.fontKerning, 'normal');
  assert.equal(config.subtitleStyle.textRendering, 'geometricPrecision');
  assert.equal(config.subtitleStyle.textShadow, '0 3px 10px rgba(0,0,0,0.69)');
  assert.equal(config.subtitleStyle.backdropFilter, 'blur(6px)');
  assert.equal(config.subtitleStyle.secondary.fontFamily, 'Manrope, Inter');
  assert.equal(config.subtitleStyle.secondary.fontColor, '#cad3f5');
  assert.equal(config.immersionTracking.enabled, true);
  assert.equal(config.immersionTracking.dbPath, '');
  assert.equal(config.immersionTracking.batchSize, 25);
  assert.equal(config.immersionTracking.flushIntervalMs, 500);
  assert.equal(config.immersionTracking.queueCap, 1000);
  assert.equal(config.immersionTracking.payloadCapBytes, 256);
  assert.equal(config.immersionTracking.maintenanceIntervalMs, 86_400_000);
  assert.equal(config.immersionTracking.retention.eventsDays, 7);
  assert.equal(config.immersionTracking.retention.telemetryDays, 30);
  assert.equal(config.immersionTracking.retention.dailyRollupsDays, 365);
  assert.equal(config.immersionTracking.retention.monthlyRollupsDays, 1825);
  assert.equal(config.immersionTracking.retention.vacuumIntervalDays, 7);
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

test('parses jellyfin remote control fields', () => {
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
  assert.equal(config.jellyfin.remoteControlDeviceName, 'SubMiner');
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
        "retention": {
          "eventsDays": 14,
          "telemetryDays": 45,
          "dailyRollupsDays": 730,
          "monthlyRollupsDays": 3650,
          "vacuumIntervalDays": 14
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
  assert.equal(config.immersionTracking.retention.dailyRollupsDays, 730);
  assert.equal(config.immersionTracking.retention.monthlyRollupsDays, 3650);
  assert.equal(config.immersionTracking.retention.vacuumIntervalDays, 14);
});

test('falls back for invalid immersion tracking tuning values', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "immersionTracking": {
        "batchSize": 0,
        "flushIntervalMs": 1,
        "queueCap": 5,
        "payloadCapBytes": 16,
        "maintenanceIntervalMs": 1000,
        "retention": {
          "eventsDays": 0,
          "telemetryDays": 99999,
          "dailyRollupsDays": 0,
          "monthlyRollupsDays": 999999,
          "vacuumIntervalDays": 0
        }
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
  assert.equal(config.immersionTracking.retention.eventsDays, 7);
  assert.equal(config.immersionTracking.retention.telemetryDays, 30);
  assert.equal(config.immersionTracking.retention.dailyRollupsDays, 365);
  assert.equal(config.immersionTracking.retention.monthlyRollupsDays, 1825);
  assert.equal(config.immersionTracking.retention.vacuumIntervalDays, 7);

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
    warnings.some((warning) => warning.path === 'immersionTracking.retention.dailyRollupsDays'),
  );
  assert.ok(
    warnings.some((warning) => warning.path === 'immersionTracking.retention.monthlyRollupsDays'),
  );
  assert.ok(
    warnings.some((warning) => warning.path === 'immersionTracking.retention.vacuumIntervalDays'),
  );
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
      "youtubeSubgen": {
        "primarySubLanguages": ["ja", "en",],
      },
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  assert.equal(config.websocket.port, 7788);
  assert.deepEqual(config.youtubeSubgen.primarySubLanguages, ['ja', 'en']);
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
    ['unknownFeature', 'websocket.enabled', 'websocket.port', 'logging.level'],
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
      "shortcuts": {
        "toggleVisibleOverlayGlobal": "Alt+Shift+U",
        "openJimaku": "Ctrl+Alt+J"
      },
      "youtubeSubgen": {
        "primarySubLanguages": ["ja", "jpn", "jp"]
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  assert.equal(config.shortcuts.toggleVisibleOverlayGlobal, 'Alt+Shift+U');
  assert.equal(config.shortcuts.openJimaku, 'Ctrl+Alt+J');
  assert.deepEqual(config.youtubeSubgen.primarySubLanguages, ['ja', 'jpn', 'jp']);
});

test('runtime options registry is centralized', () => {
  const ids = RUNTIME_OPTION_REGISTRY.map((entry) => entry.id);
  assert.deepEqual(ids, [
    'anki.autoUpdateNewCards',
    'subtitle.annotation.nPlusOne',
    'subtitle.annotation.jlpt',
    'subtitle.annotation.frequency',
    'anki.nPlusOneMatchMode',
    'anki.kikuFieldGrouping',
  ]);
});

test('validates ankiConnect n+1 behavior values', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "highlightEnabled": "yes",
          "refreshMinutes": -5
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(
    config.ankiConnect.nPlusOne.highlightEnabled,
    DEFAULT_CONFIG.ankiConnect.nPlusOne.highlightEnabled,
  );
  assert.equal(
    config.ankiConnect.nPlusOne.refreshMinutes,
    DEFAULT_CONFIG.ankiConnect.nPlusOne.refreshMinutes,
  );
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.nPlusOne.highlightEnabled'));
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.nPlusOne.refreshMinutes'));
});

test('accepts valid ankiConnect n+1 behavior values', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "highlightEnabled": true,
          "refreshMinutes": 120
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.ankiConnect.nPlusOne.highlightEnabled, true);
  assert.equal(config.ankiConnect.nPlusOne.refreshMinutes, 120);
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

test('validates ankiConnect n+1 match mode values', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "nPlusOne": {
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
    config.ankiConnect.nPlusOne.matchMode,
    DEFAULT_CONFIG.ankiConnect.nPlusOne.matchMode,
  );
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.nPlusOne.matchMode'));
});

test('accepts valid ankiConnect n+1 match mode values', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "matchMode": "surface"
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.ankiConnect.nPlusOne.matchMode, 'surface');
});

test('validates ankiConnect n+1 color values', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "nPlusOne": "not-a-color",
          "knownWord": 123
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(config.ankiConnect.nPlusOne.nPlusOne, DEFAULT_CONFIG.ankiConnect.nPlusOne.nPlusOne);
  assert.equal(
    config.ankiConnect.nPlusOne.knownWord,
    DEFAULT_CONFIG.ankiConnect.nPlusOne.knownWord,
  );
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.nPlusOne.nPlusOne'));
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.nPlusOne.knownWord'));
});

test('accepts valid ankiConnect n+1 color values', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "nPlusOne": "#c6a0f6",
          "knownWord": "#a6da95"
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.ankiConnect.nPlusOne.nPlusOne, '#c6a0f6');
  assert.equal(config.ankiConnect.nPlusOne.knownWord, '#a6da95');
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

  assert.equal(config.ankiConnect.nPlusOne.highlightEnabled, true);
  assert.equal(config.ankiConnect.nPlusOne.refreshMinutes, 90);
  assert.equal(config.ankiConnect.nPlusOne.matchMode, 'surface');
  assert.ok(
    warnings.some(
      (warning) =>
        warning.path === 'ankiConnect.behavior.nPlusOneHighlightEnabled' ||
        warning.path === 'ankiConnect.behavior.nPlusOneRefreshMinutes' ||
        warning.path === 'ankiConnect.behavior.nPlusOneMatchMode',
    ),
  );
});

test('warns when ankiConnect.openRouter is used and migrates to ai', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "openRouter": {
          "model": "openrouter/test-model"
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal((config.ankiConnect.ai as Record<string, unknown>).model, 'openrouter/test-model');
  assert.ok(
    warnings.some(
      (warning) =>
        warning.path === 'ankiConnect.openRouter' && warning.message.includes('ankiConnect.ai'),
    ),
  );
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

test('accepts valid ankiConnect n+1 deck list', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "decks": ["Deck One", "Deck Two"]
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.deepEqual(config.ankiConnect.nPlusOne.decks, ['Deck One', 'Deck Two']);
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

test('falls back to default when ankiConnect n+1 deck list is invalid', () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, 'config.jsonc'),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "decks": "not-an-array"
        }
      }
    }`,
    'utf-8',
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.deepEqual(config.ankiConnect.nPlusOne.decks, []);
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.nPlusOne.decks'));
});

test('template generator includes known keys', () => {
  const output = generateConfigTemplate(DEFAULT_CONFIG);
  assert.match(output, /"ankiConnect":/);
  assert.match(output, /"logging":/);
  assert.match(output, /"websocket":/);
  assert.match(output, /"discordPresence":/);
  assert.match(output, /"startupWarmups":/);
  assert.match(output, /"youtubeSubgen":/);
  assert.match(output, /"preserveLineBreaks": false/);
  assert.match(output, /"nPlusOne"\s*:\s*\{/);
  assert.match(output, /"nPlusOne": "#c6a0f6"/);
  assert.match(output, /"knownWord": "#a6da95"/);
  assert.match(output, /"minSentenceWords": 3/);
  assert.match(output, /auto-generated from src\/config\/definitions.ts/);
  assert.match(
    output,
    /"level": "info",? \/\/ Minimum log level for runtime logging\. Values: debug \| info \| warn \| error/,
  );
  assert.match(
    output,
    /"enabled": "auto",? \/\/ Built-in subtitle websocket server mode\. Values: auto \| true \| false/,
  );
  assert.match(
    output,
    /"enabled": false,? \/\/ Enable AnkiConnect integration\. Values: true \| false/,
  );
});
