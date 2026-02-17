import test from "node:test";
import assert from "node:assert/strict";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { ConfigService } from "./service";
import { DEFAULT_CONFIG, RUNTIME_OPTION_REGISTRY } from "./definitions";
import { generateConfigTemplate } from "./template";

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), "subminer-config-test-"));
}

test("loads defaults when config is missing", () => {
  const dir = makeTempDir();
  const service = new ConfigService(dir);
  const config = service.getConfig();
  assert.equal(config.websocket.port, DEFAULT_CONFIG.websocket.port);
  assert.equal(config.ankiConnect.behavior.autoUpdateNewCards, true);
  assert.equal(config.anilist.enabled, false);
});

test("parses anilist.enabled and warns for invalid value", () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, "config.jsonc"),
    `{
      "anilist": {
        "enabled": "yes"
      }
    }`,
    "utf-8",
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(config.anilist.enabled, DEFAULT_CONFIG.anilist.enabled);
  assert.ok(warnings.some((warning) => warning.path === "anilist.enabled"));

  service.patchRawConfig({ anilist: { enabled: true } });
  assert.equal(service.getConfig().anilist.enabled, true);
});

test("parses jsonc and warns/falls back on invalid value", () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, "config.jsonc"),
    `{
      // invalid websocket port
      "websocket": { "port": "bad" }
    }`,
    "utf-8",
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  assert.equal(config.websocket.port, DEFAULT_CONFIG.websocket.port);
  assert.ok(service.getWarnings().some((w) => w.path === "websocket.port"));
});

test("accepts valid logging.level", () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, "config.jsonc"),
    `{
      "logging": {
        "level": "warn"
      }
    }`,
    "utf-8",
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.logging.level, "warn");
});

test("falls back for invalid logging.level and reports warning", () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, "config.jsonc"),
    `{
      "logging": {
        "level": "trace"
      }
    }`,
    "utf-8",
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(config.logging.level, DEFAULT_CONFIG.logging.level);
  assert.ok(
    warnings.some((warning) => warning.path === "logging.level"),
  );
});

test("parses invisible overlay config and new global shortcuts", () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, "config.jsonc"),
    `{
      "shortcuts": {
        "toggleVisibleOverlayGlobal": "Alt+Shift+U",
        "toggleInvisibleOverlayGlobal": "Alt+Shift+I",
        "openJimaku": "Ctrl+Alt+J"
      },
      "invisibleOverlay": {
        "startupVisibility": "hidden"
      },
      "bind_visible_overlay_to_mpv_sub_visibility": false,
      "youtubeSubgen": {
        "primarySubLanguages": ["ja", "jpn", "jp"]
      }
    }`,
    "utf-8",
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  assert.equal(config.shortcuts.toggleVisibleOverlayGlobal, "Alt+Shift+U");
  assert.equal(config.shortcuts.toggleInvisibleOverlayGlobal, "Alt+Shift+I");
  assert.equal(config.shortcuts.openJimaku, "Ctrl+Alt+J");
  assert.equal(config.invisibleOverlay.startupVisibility, "hidden");
  assert.equal(config.bind_visible_overlay_to_mpv_sub_visibility, false);
  assert.deepEqual(config.youtubeSubgen.primarySubLanguages, ["ja", "jpn", "jp"]);
});

test("runtime options registry is centralized", () => {
  const ids = RUNTIME_OPTION_REGISTRY.map((entry) => entry.id);
  assert.deepEqual(ids, [
    "anki.autoUpdateNewCards",
    "anki.nPlusOneMatchMode",
    "anki.kikuFieldGrouping",
  ]);
});

test("validates ankiConnect n+1 behavior values", () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, "config.jsonc"),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "highlightEnabled": "yes",
          "refreshMinutes": -5
        }
      }
    }`,
    "utf-8",
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
  assert.ok(
    warnings.some(
      (warning) => warning.path === "ankiConnect.nPlusOne.highlightEnabled",
    ),
  );
  assert.ok(
    warnings.some(
      (warning) => warning.path === "ankiConnect.nPlusOne.refreshMinutes",
    ),
  );
});

test("accepts valid ankiConnect n+1 behavior values", () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, "config.jsonc"),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "highlightEnabled": true,
          "refreshMinutes": 120
        }
      }
    }`,
    "utf-8",
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.ankiConnect.nPlusOne.highlightEnabled, true);
  assert.equal(config.ankiConnect.nPlusOne.refreshMinutes, 120);
});

test("validates ankiConnect n+1 minimum sentence word count", () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, "config.jsonc"),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "minSentenceWords": 0
        }
      }
    }`,
    "utf-8",
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(
    config.ankiConnect.nPlusOne.minSentenceWords,
    DEFAULT_CONFIG.ankiConnect.nPlusOne.minSentenceWords,
  );
  assert.ok(
    warnings.some(
      (warning) => warning.path === "ankiConnect.nPlusOne.minSentenceWords",
    ),
  );
});

test("accepts valid ankiConnect n+1 minimum sentence word count", () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, "config.jsonc"),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "minSentenceWords": 4
        }
      }
    }`,
    "utf-8",
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.ankiConnect.nPlusOne.minSentenceWords, 4);
});

test("validates ankiConnect n+1 match mode values", () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, "config.jsonc"),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "matchMode": "bad-mode"
        }
      }
    }`,
    "utf-8",
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(
    config.ankiConnect.nPlusOne.matchMode,
    DEFAULT_CONFIG.ankiConnect.nPlusOne.matchMode,
  );
  assert.ok(
    warnings.some((warning) =>
      warning.path === "ankiConnect.nPlusOne.matchMode",
    ),
  );
});

test("accepts valid ankiConnect n+1 match mode values", () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, "config.jsonc"),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "matchMode": "surface"
        }
      }
    }`,
    "utf-8",
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.ankiConnect.nPlusOne.matchMode, "surface");
});

test("validates ankiConnect n+1 color values", () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, "config.jsonc"),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "nPlusOne": "not-a-color",
          "knownWord": 123
        }
      }
    }`,
    "utf-8",
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(
    config.ankiConnect.nPlusOne.nPlusOne,
    DEFAULT_CONFIG.ankiConnect.nPlusOne.nPlusOne,
  );
  assert.equal(
    config.ankiConnect.nPlusOne.knownWord,
    DEFAULT_CONFIG.ankiConnect.nPlusOne.knownWord,
  );
  assert.ok(
    warnings.some((warning) => warning.path === "ankiConnect.nPlusOne.nPlusOne"),
  );
  assert.ok(
    warnings.some((warning) => warning.path === "ankiConnect.nPlusOne.knownWord"),
  );
});

test("accepts valid ankiConnect n+1 color values", () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, "config.jsonc"),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "nPlusOne": "#c6a0f6",
          "knownWord": "#a6da95"
        }
      }
    }`,
    "utf-8",
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.equal(config.ankiConnect.nPlusOne.nPlusOne, "#c6a0f6");
  assert.equal(config.ankiConnect.nPlusOne.knownWord, "#a6da95");
});

test("supports legacy ankiConnect.behavior N+1 settings as fallback", () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, "config.jsonc"),
    `{
      "ankiConnect": {
        "behavior": {
          "nPlusOneHighlightEnabled": true,
          "nPlusOneRefreshMinutes": 90,
          "nPlusOneMatchMode": "surface"
        }
      }
    }`,
    "utf-8",
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.equal(config.ankiConnect.nPlusOne.highlightEnabled, true);
  assert.equal(config.ankiConnect.nPlusOne.refreshMinutes, 90);
  assert.equal(config.ankiConnect.nPlusOne.matchMode, "surface");
  assert.ok(
    warnings.some(
      (warning) =>
        warning.path === "ankiConnect.behavior.nPlusOneHighlightEnabled" ||
        warning.path === "ankiConnect.behavior.nPlusOneRefreshMinutes" ||
        warning.path === "ankiConnect.behavior.nPlusOneMatchMode",
    ),
  );
});

test("accepts valid ankiConnect n+1 deck list", () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, "config.jsonc"),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "decks": ["Deck One", "Deck Two"]
        }
      }
    }`,
    "utf-8",
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();

  assert.deepEqual(config.ankiConnect.nPlusOne.decks, ["Deck One", "Deck Two"]);
});

test("falls back to default when ankiConnect n+1 deck list is invalid", () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, "config.jsonc"),
    `{
      "ankiConnect": {
        "nPlusOne": {
          "decks": "not-an-array"
        }
      }
    }`,
    "utf-8",
  );

  const service = new ConfigService(dir);
  const config = service.getConfig();
  const warnings = service.getWarnings();

  assert.deepEqual(config.ankiConnect.nPlusOne.decks, []);
  assert.ok(
    warnings.some((warning) => warning.path === "ankiConnect.nPlusOne.decks"),
  );
});

test("template generator includes known keys", () => {
  const output = generateConfigTemplate(DEFAULT_CONFIG);
  assert.match(output, /"ankiConnect":/);
  assert.match(output, /"logging":/);
  assert.match(output, /"websocket":/);
  assert.match(output, /"youtubeSubgen":/);
  assert.match(output, /"nPlusOne"\s*:\s*\{/);
  assert.match(output, /"nPlusOne": "#c6a0f6"/);
  assert.match(output, /"knownWord": "#a6da95"/);
  assert.match(output, /"minSentenceWords": 3/);
  assert.match(output, /auto-generated from src\/config\/definitions.ts/);
});
