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

test("parses invisible overlay config and new global shortcuts", () => {
  const dir = makeTempDir();
  fs.writeFileSync(
    path.join(dir, "config.jsonc"),
    `{
      "shortcuts": {
        "toggleVisibleOverlayGlobal": "Alt+Shift+U",
        "toggleInvisibleOverlayGlobal": "Alt+Shift+I"
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
  assert.equal(config.invisibleOverlay.startupVisibility, "hidden");
  assert.equal(config.bind_visible_overlay_to_mpv_sub_visibility, false);
  assert.deepEqual(config.youtubeSubgen.primarySubLanguages, ["ja", "jpn", "jp"]);
});

test("runtime options registry is centralized", () => {
  const ids = RUNTIME_OPTION_REGISTRY.map((entry) => entry.id);
  assert.deepEqual(ids, ["anki.autoUpdateNewCards", "anki.kikuFieldGrouping"]);
});

test("template generator includes known keys", () => {
  const output = generateConfigTemplate(DEFAULT_CONFIG);
  assert.match(output, /"ankiConnect":/);
  assert.match(output, /"websocket":/);
  assert.match(output, /"youtubeSubgen":/);
  assert.match(output, /auto-generated from src\/config\/definitions.ts/);
});
