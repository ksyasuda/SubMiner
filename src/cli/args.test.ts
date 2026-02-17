import test from "node:test";
import assert from "node:assert/strict";
import { hasExplicitCommand, parseArgs, shouldStartApp } from "./args";

test("parseArgs parses booleans and value flags", () => {
  const args = parseArgs([
    "--start",
    "--socket",
    "/tmp/mpv.sock",
    "--backend=hyprland",
    "--port",
    "6000",
    "--log-level",
    "warn",
    "--debug",
  ]);

  assert.equal(args.start, true);
  assert.equal(args.socketPath, "/tmp/mpv.sock");
  assert.equal(args.backend, "hyprland");
  assert.equal(args.texthookerPort, 6000);
  assert.equal(args.logLevel, "warn");
  assert.equal(args.debug, true);
});

test("parseArgs ignores missing value after --log-level", () => {
  const args = parseArgs(["--log-level", "--start"]);
  assert.equal(args.logLevel, undefined);
  assert.equal(args.start, true);
});

test("hasExplicitCommand and shouldStartApp preserve command intent", () => {
  const stopOnly = parseArgs(["--stop"]);
  assert.equal(hasExplicitCommand(stopOnly), true);
  assert.equal(shouldStartApp(stopOnly), false);

  const toggle = parseArgs(["--toggle-visible-overlay"]);
  assert.equal(hasExplicitCommand(toggle), true);
  assert.equal(shouldStartApp(toggle), true);

  const noCommand = parseArgs(["--log-level", "warn"]);
  assert.equal(hasExplicitCommand(noCommand), false);
  assert.equal(shouldStartApp(noCommand), false);

  const refreshKnownWords = parseArgs(["--refresh-known-words"]);
  assert.equal(refreshKnownWords.help, false);
  assert.equal(hasExplicitCommand(refreshKnownWords), true);
  assert.equal(shouldStartApp(refreshKnownWords), false);

  const anilistStatus = parseArgs(["--anilist-status"]);
  assert.equal(anilistStatus.anilistStatus, true);
  assert.equal(hasExplicitCommand(anilistStatus), true);
  assert.equal(shouldStartApp(anilistStatus), false);

  const anilistRetryQueue = parseArgs(["--anilist-retry-queue"]);
  assert.equal(anilistRetryQueue.anilistRetryQueue, true);
  assert.equal(hasExplicitCommand(anilistRetryQueue), true);
  assert.equal(shouldStartApp(anilistRetryQueue), false);
});
