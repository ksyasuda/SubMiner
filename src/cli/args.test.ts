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
    "--verbose",
  ]);

  assert.equal(args.start, true);
  assert.equal(args.socketPath, "/tmp/mpv.sock");
  assert.equal(args.backend, "hyprland");
  assert.equal(args.texthookerPort, 6000);
  assert.equal(args.logLevel, "warn");
  assert.equal(args.verbose, true);
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

  const noCommand = parseArgs(["--verbose"]);
  assert.equal(hasExplicitCommand(noCommand), false);
  assert.equal(shouldStartApp(noCommand), false);
});
