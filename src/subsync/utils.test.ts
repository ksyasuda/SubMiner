import test from "node:test";
import assert from "node:assert/strict";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { getSubsyncConfig, runCommand } from "./utils";

test("getSubsyncConfig applies fallback executable paths for blank values", () => {
  const config = getSubsyncConfig({
    defaultMode: "manual",
    alass_path: " ",
    ffsubsync_path: "",
    ffmpeg_path: undefined,
  });

  assert.equal(config.defaultMode, "manual");
  assert.equal(config.alassPath, "/usr/bin/alass");
  assert.equal(config.ffsubsyncPath, "/usr/bin/ffsubsync");
  assert.equal(config.ffmpegPath, "/usr/bin/ffmpeg");
});

test("runCommand returns failure on timeout", async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "subsync-utils-"));
  const sleeperPath = path.join(tmpDir, "sleeper.sh");
  fs.writeFileSync(sleeperPath, "#!/bin/sh\nsleep 2\n", {
    encoding: "utf8",
    mode: 0o755,
  });
  fs.chmodSync(sleeperPath, 0o755);

  const result = await runCommand(sleeperPath, [], 50);

  assert.equal(result.ok, false);
});
