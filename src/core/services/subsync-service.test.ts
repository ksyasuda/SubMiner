import test from "node:test";
import assert from "node:assert/strict";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import {
  TriggerSubsyncFromConfigDeps,
  runSubsyncManualService,
  triggerSubsyncFromConfigService,
} from "./subsync-service";

function makeDeps(
  overrides: Partial<TriggerSubsyncFromConfigDeps> = {},
): TriggerSubsyncFromConfigDeps {
  const mpvClient = {
    connected: true,
    currentAudioStreamIndex: null,
    send: () => {},
    requestProperty: async (name: string) => {
      if (name === "path") return "/tmp/video.mkv";
      if (name === "sid") return 1;
      if (name === "secondary-sid") return null;
      if (name === "track-list") {
        return [
          { id: 1, type: "sub", selected: true, lang: "jpn" },
          {
            id: 2,
            type: "sub",
            selected: false,
            external: true,
            lang: "eng",
            "external-filename": "/tmp/ref.srt",
          },
          { id: 3, type: "audio", selected: true, "ff-index": 1 },
        ];
      }
      return null;
    },
  };

  return {
    getMpvClient: () => mpvClient,
    getResolvedConfig: () => ({
      defaultMode: "manual",
      alassPath: "/usr/bin/alass",
      ffsubsyncPath: "/usr/bin/ffsubsync",
      ffmpegPath: "/usr/bin/ffmpeg",
    }),
    isSubsyncInProgress: () => false,
    setSubsyncInProgress: () => {},
    showMpvOsd: () => {},
    runWithSubsyncSpinner: async <T>(task: () => Promise<T>) => task(),
    openManualPicker: () => {},
    ...overrides,
  };
}

test("triggerSubsyncFromConfigService returns early when already in progress", async () => {
  const osd: string[] = [];
  await triggerSubsyncFromConfigService(
    makeDeps({
      isSubsyncInProgress: () => true,
      showMpvOsd: (text) => {
        osd.push(text);
      },
    }),
  );
  assert.deepEqual(osd, ["Subsync already running"]);
});

test("triggerSubsyncFromConfigService opens manual picker in manual mode", async () => {
  const osd: string[] = [];
  let payloadTrackCount = 0;
  let inProgressState: boolean | null = null;

  await triggerSubsyncFromConfigService(
    makeDeps({
      openManualPicker: (payload) => {
        payloadTrackCount = payload.sourceTracks.length;
      },
      showMpvOsd: (text) => {
        osd.push(text);
      },
      setSubsyncInProgress: (value) => {
        inProgressState = value;
      },
    }),
  );

  assert.equal(payloadTrackCount, 1);
  assert.ok(osd.includes("Subsync: choose engine and source"));
  assert.equal(inProgressState, false);
});

test("triggerSubsyncFromConfigService reports failures to OSD", async () => {
  const osd: string[] = [];
  await triggerSubsyncFromConfigService(
    makeDeps({
      getMpvClient: () => null,
      showMpvOsd: (text) => {
        osd.push(text);
      },
    }),
  );

  assert.ok(osd.some((line) => line.startsWith("Subsync failed: MPV not connected")));
});

test("runSubsyncManualService requires a source track for alass", async () => {
  const result = await runSubsyncManualService(
    { engine: "alass", sourceTrackId: null },
    makeDeps(),
  );

  assert.deepEqual(result, {
    ok: false,
    message: "Select a subtitle source track for alass",
  });
});

test("triggerSubsyncFromConfigService reports path validation failures", async () => {
  const osd: string[] = [];
  const inProgress: boolean[] = [];

  await triggerSubsyncFromConfigService(
    makeDeps({
      getResolvedConfig: () => ({
        defaultMode: "auto",
        alassPath: "/missing/alass",
        ffsubsyncPath: "/missing/ffsubsync",
        ffmpegPath: "/missing/ffmpeg",
      }),
      setSubsyncInProgress: (value) => {
        inProgress.push(value);
      },
      showMpvOsd: (text) => {
        osd.push(text);
      },
    }),
  );

  assert.deepEqual(inProgress, [true, false]);
  assert.ok(
    osd.some((line) =>
      line.startsWith("Subsync failed: Configured ffmpeg executable not found"),
    ),
  );
});

function writeExecutableScript(filePath: string, content: string): void {
  fs.writeFileSync(filePath, content, { encoding: "utf8", mode: 0o755 });
  fs.chmodSync(filePath, 0o755);
}

test("runSubsyncManualService constructs ffsubsync command and returns success", async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "subsync-ffsubsync-"));
  const ffsubsyncLogPath = path.join(tmpDir, "ffsubsync-args.log");
  const ffsubsyncPath = path.join(tmpDir, "ffsubsync.sh");
  const ffmpegPath = path.join(tmpDir, "ffmpeg.sh");
  const alassPath = path.join(tmpDir, "alass.sh");
  const videoPath = path.join(tmpDir, "video.mkv");
  const primaryPath = path.join(tmpDir, "primary.srt");

  fs.writeFileSync(videoPath, "video");
  fs.writeFileSync(primaryPath, "sub");
  writeExecutableScript(
    ffmpegPath,
    "#!/bin/sh\nexit 0\n",
  );
  writeExecutableScript(
    alassPath,
    "#!/bin/sh\nexit 0\n",
  );
  writeExecutableScript(
    ffsubsyncPath,
    `#!/bin/sh\n: > "${ffsubsyncLogPath}"\nfor arg in "$@"; do printf '%s\\n' "$arg" >> "${ffsubsyncLogPath}"; done\nout=\"\"\nprev=\"\"\nfor arg in \"$@\"; do\n  if [ \"$prev\" = \"-o\" ]; then out=\"$arg\"; fi\n  prev=\"$arg\"\ndone\nif [ -n \"$out\" ]; then : > \"$out\"; fi\nexit 0\n`,
  );

  const sentCommands: Array<Array<string | number>> = [];
  const deps = makeDeps({
    getMpvClient: () => ({
      connected: true,
      currentAudioStreamIndex: 2,
      send: (payload) => {
        sentCommands.push(payload.command);
      },
      requestProperty: async (name: string) => {
        if (name === "path") return videoPath;
        if (name === "sid") return 1;
        if (name === "secondary-sid") return null;
        if (name === "track-list") {
          return [
            {
              id: 1,
              type: "sub",
              selected: true,
              external: true,
              "external-filename": primaryPath,
            },
          ];
        }
        return null;
      },
    }),
    getResolvedConfig: () => ({
      defaultMode: "manual",
      alassPath,
      ffsubsyncPath,
      ffmpegPath,
    }),
  });

  const result = await runSubsyncManualService(
    { engine: "ffsubsync", sourceTrackId: null },
    deps,
  );

  assert.equal(result.ok, true);
  assert.equal(result.message, "Subtitle synchronized with ffsubsync");
  const ffArgs = fs.readFileSync(ffsubsyncLogPath, "utf8").trim().split("\n");
  assert.equal(ffArgs[0], videoPath);
  assert.ok(ffArgs.includes("-i"));
  assert.ok(ffArgs.includes(primaryPath));
  assert.ok(ffArgs.includes("--reference-stream"));
  assert.ok(ffArgs.includes("0:2"));
  assert.equal(sentCommands[0]?.[0], "sub_add");
  assert.deepEqual(sentCommands[1], ["set_property", "sub-delay", 0]);
});

test("runSubsyncManualService constructs alass command and returns failure on non-zero exit", async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "subsync-alass-"));
  const alassLogPath = path.join(tmpDir, "alass-args.log");
  const alassPath = path.join(tmpDir, "alass.sh");
  const ffmpegPath = path.join(tmpDir, "ffmpeg.sh");
  const ffsubsyncPath = path.join(tmpDir, "ffsubsync.sh");
  const videoPath = path.join(tmpDir, "video.mkv");
  const primaryPath = path.join(tmpDir, "primary.srt");
  const sourcePath = path.join(tmpDir, "source.srt");

  fs.writeFileSync(videoPath, "video");
  fs.writeFileSync(primaryPath, "sub");
  fs.writeFileSync(sourcePath, "sub2");
  writeExecutableScript(ffmpegPath, "#!/bin/sh\nexit 0\n");
  writeExecutableScript(ffsubsyncPath, "#!/bin/sh\nexit 0\n");
  writeExecutableScript(
    alassPath,
    `#!/bin/sh\n: > "${alassLogPath}"\nfor arg in "$@"; do printf '%s\\n' "$arg" >> "${alassLogPath}"; done\nexit 1\n`,
  );

  const deps = makeDeps({
    getMpvClient: () => ({
      connected: true,
      currentAudioStreamIndex: null,
      send: () => {},
      requestProperty: async (name: string) => {
        if (name === "path") return videoPath;
        if (name === "sid") return 1;
        if (name === "secondary-sid") return null;
        if (name === "track-list") {
          return [
            {
              id: 1,
              type: "sub",
              selected: true,
              external: true,
              "external-filename": primaryPath,
            },
            {
              id: 2,
              type: "sub",
              selected: false,
              external: true,
              "external-filename": sourcePath,
            },
          ];
        }
        return null;
      },
    }),
    getResolvedConfig: () => ({
      defaultMode: "manual",
      alassPath,
      ffsubsyncPath,
      ffmpegPath,
    }),
  });

  const result = await runSubsyncManualService(
    { engine: "alass", sourceTrackId: 2 },
    deps,
  );

  assert.equal(result.ok, false);
  assert.equal(typeof result.message, "string");
  assert.equal(result.message.startsWith("alass synchronization failed"), true);
  const alassArgs = fs.readFileSync(alassLogPath, "utf8").trim().split("\n");
  assert.equal(alassArgs[0], sourcePath);
  assert.equal(alassArgs[1], primaryPath);
});

test("runSubsyncManualService resolves string sid values from mpv stream properties", async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "subsync-stream-sid-"));
  const ffsubsyncPath = path.join(tmpDir, "ffsubsync.sh");
  const ffsubsyncLogPath = path.join(tmpDir, "ffsubsync-args.log");
  const ffmpegPath = path.join(tmpDir, "ffmpeg.sh");
  const alassPath = path.join(tmpDir, "alass.sh");
  const videoPath = path.join(tmpDir, "video.mkv");
  const primaryPath = path.join(tmpDir, "primary.srt");

  fs.writeFileSync(videoPath, "video");
  fs.writeFileSync(primaryPath, "subtitle");
  writeExecutableScript(ffmpegPath, "#!/bin/sh\nexit 0\n");
  writeExecutableScript(alassPath, "#!/bin/sh\nexit 0\n");
  writeExecutableScript(
    ffsubsyncPath,
    `#!/bin/sh\nmkdir -p "${tmpDir}"\n: > "${ffsubsyncLogPath}"\nfor arg in "$@"; do printf '%s\\n' "$arg" >> "${ffsubsyncLogPath}"; done\nprev=""\nout=""\nfor arg in "$@"; do\n  if [ "$prev" = "--reference-stream" ]; then :; fi\n  if [ "$prev" = "-o" ]; then out="$arg"; fi\n  prev="$arg"\ndone\nif [ -n "$out" ]; then : > "$out"; fi`,
  );

  const deps = makeDeps({
    getMpvClient: () => ({
      connected: true,
      currentAudioStreamIndex: null,
      send: () => {},
      requestProperty: async (name: string) => {
        if (name === "path") return videoPath;
        if (name === "sid") return "1";
        if (name === "secondary-sid") return "2";
        if (name === "track-list") {
          return [
            {
              id: "1",
              type: "sub",
              selected: true,
              external: true,
              "external-filename": primaryPath,
            },
          ];
        }
        return null;
      },
    }),
    getResolvedConfig: () => ({
      defaultMode: "manual",
      alassPath,
      ffsubsyncPath,
      ffmpegPath,
    }),
  });

  const result = await runSubsyncManualService(
    { engine: "ffsubsync", sourceTrackId: null },
    deps,
  );

  assert.equal(result.ok, true);
  assert.equal(result.message, "Subtitle synchronized with ffsubsync");
  const ffArgs = fs.readFileSync(ffsubsyncLogPath, "utf8").trim().split("\n");
  const syncOutputIndex = ffArgs.indexOf("-o");
  assert.equal(syncOutputIndex >= 0, true);
  const outputPath = ffArgs[syncOutputIndex + 1];
  assert.equal(typeof outputPath, "string");
  assert.ok(outputPath.length > 0);
  assert.equal(fs.readFileSync(outputPath, "utf8"), "");
});
