import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import {
  SubsyncManualPayload,
  SubsyncManualRunRequest,
  SubsyncResult,
} from "../../types";
import {
  CommandResult,
  codecToExtension,
  fileExists,
  formatTrackLabel,
  getTrackById,
  hasPathSeparators,
  MpvTrack,
  runCommand,
  SubsyncContext,
  SubsyncResolvedConfig,
} from "../../subsync/utils";

interface FileExtractionResult {
  path: string;
  temporary: boolean;
}

interface MpvClientLike {
  connected: boolean;
  currentAudioStreamIndex: number | null;
  send: (payload: { command: (string | number)[] }) => void;
  requestProperty: (name: string) => Promise<unknown>;
}

interface SubsyncCoreDeps {
  getMpvClient: () => MpvClientLike | null;
  getResolvedConfig: () => SubsyncResolvedConfig;
}

export interface TriggerSubsyncFromConfigDeps extends SubsyncCoreDeps {
  isSubsyncInProgress: () => boolean;
  setSubsyncInProgress: (inProgress: boolean) => void;
  showMpvOsd: (text: string) => void;
  runWithSubsyncSpinner: <T>(task: () => Promise<T>) => Promise<T>;
  openManualPicker: (payload: SubsyncManualPayload) => void;
}

function getMpvClientForSubsync(deps: SubsyncCoreDeps): MpvClientLike {
  const client = deps.getMpvClient();
  if (!client || !client.connected) {
    throw new Error("MPV not connected");
  }
  return client;
}

async function gatherSubsyncContext(
  client: MpvClientLike,
): Promise<SubsyncContext> {
  const [videoPathRaw, sidRaw, secondarySidRaw, trackListRaw] =
    await Promise.all([
      client.requestProperty("path"),
      client.requestProperty("sid"),
      client.requestProperty("secondary-sid"),
      client.requestProperty("track-list"),
    ]);

  const videoPath = typeof videoPathRaw === "string" ? videoPathRaw : "";
  if (!videoPath) {
    throw new Error("No video is currently loaded");
  }

  const tracks = Array.isArray(trackListRaw)
    ? (trackListRaw as MpvTrack[])
    : [];
  const subtitleTracks = tracks.filter((track) => track.type === "sub");
  const sid = typeof sidRaw === "number" ? sidRaw : null;
  const secondarySid =
    typeof secondarySidRaw === "number" ? secondarySidRaw : null;

  const primaryTrack = subtitleTracks.find((track) => track.id === sid);
  if (!primaryTrack) {
    throw new Error("No active subtitle track found");
  }

  const secondaryTrack =
    subtitleTracks.find((track) => track.id === secondarySid) ?? null;
  const sourceTracks = subtitleTracks
    .filter((track) => track.id !== sid)
    .filter((track) => {
      if (!track.external) return true;
      const filename = track["external-filename"];
      return typeof filename === "string" && filename.length > 0;
    });

  return {
    videoPath,
    primaryTrack,
    secondaryTrack,
    sourceTracks,
    audioStreamIndex: client.currentAudioStreamIndex,
  };
}

function ensureExecutablePath(pathOrName: string, name: string): string {
  if (!pathOrName) {
    throw new Error(`Missing ${name} path in config`);
  }

  if (hasPathSeparators(pathOrName) && !fileExists(pathOrName)) {
    throw new Error(`Configured ${name} executable not found: ${pathOrName}`);
  }
  return pathOrName;
}

async function extractSubtitleTrackToFile(
  ffmpegPath: string,
  videoPath: string,
  track: MpvTrack,
): Promise<FileExtractionResult> {
  if (track.external) {
    const externalPath = track["external-filename"];
    if (typeof externalPath !== "string" || externalPath.length === 0) {
      throw new Error("External subtitle track has no file path");
    }
    if (!fileExists(externalPath)) {
      throw new Error(`Subtitle file not found: ${externalPath}`);
    }
    return { path: externalPath, temporary: false };
  }

  const ffIndex = track["ff-index"];
  const extension = codecToExtension(track.codec);
  if (
    typeof ffIndex !== "number" ||
    !Number.isInteger(ffIndex) ||
    ffIndex < 0
  ) {
    throw new Error("Internal subtitle track has no valid ff-index");
  }
  if (!extension) {
    throw new Error(`Unsupported subtitle codec: ${track.codec ?? "unknown"}`);
  }

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "subminer-subsync-"));
  const outputPath = path.join(tempDir, `track_${ffIndex}.${extension}`);
  const extraction = await runCommand(ffmpegPath, [
    "-hide_banner",
    "-nostdin",
    "-y",
    "-loglevel",
    "quiet",
    "-an",
    "-vn",
    "-i",
    videoPath,
    "-map",
    `0:${ffIndex}`,
    "-f",
    extension,
    outputPath,
  ]);

  if (!extraction.ok || !fileExists(outputPath)) {
    throw new Error("Failed to extract internal subtitle track with ffmpeg");
  }

  return { path: outputPath, temporary: true };
}

function cleanupTemporaryFile(extraction: FileExtractionResult): void {
  if (!extraction.temporary) return;
  try {
    if (fileExists(extraction.path)) {
      fs.unlinkSync(extraction.path);
    }
  } catch {}
  try {
    const dir = path.dirname(extraction.path);
    if (fs.existsSync(dir)) {
      fs.rmdirSync(dir);
    }
  } catch {}
}

function buildRetimedPath(subPath: string): string {
  const parsed = path.parse(subPath);
  const suffix = `_retimed_${Date.now()}`;
  return path.join(
    parsed.dir,
    `${parsed.name}${suffix}${parsed.ext || ".srt"}`,
  );
}

async function runAlassSync(
  alassPath: string,
  referenceFile: string,
  inputSubtitlePath: string,
  outputPath: string,
): Promise<CommandResult> {
  return runCommand(alassPath, [referenceFile, inputSubtitlePath, outputPath]);
}

async function runFfsubsyncSync(
  ffsubsyncPath: string,
  videoPath: string,
  inputSubtitlePath: string,
  outputPath: string,
  audioStreamIndex: number | null,
): Promise<CommandResult> {
  const args = [videoPath, "-i", inputSubtitlePath, "-o", outputPath];
  if (audioStreamIndex !== null) {
    args.push("--reference-stream", `0:${audioStreamIndex}`);
  }
  return runCommand(ffsubsyncPath, args);
}

function loadSyncedSubtitle(
  client: MpvClientLike,
  pathToLoad: string,
): void {
  if (!client.connected) {
    throw new Error("MPV disconnected while loading subtitle");
  }
  client.send({ command: ["sub_add", pathToLoad] });
  client.send({ command: ["set_property", "sub-delay", 0] });
}

async function subsyncToReference(
  engine: "alass" | "ffsubsync",
  referenceFilePath: string,
  context: SubsyncContext,
  resolved: SubsyncResolvedConfig,
  client: MpvClientLike,
): Promise<SubsyncResult> {
  const ffmpegPath = ensureExecutablePath(resolved.ffmpegPath, "ffmpeg");
  const primaryExtraction = await extractSubtitleTrackToFile(
    ffmpegPath,
    context.videoPath,
    context.primaryTrack,
  );
  const outputPath = buildRetimedPath(primaryExtraction.path);

  try {
    let result: CommandResult;
    if (engine === "alass") {
      const alassPath = ensureExecutablePath(resolved.alassPath, "alass");
      result = await runAlassSync(
        alassPath,
        referenceFilePath,
        primaryExtraction.path,
        outputPath,
      );
    } else {
      const ffsubsyncPath = ensureExecutablePath(
        resolved.ffsubsyncPath,
        "ffsubsync",
      );
      result = await runFfsubsyncSync(
        ffsubsyncPath,
        context.videoPath,
        primaryExtraction.path,
        outputPath,
        context.audioStreamIndex,
      );
    }

    if (!result.ok || !fileExists(outputPath)) {
      return {
        ok: false,
        message: `${engine} synchronization failed`,
      };
    }

    loadSyncedSubtitle(client, outputPath);
    return {
      ok: true,
      message: `Subtitle synchronized with ${engine}`,
    };
  } finally {
    cleanupTemporaryFile(primaryExtraction);
  }
}

async function runSubsyncAutoInternal(
  deps: SubsyncCoreDeps,
): Promise<SubsyncResult> {
  const client = getMpvClientForSubsync(deps);
  const context = await gatherSubsyncContext(client);
  const resolved = deps.getResolvedConfig();
  const ffmpegPath = ensureExecutablePath(resolved.ffmpegPath, "ffmpeg");

  if (context.secondaryTrack) {
    let secondaryExtraction: FileExtractionResult | null = null;
    try {
      secondaryExtraction = await extractSubtitleTrackToFile(
        ffmpegPath,
        context.videoPath,
        context.secondaryTrack,
      );
      const alassResult = await subsyncToReference(
        "alass",
        secondaryExtraction.path,
        context,
        resolved,
        client,
      );
      if (alassResult.ok) {
        return alassResult;
      }
    } catch (error) {
      console.warn("Auto alass sync failed, trying ffsubsync fallback:", error);
    } finally {
      if (secondaryExtraction) {
        cleanupTemporaryFile(secondaryExtraction);
      }
    }
  }

  const ffsubsyncPath = ensureExecutablePath(
    resolved.ffsubsyncPath,
    "ffsubsync",
  );
  if (!ffsubsyncPath) {
    return {
      ok: false,
      message: "No secondary subtitle for alass and ffsubsync not configured",
    };
  }
  return subsyncToReference(
    "ffsubsync",
    context.videoPath,
    context,
    resolved,
    client,
  );
}

export async function runSubsyncManualService(
  request: SubsyncManualRunRequest,
  deps: SubsyncCoreDeps,
): Promise<SubsyncResult> {
  const client = getMpvClientForSubsync(deps);
  const context = await gatherSubsyncContext(client);
  const resolved = deps.getResolvedConfig();

  if (request.engine === "ffsubsync") {
    return subsyncToReference(
      "ffsubsync",
      context.videoPath,
      context,
      resolved,
      client,
    );
  }

  const sourceTrack = getTrackById(
    context.sourceTracks,
    request.sourceTrackId ?? null,
  );
  if (!sourceTrack) {
    return { ok: false, message: "Select a subtitle source track for alass" };
  }

  const ffmpegPath = ensureExecutablePath(resolved.ffmpegPath, "ffmpeg");
  let sourceExtraction: FileExtractionResult | null = null;
  try {
    sourceExtraction = await extractSubtitleTrackToFile(
      ffmpegPath,
      context.videoPath,
      sourceTrack,
    );
    return subsyncToReference(
      "alass",
      sourceExtraction.path,
      context,
      resolved,
      client,
    );
  } finally {
    if (sourceExtraction) {
      cleanupTemporaryFile(sourceExtraction);
    }
  }
}

export async function openSubsyncManualPickerService(
  deps: TriggerSubsyncFromConfigDeps,
): Promise<void> {
  const client = getMpvClientForSubsync(deps);
  const context = await gatherSubsyncContext(client);
  const payload: SubsyncManualPayload = {
    sourceTracks: context.sourceTracks
      .filter((track) => typeof track.id === "number")
      .map((track) => ({
        id: track.id as number,
        label: formatTrackLabel(track),
      })),
  };
  deps.openManualPicker(payload);
}

export async function triggerSubsyncFromConfigService(
  deps: TriggerSubsyncFromConfigDeps,
): Promise<void> {
  if (deps.isSubsyncInProgress()) {
    deps.showMpvOsd("Subsync already running");
    return;
  }

  const resolved = deps.getResolvedConfig();
  try {
    if (resolved.defaultMode === "manual") {
      await openSubsyncManualPickerService(deps);
      deps.showMpvOsd("Subsync: choose engine and source");
      return;
    }

    deps.setSubsyncInProgress(true);
    const result = await deps.runWithSubsyncSpinner(() =>
      runSubsyncAutoInternal(deps),
    );
    deps.showMpvOsd(result.message);
  } catch (error) {
    deps.showMpvOsd(`Subsync failed: ${(error as Error).message}`);
  } finally {
    deps.setSubsyncInProgress(false);
  }
}
