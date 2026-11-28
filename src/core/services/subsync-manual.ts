import * as path from 'path';
import { SubsyncManualRunRequest, SubsyncResult } from '../../types';
import {
  CommandResult,
  fileExists,
  getTrackById,
  MpvTrack,
  SubsyncContext,
  SubsyncResolvedConfig,
} from '../../subsync/utils';
import { alignSubtitleFileToReferenceSpans } from '../../subsync/alass-align';
import { extractAudioSpeechSpans, getMediaFileSignature } from '../../subsync/audio-speech';
import { buildTrackReferenceLabel } from '../../subsync/reference-discovery';
import {
  applyTimingToSubtitleFile,
  extractSpansFromParsedSubtitle,
  type SubtitleTimeSpan,
} from '../../subsync/subtitle-lines';
import { isRemoteMediaPath } from '../../jimaku/utils';
import {
  buildManualPayload,
  discoverAndEvaluateSubtitleCandidates,
  prepareRipCandidatesAndAudio,
} from './subsync-candidates';
import {
  gatherSubsyncContext,
  isPlaybackSessionUnchanged,
  normalizeMediaVideoPath,
  normalizeSubtitlePathForCompare,
  resolveTargetSlot,
  resolveTargetTrack,
} from './subsync-context';
import {
  LOSSY_REWRITE_ERROR_MESSAGE,
  buildRetimedPath,
  cleanupTemporaryFile,
  extractSubtitleTrackToFile,
  installRetimedFile,
  loadSyncedSubtitle,
  maybeWriteCleanedReferenceSrt,
  readAndParseSubtitleFile,
  removeIfExists,
  resolveReferenceFilePath,
  runAlassSync,
  runFfsubsyncSync,
  validateFfsubsyncReference,
  writeRetimedContent,
} from './subsync-files';
import {
  AUDIO_SPEECH_LEAD_SEC,
  type CandidateWithSpans,
  type FileExtractionResult,
  type MpvClientLike,
  PLAYBACK_CHANGED_RESULT_MESSAGE,
  type SubsyncCoreDeps,
  type SubsyncIo,
  type SubtitleSlot,
  type TriggerSubsyncFromConfigDeps,
  ensureExecutablePath,
  formatAlignmentSummary,
  formatOffsetSummary,
  getMpvClientForSubsync,
  isExecutableAvailable,
  summarizeCommandFailure,
} from './subsync-shared';

async function runBuiltinSubsync(
  request: SubsyncManualRunRequest,
  targetTrack: MpvTrack,
  targetSlot: SubtitleSlot,
  context: SubsyncContext,
  resolved: SubsyncResolvedConfig,
  client: MpvClientLike,
  deps: SubsyncCoreDeps,
): Promise<SubsyncResult> {
  const stripSigns = request.stripSigns !== false;
  const allowSplits = request.allowSplits !== false;

  const targetExtraction = await extractSubtitleTrackToFile(
    resolved.ffmpegPath,
    context.videoPath,
    targetTrack,
    context.workingDirectory,
    deps,
  );
  const replaceTarget = resolved.replace !== false && !targetExtraction.temporary;
  const outputPath = buildRetimedPath(targetExtraction.path, replaceTarget);

  let refExtraction: FileExtractionResult | null = null;
  let keepTempTarget = false;
  try {
    const parsedTarget = readAndParseSubtitleFile(targetExtraction.path);
    if (parsedTarget.events.length === 0) {
      return { ok: false, message: 'Target subtitle file has no readable dialogue cues.' };
    }
    if (parsedTarget.lossless === false) {
      return { ok: false, message: LOSSY_REWRITE_ERROR_MESSAGE };
    }

    let referenceSpans: SubtitleTimeSpan[] = [];
    let referenceLabel = 'reference';
    let calibration: 'clean' | 'noisy' = 'clean';
    let leadSec = 0;
    let lengthSec = context.durationSec ?? 0;

    if (request.referenceMode === 'video') {
      if (isRemoteMediaPath(context.videoPath)) {
        return {
          ok: false,
          message:
            'Built-in audio sync cannot read remote stream URLs directly. Pick a reference subtitle track or browse a local subtitle file.',
        };
      }
      const ffmpegPath = ensureExecutablePath(resolved.ffmpegPath, 'ffmpeg');
      const vadPaths = deps.getVadPaths?.();
      const audioSpeech = await extractAudioSpeechSpans({
        ffmpegPath,
        videoPath: context.videoPath,
        audioStreamIndex: context.audioStreamIndex,
        signal: deps.signal,
        vadExecutablePath: vadPaths?.vadExecutablePath ?? null,
        vadModelPath: vadPaths?.vadModelPath ?? null,
      });
      if (audioSpeech.spans.length === 0) {
        return {
          ok: false,
          message: 'No spoken dialogue was detected in the video audio.',
        };
      }
      referenceSpans = audioSpeech.spans;
      referenceLabel = audioSpeech.source === 'silero-vad' ? 'audio (Silero VAD)' : 'video audio';
      calibration = audioSpeech.calibration;
      leadSec = AUDIO_SPEECH_LEAD_SEC;
      lengthSec = Math.max(lengthSec, audioSpeech.durationSec);
    } else if (request.referenceMode === 'file') {
      const refPath = resolveReferenceFilePath(request.referenceFilePath);
      if (!refPath) {
        return { ok: false, message: 'Select an existing reference subtitle file (up to 20 MB).' };
      }
      const parsedRef = readAndParseSubtitleFile(refPath);
      referenceSpans = extractSpansFromParsedSubtitle(parsedRef, { stripSigns }).spans;
      referenceLabel = path.basename(refPath);
    } else {
      const referenceTrack = getTrackById(context.subtitleTracks, request.referenceTrackId ?? null);
      if (!referenceTrack) {
        return { ok: false, message: 'Select a reference subtitle track.' };
      }
      if (referenceTrack.id === targetTrack.id) {
        return {
          ok: false,
          message: 'Reference and out-of-sync subtitles must be different tracks',
        };
      }
      refExtraction = await extractSubtitleTrackToFile(
        resolved.ffmpegPath,
        context.videoPath,
        referenceTrack,
        context.workingDirectory,
        deps,
      );
      const parsedRef = readAndParseSubtitleFile(refExtraction.path);
      referenceSpans = extractSpansFromParsedSubtitle(parsedRef, { stripSigns }).spans;
      referenceLabel = buildTrackReferenceLabel(
        referenceTrack,
        context.secondaryTrack?.id === referenceTrack.id,
      );
    }

    if (referenceSpans.length === 0) {
      return { ok: false, message: 'Selected reference has no usable dialogue timestamps.' };
    }

    const alignment = alignSubtitleFileToReferenceSpans(parsedTarget, referenceSpans, {
      stripSigns,
      allowSplits,
      calibration,
      leadSec,
      lengthSec,
    });

    if (!(await isPlaybackSessionUnchanged(client, context, targetTrack, targetSlot))) {
      return { ok: false, message: PLAYBACK_CHANGED_RESULT_MESSAGE };
    }

    writeRetimedContent(outputPath, alignment.rewrittenContent, replaceTarget);
    if (targetExtraction.temporary) {
      keepTempTarget = true;
    }
    await loadSyncedSubtitle(client, outputPath, targetSlot, targetTrack);
    return {
      ok: true,
      message: formatAlignmentSummary(referenceLabel, alignment),
    };
  } finally {
    cleanupTemporaryFile(refExtraction);
    if (!keepTempTarget) {
      cleanupTemporaryFile(targetExtraction);
    }
  }
}

async function runManualOffsetSubsync(
  offsetMs: number,
  targetTrack: MpvTrack,
  targetSlot: SubtitleSlot,
  context: SubsyncContext,
  resolved: SubsyncResolvedConfig,
  client: MpvClientLike,
  io: SubsyncIo = {},
): Promise<SubsyncResult> {
  const targetExtraction = await extractSubtitleTrackToFile(
    resolved.ffmpegPath,
    context.videoPath,
    targetTrack,
    context.workingDirectory,
    io,
  );
  const replaceTarget = resolved.replace !== false && !targetExtraction.temporary;
  const outputPath = buildRetimedPath(targetExtraction.path, replaceTarget);

  let keepTempTarget = false;
  try {
    const parsedTarget = readAndParseSubtitleFile(targetExtraction.path);
    if (parsedTarget.events.length === 0) {
      return { ok: false, message: 'Target subtitle file has no readable cues to shift.' };
    }
    if (parsedTarget.lossless === false) {
      return { ok: false, message: LOSSY_REWRITE_ERROR_MESSAGE };
    }
    const allIndices = parsedTarget.events.map((_, idx) => idx);
    const moves = new Array<number>(allIndices.length).fill(offsetMs / 1000);
    const rewritten = applyTimingToSubtitleFile(parsedTarget, allIndices, moves, 1.0);
    if (!(await isPlaybackSessionUnchanged(client, context, targetTrack, targetSlot))) {
      return { ok: false, message: PLAYBACK_CHANGED_RESULT_MESSAGE };
    }
    writeRetimedContent(outputPath, rewritten, replaceTarget);
    if (targetExtraction.temporary) {
      keepTempTarget = true;
    }
    await loadSyncedSubtitle(client, outputPath, targetSlot, targetTrack);
    return {
      ok: true,
      message: `Shifted subtitle by ${formatOffsetSummary(offsetMs)}`,
    };
  } finally {
    if (!keepTempTarget) {
      cleanupTemporaryFile(targetExtraction);
    }
  }
}

async function subsyncToReference(
  engine: 'alass' | 'ffsubsync',
  referenceFilePath: string,
  targetTrack: MpvTrack,
  context: SubsyncContext,
  resolved: SubsyncResolvedConfig,
  client: MpvClientLike,
  slot: SubtitleSlot,
  io: SubsyncIo = {},
): Promise<SubsyncResult> {
  const targetExtraction = await extractSubtitleTrackToFile(
    resolved.ffmpegPath,
    context.videoPath,
    targetTrack,
    context.workingDirectory,
    io,
  );
  const replaceTarget = resolved.replace !== false && !targetExtraction.temporary;
  const outputPath = buildRetimedPath(targetExtraction.path, replaceTarget);
  const parsedOutput = path.parse(outputPath);
  const stagedOutputPath = path.join(
    parsedOutput.dir,
    `.subminer-subsync-${process.pid}-${Date.now()}-${Math.random().toString(16).slice(2)}${parsedOutput.ext || '.srt'}`,
  );

  try {
    let result: CommandResult;
    if (engine === 'alass') {
      const alassPath = ensureExecutablePath(resolved.alassPath, 'alass');
      result = await runAlassSync(
        alassPath,
        referenceFilePath,
        targetExtraction.path,
        stagedOutputPath,
        io.signal,
      );
    } else {
      const ffsubsyncPath = ensureExecutablePath(resolved.ffsubsyncPath, 'ffsubsync');
      result = await runFfsubsyncSync(
        ffsubsyncPath,
        context.videoPath,
        targetExtraction.path,
        stagedOutputPath,
        context.audioStreamIndex,
        io.signal,
      );
    }

    if (!result.ok || !fileExists(stagedOutputPath)) {
      const details = summarizeCommandFailure(engine, result);
      return {
        ok: false,
        message: `${engine} synchronization failed: ${details}`,
      };
    }

    if (!(await isPlaybackSessionUnchanged(client, context, targetTrack, slot))) {
      return { ok: false, message: PLAYBACK_CHANGED_RESULT_MESSAGE };
    }

    installRetimedFile(stagedOutputPath, outputPath, replaceTarget);
    await loadSyncedSubtitle(client, outputPath, slot, targetTrack);
    return {
      ok: true,
      message: `Subtitle synchronized with ${engine}`,
    };
  } finally {
    removeIfExists(stagedOutputPath);
    cleanupTemporaryFile(targetExtraction);
  }
}

export async function runSubsyncManualUnguarded(
  request: SubsyncManualRunRequest,
  deps: SubsyncCoreDeps,
): Promise<SubsyncResult> {
  const client = getMpvClientForSubsync(deps);
  const context = await gatherSubsyncContext(client);
  const resolved = deps.getResolvedConfig();

  if (request.mediaPath !== undefined && request.mediaPath !== null) {
    const expectedVideoPath =
      request.mediaPath.trim().length > 0
        ? normalizeMediaVideoPath(request.mediaPath, context.workingDirectory)
        : '';
    if (
      !expectedVideoPath ||
      normalizeSubtitlePathForCompare(expectedVideoPath) !==
        normalizeSubtitlePathForCompare(context.videoPath)
    ) {
      return {
        ok: false,
        message:
          'Active video changed since the subsync picker opened. Reopen subsync for the current video.',
      };
    }
  }

  const targetTrack = resolveTargetTrack(request, context);
  if (!targetTrack) {
    return { ok: false, message: 'Select the out-of-sync subtitle track to retime' };
  }
  const targetSlot = resolveTargetSlot(targetTrack, context);

  if (request.engine === 'offset') {
    const offsetMs = request.offsetMs ?? 0;
    if (!Number.isFinite(offsetMs) || offsetMs === 0) {
      return { ok: false, message: 'Enter a non-zero offset in milliseconds' };
    }
    return runManualOffsetSubsync(
      offsetMs,
      targetTrack,
      targetSlot,
      context,
      resolved,
      client,
      deps,
    );
  }

  if (request.engine === 'builtin') {
    return runBuiltinSubsync(request, targetTrack, targetSlot, context, resolved, client, deps);
  }

  if (request.engine === 'ffsubsync') {
    try {
      validateFfsubsyncReference(context.videoPath);
    } catch (error) {
      return {
        ok: false,
        message: `ffsubsync synchronization failed: ${(error as Error).message}`,
      };
    }
    return subsyncToReference(
      'ffsubsync',
      context.videoPath,
      targetTrack,
      context,
      resolved,
      client,
      targetSlot,
    );
  }

  if (request.referenceMode === 'video') {
    if (isRemoteMediaPath(context.videoPath)) {
      return {
        ok: false,
        message:
          'alass cannot use a stream URL as reference. Pick a reference subtitle track instead.',
      };
    }
    return subsyncToReference(
      'alass',
      context.videoPath,
      targetTrack,
      context,
      resolved,
      client,
      targetSlot,
    );
  }

  if (request.referenceMode === 'file') {
    const refPath = resolveReferenceFilePath(request.referenceFilePath);
    if (!refPath) {
      return {
        ok: false,
        message: 'Select an existing reference subtitle file (up to 20 MB) for alass',
      };
    }
    const cleaned = maybeWriteCleanedReferenceSrt(refPath, request.stripSigns !== false);
    try {
      return await subsyncToReference(
        'alass',
        cleaned ? cleaned.path : refPath,
        targetTrack,
        context,
        resolved,
        client,
        targetSlot,
        deps,
      );
    } finally {
      cleanupTemporaryFile(cleaned);
    }
  }

  const referenceTrack = getTrackById(context.subtitleTracks, request.referenceTrackId ?? null);
  if (!referenceTrack) {
    return { ok: false, message: 'Select a reference subtitle track for alass' };
  }
  if (referenceTrack.id === targetTrack.id) {
    return { ok: false, message: 'Reference and out-of-sync subtitles must be different tracks' };
  }

  let referenceExtraction: FileExtractionResult | null = null;
  let cleanedRef: FileExtractionResult | null = null;
  try {
    referenceExtraction = await extractSubtitleTrackToFile(
      resolved.ffmpegPath,
      context.videoPath,
      referenceTrack,
      context.workingDirectory,
      deps,
    );
    cleanedRef = maybeWriteCleanedReferenceSrt(
      referenceExtraction.path,
      request.stripSigns !== false,
    );
    return await subsyncToReference(
      'alass',
      cleanedRef ? cleanedRef.path : referenceExtraction.path,
      targetTrack,
      context,
      resolved,
      client,
      targetSlot,
      deps,
    );
  } finally {
    cleanupTemporaryFile(cleanedRef);
    cleanupTemporaryFile(referenceExtraction);
  }
}

export async function openSubsyncManualPicker(
  deps: TriggerSubsyncFromConfigDeps,
  options: {
    guidanceMessage?: string | null;
    skipAudioAnalysis?: boolean;
  } = {},
): Promise<boolean> {
  const client = getMpvClientForSubsync(deps);
  const context = await gatherSubsyncContext(client);
  const resolved = deps.getResolvedConfig();
  const isRemote = isRemoteMediaPath(context.videoPath);

  let evaluatedCandidates: CandidateWithSpans[] | undefined;
  let audioSpeechAvailable: boolean | undefined;

  if (!options.skipAudioAnalysis) {
    const initialMediaSignature = !isRemote ? getMediaFileSignature(context.videoPath) : null;
    const prepared = await prepareRipCandidatesAndAudio(
      context,
      context.primaryTrack,
      resolved,
      deps,
      {
        allowSampledAudioFastPath: true,
      },
    );
    deps.signal?.throwIfAborted();
    audioSpeechAvailable =
      prepared.audioSpeech !== null ||
      prepared.sampledAudioWindows !== null ||
      (!isRemote && isExecutableAvailable(resolved.ffmpegPath, ['ffmpeg']));
    evaluatedCandidates = await discoverAndEvaluateSubtitleCandidates(
      context,
      context.primaryTrack,
      resolved,
      prepared.audioSpeech,
      {
        sampledAudioWindows: prepared.sampledAudioWindows,
        preExtractedEmbedded: prepared.preExtractedEmbedded,
        // Never stream a remote video over HTTP or full-read a local MKV during sampled fast-path just to populate the picker.
        skipEmbeddedExtraction: isRemote || prepared.sampledAudioWindows !== null,
        fetchImpl: deps.fetchImpl,
        signal: deps.signal,
      },
    );
    if (
      !(await isPlaybackSessionUnchanged(client, context, context.primaryTrack, 'primary')) ||
      (!isRemote && getMediaFileSignature(context.videoPath) !== initialMediaSignature)
    ) {
      return false;
    }
  }

  const payload = buildManualPayload(context, resolved, deps, {
    evaluatedCandidates,
    audioSpeechAvailable,
    guidanceMessage: options.guidanceMessage ?? null,
  });
  deps.openManualPicker(payload);
  return true;
}
