import { SubsyncEngine, SubsyncManualPayload, SubsyncReferenceCandidate } from '../../types';
import {
  canFfmpegExtractSubtitleExtension,
  codecToExtension,
  fileExists,
  formatTrackLabel,
  MpvTrack,
  SubsyncContext,
  SubsyncResolvedConfig,
} from '../../subsync/utils';
import {
  evaluateReferenceAlignmentToAudio,
  evaluateReferenceAlignmentToSampledAudio,
  type SampledAudioWindow,
} from '../../subsync/alass-align';
import {
  cacheEmbeddedSubtitleTrack,
  extractAudioSpeechSpans,
  extractSampledAudioWindows,
  extractSinglePassRipMedia,
  getCachedAudioSpeechSpans,
  getCachedEmbeddedSubtitles,
  getMediaFileSignature,
  type AudioSpeechResult,
  type EmbeddedTrackSpec,
} from '../../subsync/audio-speech';
import {
  buildTrackReferenceLabel,
  discoverSidecarReferenceFiles,
  formatReferenceCandidateDetail,
  isEnglishReferenceMeta,
  isFullDialogueHintMeta,
  rankReferenceCandidates,
  type EvaluatedSubtitleReference,
} from '../../subsync/reference-discovery';
import {
  extractSpansFromParsedSubtitle,
  isSignsOrSongsTrackMeta,
  type ParsedSubtitleFile,
} from '../../subsync/subtitle-lines';
import { isRemoteMediaPath } from '../../jimaku/utils';
import { resolveLocalSubtitlePath } from './subsync-context';
import {
  cleanupTemporaryFile,
  extractSubtitleTrackToFile,
  readAndParseSubtitleFile,
} from './subsync-files';
import {
  AUDIO_SPEECH_LEAD_SEC,
  type CandidateWithSpans,
  type FileExtractionResult,
  MIN_REFERENCE_DIALOGUE_LINES,
  type SubsyncCoreDeps,
  isExecutableAvailable,
  yieldToEventLoop,
} from './subsync-shared';

async function resolveAudioSpeechForContext(
  context: SubsyncContext,
  resolved: SubsyncResolvedConfig,
  deps: SubsyncCoreDeps,
): Promise<AudioSpeechResult | null> {
  if (isRemoteMediaPath(context.videoPath) || !fileExists(context.videoPath)) {
    return null;
  }
  if (!isExecutableAvailable(resolved.ffmpegPath, ['ffmpeg'])) {
    return null;
  }
  const vadPaths = deps.getVadPaths?.();
  try {
    const speech = await extractAudioSpeechSpans({
      ffmpegPath: resolved.ffmpegPath,
      videoPath: context.videoPath,
      audioStreamIndex: context.audioStreamIndex,
      signal: deps.signal,
      vadExecutablePath: vadPaths?.vadExecutablePath ?? null,
      vadModelPath: vadPaths?.vadModelPath ?? null,
    });
    return speech.spans.length > 0 ? speech : null;
  } catch {
    return null;
  }
}

function collectExtractableEmbeddedRefTracks(
  context: SubsyncContext,
  targetTrack: MpvTrack,
): EmbeddedTrackSpec[] {
  const specs: EmbeddedTrackSpec[] = [];
  for (const track of context.subtitleTracks) {
    if (typeof track.id !== 'number' || track.id === targetTrack.id) continue;
    if (track.external) continue;
    // Skip embedded tracks already tagged as signs/songs or forced so we don't waste demux I/O.
    if (isSignsOrSongsTrackMeta(track)) continue;
    const ffIndex = track['ff-index'];
    const extension = codecToExtension(track.codec);
    if (
      typeof ffIndex === 'number' &&
      Number.isInteger(ffIndex) &&
      ffIndex >= 0 &&
      extension &&
      canFfmpegExtractSubtitleExtension(extension)
    ) {
      specs.push({ ffIndex, extension });
    }
  }
  return specs;
}

export async function discoverAndEvaluateSubtitleCandidates(
  context: SubsyncContext,
  targetTrack: MpvTrack,
  resolved: SubsyncResolvedConfig,
  audioSpeech: AudioSpeechResult | null,
  options: {
    sampledAudioWindows?: SampledAudioWindow[] | null;
    preExtractedEmbedded?: ReadonlyMap<number, ParsedSubtitleFile> | null;
    skipSignsEmbeddedExtraction?: boolean;
    skipEmbeddedExtraction?: boolean;
    fetchImpl?: typeof fetch;
    signal?: AbortSignal;
  } = {},
): Promise<CandidateWithSpans[]> {
  const results: CandidateWithSpans[] = [];
  const loadedPaths: string[] = [];
  const ffmpegAvailable = isExecutableAvailable(resolved.ffmpegPath, ['ffmpeg']);
  const sampledWindows = options.sampledAudioWindows ?? null;
  const preExtracted = options.preExtractedEmbedded ?? null;
  const skipSignsEmbedded = options.skipSignsEmbeddedExtraction !== false;
  const skipEmbeddedExtraction = options.skipEmbeddedExtraction === true;

  for (const track of context.subtitleTracks) {
    if (track.external && track['external-filename']) {
      const resolvedPath = resolveLocalSubtitlePath(
        track['external-filename'],
        context.workingDirectory,
      );
      if (resolvedPath) loadedPaths.push(resolvedPath);
    }
  }

  const evaluateParsedCandidate = (
    meta: Omit<
      EvaluatedSubtitleReference,
      'dialogueCount' | 'signCount' | 'audioStatus' | 'audioOffsetMs' | 'audioConfidence'
    >,
    parsed: ParsedSubtitleFile | null,
  ): CandidateWithSpans => {
    const stripped = parsed
      ? extractSpansFromParsedSubtitle(parsed, { stripSigns: true })
      : { spans: [], dialogueCount: 0, signCount: 0 };
    const all = parsed
      ? extractSpansFromParsedSubtitle(parsed, { stripSigns: false })
      : { spans: [] };

    const effectiveSignsOnly =
      meta.isSignsOnly ||
      (parsed !== null && stripped.dialogueCount < MIN_REFERENCE_DIALOGUE_LINES);

    let audioStatus: EvaluatedSubtitleReference['audioStatus'] = 'unverified';
    let audioOffsetMs: number | null = null;
    let audioConfidence = 0;

    if (!effectiveSignsOnly && stripped.spans.length >= MIN_REFERENCE_DIALOGUE_LINES) {
      if (audioSpeech && audioSpeech.spans.length > 0) {
        const audioEval = evaluateReferenceAlignmentToAudio(stripped.spans, audioSpeech.spans, {
          lengthSec: context.durationSec ?? audioSpeech.durationSec,
          audioCalibration: audioSpeech.calibration,
          audioLeadSec: AUDIO_SPEECH_LEAD_SEC,
        });
        audioStatus = audioEval.status;
        audioOffsetMs = audioEval.offsetMs;
        audioConfidence = audioEval.confidence;
      } else if (sampledWindows && sampledWindows.length >= 2) {
        const sampleEval = evaluateReferenceAlignmentToSampledAudio(
          stripped.spans,
          sampledWindows,
          { audioLeadSec: AUDIO_SPEECH_LEAD_SEC },
        );
        audioStatus = sampleEval.status;
        audioOffsetMs = sampleEval.offsetMs;
        audioConfidence = sampleEval.confidence;
      }
    }

    return {
      evaluated: {
        ...meta,
        isSignsOnly: effectiveSignsOnly,
        dialogueCount: stripped.dialogueCount,
        signCount: stripped.signCount,
        audioStatus,
        audioOffsetMs,
        audioConfidence,
      },
      dialogueSpans: stripped.spans,
      allSpans: all.spans,
    };
  };

  // 1. Evaluate all mpv subtitle tracks other than target.
  for (const track of context.subtitleTracks) {
    options.signal?.throwIfAborted();
    if (typeof track.id !== 'number' || track.id === targetTrack.id) continue;
    const isSecondary = context.secondaryTrack !== null && track.id === context.secondaryTrack.id;
    const isEmbedded = !track.external;
    const isSignsTrack = isSignsOrSongsTrackMeta(track);
    const externalFilename = track['external-filename'];
    const baseMeta = {
      key: `track:${track.id}`,
      kind: 'track' as const,
      trackId: track.id,
      filePath: track.external
        ? (resolveLocalSubtitlePath(externalFilename ?? '', context.workingDirectory) ?? null)
        : null,
      label: buildTrackReferenceLabel(track, isSecondary),
      isSecondary,
      isEmbedded,
      isEnglish: isEnglishReferenceMeta({
        lang: track.lang,
        title: track.title,
        filename: externalFilename,
      }),
      isFullDialogueHint: isFullDialogueHintMeta({
        title: track.title,
        filename: externalFilename,
      }),
      isSignsOnly: isSignsTrack,
    };

    if (isEmbedded && skipSignsEmbedded && isSignsTrack) {
      results.push(evaluateParsedCandidate(baseMeta, null));
      continue;
    }

    if (isEmbedded && preExtracted && typeof track['ff-index'] === 'number') {
      const preParsed = preExtracted.get(track['ff-index']);
      if (preParsed && preParsed.events.length > 0) {
        await yieldToEventLoop(options.signal);
        results.push(evaluateParsedCandidate(baseMeta, preParsed));
        continue;
      }
    }

    if (isEmbedded && skipEmbeddedExtraction) {
      results.push(evaluateParsedCandidate(baseMeta, null));
      continue;
    }

    let extraction: FileExtractionResult | null = null;
    try {
      const embeddedExtension = isEmbedded ? codecToExtension(track.codec) : null;
      if (
        isEmbedded &&
        (!ffmpegAvailable ||
          !embeddedExtension ||
          (!isRemoteMediaPath(context.videoPath) &&
            !canFfmpegExtractSubtitleExtension(embeddedExtension)))
      ) {
        results.push(evaluateParsedCandidate(baseMeta, null));
        continue;
      }
      const preExtractSignature =
        isEmbedded && !isRemoteMediaPath(context.videoPath)
          ? getMediaFileSignature(context.videoPath)
          : null;
      extraction = await extractSubtitleTrackToFile(
        resolved.ffmpegPath,
        context.videoPath,
        track,
        context.workingDirectory,
        options,
      );
      const parsed = readAndParseSubtitleFile(extraction.path);
      if (isEmbedded && !isRemoteMediaPath(context.videoPath)) {
        const postExtractSignature = getMediaFileSignature(context.videoPath);
        if (postExtractSignature !== preExtractSignature) {
          results.push(evaluateParsedCandidate(baseMeta, null));
          continue;
        }
        if (
          typeof track['ff-index'] === 'number' &&
          parsed.events.length > 0 &&
          preExtractSignature !== null
        ) {
          cacheEmbeddedSubtitleTrack(
            context.videoPath,
            track['ff-index'],
            parsed,
            preExtractSignature,
          );
        }
      }
      await yieldToEventLoop(options.signal);
      results.push(evaluateParsedCandidate(baseMeta, parsed));
    } catch {
      results.push(evaluateParsedCandidate(baseMeta, null));
    } finally {
      cleanupTemporaryFile(extraction);
    }
  }

  // 2. Discover unloaded sidecar subtitle files beside the local video.
  if (!isRemoteMediaPath(context.videoPath) && fileExists(context.videoPath)) {
    const sidecars = discoverSidecarReferenceFiles(context.videoPath, loadedPaths);
    for (const sidecar of sidecars) {
      options.signal?.throwIfAborted();
      const baseMeta = {
        key: `file:${sidecar.path}`,
        kind: 'sidecar' as const,
        trackId: null,
        filePath: sidecar.path,
        label: `Sidecar - ${sidecar.filename}`,
        isSecondary: false,
        isEmbedded: false,
        isEnglish: isEnglishReferenceMeta({
          lang: sidecar.languageHint,
          filename: sidecar.filename,
        }),
        isFullDialogueHint: isFullDialogueHintMeta({ filename: sidecar.filename }),
        isSignsOnly: sidecar.isSignsOnly,
      };
      try {
        const parsed = readAndParseSubtitleFile(sidecar.path);
        await yieldToEventLoop(options.signal);
        results.push(evaluateParsedCandidate(baseMeta, parsed));
      } catch {
        // Ignore unreadable sidecars.
      }
    }
  }

  const rankedEvaluated = rankReferenceCandidates(results.map((r) => r.evaluated));
  const byKey = new Map(results.map((r) => [r.evaluated.key, r]));
  return rankedEvaluated.map((ev) => byKey.get(ev.key)!);
}

export function buildManualPayload(
  context: SubsyncContext,
  resolved: SubsyncResolvedConfig,
  deps: SubsyncCoreDeps,
  options: {
    evaluatedCandidates?: CandidateWithSpans[];
    audioSpeechAvailable?: boolean;
    guidanceMessage?: string | null;
    defaultEngine?: SubsyncEngine;
    defaultReferenceKey?: string | null;
  } = {},
): SubsyncManualPayload {
  const subtitleTracks = context.subtitleTracks
    .filter((track) => typeof track.id === 'number')
    .map((track) => ({
      id: track.id as number,
      label: formatTrackLabel(track),
    }));
  const primaryTrackId =
    typeof context.primaryTrack.id === 'number' ? context.primaryTrack.id : null;
  const secondaryTrackId =
    typeof context.secondaryTrack?.id === 'number' ? context.secondaryTrack.id : null;

  const isLocalMedia = !isRemoteMediaPath(context.videoPath);
  const ffmpegAvailable = isExecutableAvailable(resolved.ffmpegPath, ['ffmpeg']);
  const alassAvailable = isExecutableAvailable(resolved.alassPath, ['alass', 'alass-cli']);
  const ffsubsyncAvailable =
    isLocalMedia && isExecutableAvailable(resolved.ffsubsyncPath, ['ffsubsync']);
  const audioReferenceAvailable = options.audioSpeechAvailable ?? (isLocalMedia && ffmpegAvailable);

  const referenceCandidates: SubsyncReferenceCandidate[] = [];
  if (options.evaluatedCandidates && options.evaluatedCandidates.length > 0) {
    for (const [idx, item] of options.evaluatedCandidates.entries()) {
      const ev = item.evaluated;
      referenceCandidates.push({
        key: ev.key,
        kind: ev.kind === 'sidecar' ? 'file' : 'track',
        label: ev.label,
        detail: formatReferenceCandidateDetail(ev) || undefined,
        trackId: ev.trackId,
        filePath: ev.filePath,
        audioStatus: ev.audioStatus,
        audioOffsetMs: ev.audioOffsetMs,
        isSignsOnly: ev.isSignsOnly,
        recommended: idx === 0 && ev.audioStatus === 'aligned' && !ev.isSignsOnly,
      });
    }
  } else {
    for (const track of context.subtitleTracks) {
      if (typeof track.id !== 'number' || track.id === primaryTrackId) continue;
      const isSecondary = secondaryTrackId !== null && track.id === secondaryTrackId;
      referenceCandidates.push({
        key: `track:${track.id}`,
        kind: 'track',
        label: buildTrackReferenceLabel(track, isSecondary),
        trackId: track.id,
        filePath: null,
        isSignsOnly: isSignsOrSongsTrackMeta(track),
      });
    }
  }

  if (isLocalMedia) {
    const audioRecommended =
      !referenceCandidates.some((c) => c.recommended) && audioReferenceAvailable;
    referenceCandidates.push({
      key: 'video',
      kind: 'video',
      label: 'Video audio (built-in speech detection)',
      detail: audioReferenceAvailable ? 'no reference subtitle needed' : 'requires ffmpeg',
      recommended: audioRecommended,
    });
  }

  const defaultReferenceTrackId =
    subtitleTracks.find((track) => track.id === secondaryTrackId)?.id ??
    subtitleTracks.find((track) => track.id !== primaryTrackId)?.id ??
    null;

  const recommendedKey =
    options.defaultReferenceKey ??
    referenceCandidates.find((c) => c.recommended)?.key ??
    (defaultReferenceTrackId !== null ? `track:${defaultReferenceTrackId}` : null) ??
    referenceCandidates[0]?.key ??
    null;

  return {
    subtitleTracks,
    defaultReferenceTrackId,
    defaultTargetTrackId: primaryTrackId,
    videoReferenceAvailable: isLocalMedia,
    ffsubsyncAvailable,
    mediaPath: context.videoPath,
    guidanceMessage: options.guidanceMessage ?? null,
    referenceCandidates,
    defaultReferenceKey: recommendedKey,
    defaultEngine: options.defaultEngine ?? 'builtin',
    builtinAvailable: true,
    alassAvailable,
    audioReferenceAvailable,
    canPickReferenceFile: deps.canPickReferenceFile?.() ?? false,
  };
}

export async function prepareRipCandidatesAndAudio(
  context: SubsyncContext,
  targetTrack: MpvTrack,
  resolved: SubsyncResolvedConfig,
  deps: SubsyncCoreDeps,
  options: { allowSampledAudioFastPath: boolean },
): Promise<{
  audioSpeech: AudioSpeechResult | null;
  sampledAudioWindows: SampledAudioWindow[] | null;
  preExtractedEmbedded: Map<number, ParsedSubtitleFile> | null;
}> {
  const isLocal = !isRemoteMediaPath(context.videoPath) && fileExists(context.videoPath);
  const ffmpegAvailable = isLocal && isExecutableAvailable(resolved.ffmpegPath, ['ffmpeg']);
  if (!ffmpegAvailable) {
    return { audioSpeech: null, sampledAudioWindows: null, preExtractedEmbedded: null };
  }

  const embeddedSpecs = collectExtractableEmbeddedRefTracks(context, targetTrack);
  const vadPaths = deps.getVadPaths?.();
  const cachedFull = getCachedAudioSpeechSpans(
    context.videoPath,
    context.audioStreamIndex,
    vadPaths,
  );

  // Fast path: when a non-signs external or sidecar reference already exists and
  // full audio is not cached yet, sample 4 seeked 60s windows (~0.3s local / ~3s NAS)
  // before ever streaming the full container over disk/NFS for embedded tracks or full VAD.
  const loadedExternalPaths: string[] = [];
  for (const track of context.subtitleTracks) {
    if (track.external && track['external-filename']) {
      const resolvedPath = resolveLocalSubtitlePath(
        track['external-filename'],
        context.workingDirectory,
      );
      if (resolvedPath) loadedExternalPaths.push(resolvedPath);
    }
  }
  const hasExternalOrSidecarHint =
    context.subtitleTracks.some(
      (t) =>
        typeof t.id === 'number' &&
        t.id !== targetTrack.id &&
        t.external &&
        !isSignsOrSongsTrackMeta(t),
    ) ||
    discoverSidecarReferenceFiles(context.videoPath, loadedExternalPaths).some(
      (s) => !s.isSignsOnly,
    );

  if (
    options.allowSampledAudioFastPath &&
    !cachedFull &&
    hasExternalOrSidecarHint &&
    context.durationSec !== null &&
    context.durationSec !== undefined &&
    context.durationSec >= 300
  ) {
    const sampled = await extractSampledAudioWindows({
      ffmpegPath: resolved.ffmpegPath,
      videoPath: context.videoPath,
      audioStreamIndex: context.audioStreamIndex,
      signal: deps.signal,
      durationSec: context.durationSec,
    });
    if (sampled && sampled.length >= 2) {
      return { audioSpeech: null, sampledAudioWindows: sampled, preExtractedEmbedded: null };
    }
  }

  // Single-pass demux when 1+ non-signs embedded tracks exist: reads the MKV ONCE
  // for all embedded subtitle tracks + full audio simultaneously.
  if (embeddedSpecs.length > 0) {
    const rip = await extractSinglePassRipMedia({
      ffmpegPath: resolved.ffmpegPath,
      videoPath: context.videoPath,
      audioStreamIndex: context.audioStreamIndex,
      signal: deps.signal,
      vadExecutablePath: vadPaths?.vadExecutablePath ?? null,
      vadModelPath: vadPaths?.vadModelPath ?? null,
      embeddedTracks: embeddedSpecs,
      includeAudio: true,
    });
    if (rip) {
      return {
        audioSpeech: rip.audioSpeech,
        sampledAudioWindows: null,
        preExtractedEmbedded: rip.embeddedSubtitles,
      };
    }
    // Single-pass failed (e.g. a broken track in container): fall back cleanly while preserving cached tracks.
    const audioSpeech = await resolveAudioSpeechForContext(context, resolved, deps);
    return {
      audioSpeech,
      sampledAudioWindows: null,
      preExtractedEmbedded: getCachedEmbeddedSubtitles(context.videoPath),
    };
  }

  if (cachedFull) {
    return { audioSpeech: cachedFull, sampledAudioWindows: null, preExtractedEmbedded: null };
  }

  const audioSpeech = await resolveAudioSpeechForContext(context, resolved, deps);
  return { audioSpeech, sampledAudioWindows: null, preExtractedEmbedded: null };
}

export function collectRankedRemoteEmbeddedTracks(
  context: SubsyncContext,
  targetTrack: MpvTrack,
): Array<{ track: MpvTrack; ffIndex: number; extension: string; key: string; label: string }> {
  const metadataEvaluated: Array<{
    evaluated: EvaluatedSubtitleReference;
    track: MpvTrack;
    ffIndex: number;
    extension: string;
  }> = [];

  for (const track of context.subtitleTracks) {
    if (typeof track.id !== 'number' || track.id === targetTrack.id) continue;
    if (track.external || isSignsOrSongsTrackMeta(track)) continue;
    const ffIndex = track['ff-index'];
    const extension = codecToExtension(track.codec);
    if (typeof ffIndex !== 'number' || !Number.isInteger(ffIndex) || ffIndex < 0 || !extension) {
      continue;
    }
    const isSecondary = context.secondaryTrack !== null && track.id === context.secondaryTrack.id;
    metadataEvaluated.push({
      track,
      ffIndex,
      extension,
      evaluated: {
        key: `track:${track.id}`,
        kind: 'track',
        trackId: track.id,
        filePath: null,
        label: buildTrackReferenceLabel(track, isSecondary),
        isSecondary,
        isEmbedded: true,
        isEnglish: isEnglishReferenceMeta({ lang: track.lang, title: track.title }),
        isFullDialogueHint: isFullDialogueHintMeta({ title: track.title }),
        isSignsOnly: false,
        dialogueCount: MIN_REFERENCE_DIALOGUE_LINES + 5,
        signCount: 0,
        audioStatus: 'unverified',
        audioOffsetMs: null,
        audioConfidence: 0,
      },
    });
  }

  const rankedKeys = rankReferenceCandidates(metadataEvaluated.map((item) => item.evaluated));
  const byKey = new Map(metadataEvaluated.map((item) => [item.evaluated.key, item]));
  return rankedKeys.map((ev) => {
    const item = byKey.get(ev.key)!;
    return {
      track: item.track,
      ffIndex: item.ffIndex,
      extension: item.extension,
      key: ev.key,
      label: ev.label,
    };
  });
}
