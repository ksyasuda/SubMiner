import { redactUrlCredentials } from './subtitle-generation-process';
import { buildJellyfinEmbeddedSubtitleUrl } from '../../subsync/utils';
import type { AlignmentResult, AlignSubtitleOptions } from '../../subsync/alass-align';
import { extractSinglePassRipMedia, getMediaFileSignature } from '../../subsync/audio-speech';
import {
  decodeSubtitleBytesWithStatus,
  extractSpansFromParsedSubtitle,
  parseSubtitleFileContent,
  type ParsedSubtitleFile,
  type SubtitleTimeSpan,
} from '../../subsync/subtitle-lines';
import { isRemoteMediaPath } from '../../jimaku/utils';
import {
  buildManualPayload,
  collectRankedRemoteEmbeddedTracks,
  discoverAndEvaluateSubtitleCandidates,
  prepareRipCandidatesAndAudio,
} from './subsync-candidates';
import {
  gatherSubsyncContext,
  isPlaybackSessionUnchanged,
  parseOptionalNumber,
  resolveTargetSlot,
} from './subsync-context';
import {
  LOSSY_REWRITE_ERROR_MESSAGE,
  buildRetimedPath,
  cleanupTemporaryFile,
  extractSubtitleTrackToFile,
  loadSyncedSubtitle,
  readAndParseSubtitleFile,
  tryFetchJellyfinSubtitleBytes,
  writeRetimedContent,
} from './subsync-files';
import {
  AUDIO_SPEECH_LEAD_SEC,
  AUTO_AUDIO_CONFIDENCE_THRESHOLD,
  AUTO_REMOTE_CONFIDENCE_THRESHOLD,
  AUTO_SUBTITLE_CONFIDENCE_THRESHOLD,
  type FileExtractionResult,
  MIN_REFERENCE_DIALOGUE_LINES,
  STALE_AUTO_RESULT_MESSAGE,
  type TriggerSubsyncFromConfigDeps,
  formatAlignmentSummary,
  formatOffsetSummary,
  getMpvClientForSubsync,
  isAlreadyInSync,
  isExecutableAvailable,
  type CandidateWithSpans,
} from './subsync-shared';
import { alignSubtitleOffThread } from '../../subsync/align-worker';

type BestAttempt = { alignment: AlignmentResult; sourceLabel: string; referenceKey: string };

/** Split evaluated references by audio verdict; aligned/unverified keep only usable dialogue tracks. */
function groupCandidatesByAudioStatus(candidates: readonly CandidateWithSpans[]): {
  misaligned: CandidateWithSpans[];
  aligned: CandidateWithSpans[];
  unverified: CandidateWithSpans[];
} {
  const isUsableDialogue = (c: CandidateWithSpans): boolean =>
    !c.evaluated.isSignsOnly && c.dialogueSpans.length >= MIN_REFERENCE_DIALOGUE_LINES;
  return {
    misaligned: candidates.filter(
      (c) => c.evaluated.audioStatus === 'misaligned' && !c.evaluated.isSignsOnly,
    ),
    aligned: candidates.filter((c) => c.evaluated.audioStatus === 'aligned' && isUsableDialogue(c)),
    unverified: candidates.filter(
      (c) => c.evaluated.audioStatus === 'unverified' && isUsableDialogue(c),
    ),
  };
}

export async function runAutoSubsync(deps: TriggerSubsyncFromConfigDeps): Promise<void> {
  const client = getMpvClientForSubsync(deps);
  const context = await gatherSubsyncContext(client);
  const resolved = deps.getResolvedConfig();
  const isRemote = isRemoteMediaPath(context.videoPath);
  const initialMediaSignature = !isRemote ? getMediaFileSignature(context.videoPath) : null;
  const isMediaFileUnchanged = (): boolean =>
    isRemote || getMediaFileSignature(context.videoPath) === initialMediaSignature;

  const targetTrack = context.primaryTrack;
  const targetSlot = resolveTargetSlot(targetTrack, context);

  let targetExtraction: FileExtractionResult | null = null;
  try {
    try {
      targetExtraction = await extractSubtitleTrackToFile(
        resolved.ffmpegPath,
        context.videoPath,
        targetTrack,
        context.workingDirectory,
        deps,
      );
    } catch (extractError) {
      if (
        !(await isPlaybackSessionUnchanged(client, context, targetTrack, targetSlot)) ||
        !isMediaFileUnchanged()
      ) {
        deps.showMpvOsd(STALE_AUTO_RESULT_MESSAGE);
        return;
      }
      const payload = buildManualPayload(context, resolved, deps, {
        guidanceMessage: `Automatic sync could not read the active subtitle track (${redactUrlCredentials((extractError as Error).message)}). Choose an engine and subtitle tracks below.`,
      });
      deps.openManualPicker(payload);
      deps.showMpvOsd('Subsync: choose engine and subtitles');
      return;
    }
    const parsedTarget = readAndParseSubtitleFile(targetExtraction.path);
    const targetInfo = extractSpansFromParsedSubtitle(parsedTarget, { stripSigns: true });

    if (parsedTarget.events.length === 0 || targetInfo.spans.length === 0) {
      if (
        !(await isPlaybackSessionUnchanged(client, context, targetTrack, targetSlot)) ||
        !isMediaFileUnchanged()
      ) {
        deps.showMpvOsd(STALE_AUTO_RESULT_MESSAGE);
        return;
      }
      const payload = buildManualPayload(context, resolved, deps, {
        guidanceMessage:
          'Could not read dialogue cues from the active subtitle track. Choose an engine or reference below.',
      });
      deps.openManualPicker(payload);
      deps.showMpvOsd('Subsync: choose engine and subtitles');
      return;
    }

    const prepared = await prepareRipCandidatesAndAudio(context, targetTrack, resolved, deps, {
      allowSampledAudioFastPath: true,
    });
    deps.signal?.throwIfAborted();
    let audioSpeech = prepared.audioSpeech;
    let evaluatedCandidates = await discoverAndEvaluateSubtitleCandidates(
      context,
      targetTrack,
      resolved,
      audioSpeech,
      {
        sampledAudioWindows: prepared.sampledAudioWindows,
        preExtractedEmbedded: prepared.preExtractedEmbedded,
        // Defer embedded extraction on remote streams (try external tracks first) and on
        // local sampled-audio fast path (try external/sidecar tracks before full MKV demux).
        skipEmbeddedExtraction: isRemote || prepared.sampledAudioWindows !== null,
        fetchImpl: deps.fetchImpl,
        signal: deps.signal,
      },
    );

    const replaceTarget = resolved.replace !== false && !targetExtraction.temporary;
    const outputPath = buildRetimedPath(targetExtraction.path, replaceTarget);
    let lengthSec = context.durationSec ?? audioSpeech?.durationSec ?? 0;

    const applyAlignment = async (
      alignment: AlignmentResult,
      sourceLabel: string,
    ): Promise<void> => {
      const alreadyInSync = isAlreadyInSync(alignment);
      const delayProperty = targetSlot === 'secondary' ? 'secondary-sub-delay' : 'sub-delay';
      const activeDelay = alreadyInSync
        ? parseOptionalNumber(await client.requestProperty(delayProperty).catch(() => null))
        : null;
      if (
        !(await isPlaybackSessionUnchanged(client, context, targetTrack, targetSlot)) ||
        !isMediaFileUnchanged()
      ) {
        deps.showMpvOsd(STALE_AUTO_RESULT_MESSAGE);
        return;
      }
      if (alreadyInSync) {
        if (activeDelay !== null && Math.abs(activeDelay) > 0.005) {
          client.send({ command: ['set_property', delayProperty, 0] });
          deps.showMpvOsd(`Reset subtitle delay to 0s (already in sync with ${sourceLabel})`);
          return;
        }
        deps.showMpvOsd(`Subtitles are already in sync (${sourceLabel})`);
        return;
      }
      if (parsedTarget.lossless === false) {
        deps.showMpvOsd(`Subsync failed: ${LOSSY_REWRITE_ERROR_MESSAGE}`);
        return;
      }
      writeRetimedContent(outputPath, alignment.rewrittenContent, {
        replacesSource: replaceTarget,
        sourceStamp: parsedTarget.stamp,
      });
      await loadSyncedSubtitle(client, outputPath, targetSlot, targetTrack);
      deps.showMpvOsd(formatAlignmentSummary(sourceLabel, alignment));
    };

    // Only assigned inside tryReference; the cast keeps TS from narrowing it to null.
    let bestAttempt = null as BestAttempt | null;

    /**
     * Align the target to one reference. A confident result is applied (returns true);
     * otherwise it is remembered as the best attempt for the guided modal unless
     * `trackBest` is false.
     */
    const tryReference = async (
      referenceSpans: readonly SubtitleTimeSpan[],
      source: { label: string; key: string },
      options: Pick<AlignSubtitleOptions, 'calibration' | 'confidenceThreshold' | 'leadSec'>,
      trackBest = true,
    ): Promise<boolean> => {
      const alignment = await alignSubtitleOffThread(deps.signal, parsedTarget, referenceSpans, {
        stripSigns: true,
        allowSplits: true,
        lengthSec,
        ...options,
      });
      if (alignment.confident) {
        await applyAlignment(alignment, source.label);
        return true;
      }
      if (trackBest && (!bestAttempt || alignment.confidence > bestAttempt.alignment.confidence)) {
        bestAttempt = { alignment, sourceLabel: source.label, referenceKey: source.key };
      }
      return false;
    };
    const clean = (confidenceThreshold: number) =>
      ({ calibration: 'clean', confidenceThreshold }) as const;

    // Fast path: if sampled audio already verified a clean aligned subtitle reference
    // and target alignment succeeds with high confidence, finish immediately without
    // reading the full video audio or embedded container tracks!
    if (!audioSpeech && prepared.sampledAudioWindows) {
      for (const candidate of groupCandidatesByAudioStatus(evaluatedCandidates).aligned) {
        const applied = await tryReference(
          candidate.dialogueSpans,
          candidate.evaluated,
          clean(AUTO_SUBTITLE_CONFIDENCE_THRESHOLD),
          false,
        );
        if (applied) return;
      }

      // Sampled verification did not yield a confident aligned subtitle sync (e.g. all
      // external references were misaligned or sampled windows were inconclusive). Escalate
      // to a single full-episode pass for embedded tracks + complete audio verification.
      const escalated = await prepareRipCandidatesAndAudio(context, targetTrack, resolved, deps, {
        allowSampledAudioFastPath: false,
      });
      deps.signal?.throwIfAborted();
      audioSpeech = escalated.audioSpeech;
      lengthSec = context.durationSec ?? audioSpeech?.durationSec ?? lengthSec;
      evaluatedCandidates = await discoverAndEvaluateSubtitleCandidates(
        context,
        targetTrack,
        resolved,
        audioSpeech,
        {
          // Preserve sampled-audio verdicts (including misaligned references) whenever full audio is unavailable.
          sampledAudioWindows: audioSpeech ? null : prepared.sampledAudioWindows,
          preExtractedEmbedded: escalated.preExtractedEmbedded,
          fetchImpl: deps.fetchImpl,
          signal: deps.signal,
        },
      );
    }

    let {
      misaligned: misalignedCandidates,
      aligned: alignedCandidates,
      unverified: unverifiedDialogueCandidates,
    } = groupCandidatesByAudioStatus(evaluatedCandidates);

    // 1. Best path: an audio-verified aligned dialogue subtitle reference (loaded or unloaded!).
    for (const candidate of alignedCandidates) {
      const applied = await tryReference(
        candidate.dialogueSpans,
        candidate.evaluated,
        clean(AUTO_SUBTITLE_CONFIDENCE_THRESHOLD),
      );
      if (applied) return;
    }

    // 2. Fallback when no subtitle track is audio-aligned (none exist, only signs/songs, or all
    //    reference subtitles are shifted off the audio): sync directly to video audio speech!
    if (audioSpeech && audioSpeech.spans.length >= MIN_REFERENCE_DIALOGUE_LINES) {
      const audioLabel = audioSpeech.source === 'silero-vad' ? 'audio (Silero VAD)' : 'video audio';
      const applied = await tryReference(
        audioSpeech.spans,
        { label: audioLabel, key: 'video' },
        {
          calibration: audioSpeech.calibration,
          leadSec: AUDIO_SPEECH_LEAD_SEC,
          confidenceThreshold: AUTO_AUDIO_CONFIDENCE_THRESHOLD,
        },
      );
      if (applied) return;
    }

    // 3. When audio verification wasn't available (e.g. remote stream or no ffmpeg):
    //    try already-loaded external dialogue tracks first with a stricter confidence bar.
    if (!audioSpeech) {
      for (const candidate of unverifiedDialogueCandidates) {
        const applied = await tryReference(
          candidate.dialogueSpans,
          candidate.evaluated,
          clean(AUTO_REMOTE_CONFIDENCE_THRESHOLD),
        );
        if (applied) return;
      }

      // 3b. On remote streams where no external reference succeeded, fetch embedded tracks
      //     via Jellyfin's subtitle endpoint first (KB instead of GB), or fall back to a
      //     single ffmpeg pass across embedded candidates instead of N sequential downloads.
      if (isRemote) {
        const rankedEmbedded = collectRankedRemoteEmbeddedTracks(context, targetTrack);
        if (rankedEmbedded.length > 0) {
          const remoteExtracted = new Map<number, ParsedSubtitleFile>();
          const triedRemoteEmbedded = new Set<number>();

          const tryAlignRemoteEmbeddedTrack = async (
            item: (typeof rankedEmbedded)[number],
            parsed: ParsedSubtitleFile,
          ): Promise<boolean> => {
            remoteExtracted.set(item.ffIndex, parsed);
            triedRemoteEmbedded.add(item.ffIndex);

            const extractedSpans = extractSpansFromParsedSubtitle(parsed, { stripSigns: true });
            if (extractedSpans.dialogueCount < MIN_REFERENCE_DIALOGUE_LINES) return false;
            return tryReference(
              extractedSpans.spans,
              item,
              clean(AUTO_REMOTE_CONFIDENCE_THRESHOLD),
            );
          };

          for (const item of rankedEmbedded) {
            if (
              !buildJellyfinEmbeddedSubtitleUrl(context.videoPath, item.ffIndex, item.extension)
            ) {
              break;
            }
            const bytes = await tryFetchJellyfinSubtitleBytes(
              context.videoPath,
              item.ffIndex,
              item.extension,
              deps,
            );
            if (!bytes) continue;
            const decoded = decodeSubtitleBytesWithStatus(bytes);
            const parsed = parseSubtitleFileContent(
              decoded.text,
              `track_${item.ffIndex}.${item.extension}`,
            );
            if (parsed.events.length === 0) continue;
            if (await tryAlignRemoteEmbeddedTrack(item, parsed)) {
              return;
            }
          }

          const remainingItems = rankedEmbedded.filter(
            (item) => !remoteExtracted.has(item.ffIndex),
          );
          const remainingSpecs = remainingItems.map((item) => ({
            ffIndex: item.ffIndex,
            extension: item.extension,
          }));

          if (remainingSpecs.length > 0 && isExecutableAvailable(resolved.ffmpegPath, ['ffmpeg'])) {
            const singlePass = await extractSinglePassRipMedia({
              ffmpegPath: resolved.ffmpegPath,
              videoPath: context.videoPath,
              audioStreamIndex: context.audioStreamIndex,
              signal: deps.signal,
              embeddedTracks: remainingSpecs,
              includeAudio: false,
            });
            if (singlePass) {
              for (const item of remainingItems) {
                const parsed = singlePass.embeddedSubtitles.get(item.ffIndex);
                if (!parsed || parsed.events.length === 0) continue;
                if (await tryAlignRemoteEmbeddedTrack(item, parsed)) {
                  return;
                }
              }
            }

            // Retry any ranked track omitted by the combined pass (whether the pass returned null
            // or a partial map) individually in rank order, stopping on the first confident match.
            const missingItems = remainingItems.filter(
              (item) => !remoteExtracted.has(item.ffIndex),
            );
            for (const fallbackItem of missingItems) {
              let extraction: FileExtractionResult | null = null;
              let parsed: ParsedSubtitleFile | null = null;
              try {
                extraction = await extractSubtitleTrackToFile(
                  resolved.ffmpegPath,
                  context.videoPath,
                  fallbackItem.track,
                  context.workingDirectory,
                  deps,
                );
                parsed = readAndParseSubtitleFile(extraction.path);
              } catch {
                // Continue to lower-ranked sibling tracks if extracting or parsing this track fails.
                continue;
              } finally {
                cleanupTemporaryFile(extraction);
              }
              if (parsed.events.length === 0) continue;
              if (await tryAlignRemoteEmbeddedTrack(fallbackItem, parsed)) {
                return;
              }
            }
          }

          if (remoteExtracted.size > 0) {
            evaluatedCandidates = await discoverAndEvaluateSubtitleCandidates(
              context,
              targetTrack,
              resolved,
              null,
              {
                preExtractedEmbedded: remoteExtracted,
                skipEmbeddedExtraction: true,
                fetchImpl: deps.fetchImpl,
                signal: deps.signal,
              },
            );
            ({
              misaligned: misalignedCandidates,
              aligned: alignedCandidates,
              unverified: unverifiedDialogueCandidates,
            } = groupCandidatesByAudioStatus(evaluatedCandidates));

            for (const candidate of unverifiedDialogueCandidates) {
              if (!candidate.evaluated.isEmbedded) continue;
              const matchedItem = rankedEmbedded.find(
                (item) => item.key === candidate.evaluated.key,
              );
              if (matchedItem && triedRemoteEmbedded.has(matchedItem.ffIndex)) continue;
              const applied = await tryReference(
                candidate.dialogueSpans,
                candidate.evaluated,
                clean(AUTO_REMOTE_CONFIDENCE_THRESHOLD),
              );
              if (applied) return;
            }
          }
        }
      }
    }

    // 4. Automatic sync could not finish with high confidence -> open guided modal with explanation!
    const reasons: string[] = [];
    if (misalignedCandidates.length > 0 && alignedCandidates.length === 0) {
      const sample = misalignedCandidates[0]!.evaluated;
      const shiftStr =
        sample.audioOffsetMs !== null ? formatOffsetSummary(sample.audioOffsetMs) : '';
      reasons.push(
        `Reference subtitle "${sample.label}" appears misaligned with the audio${shiftStr ? ` (${shiftStr})` : ''}.`,
      );
    } else if (
      evaluatedCandidates.length > 0 &&
      alignedCandidates.length === 0 &&
      unverifiedDialogueCandidates.length === 0
    ) {
      reasons.push(
        'Available subtitle tracks look like signs/songs or have too few dialogue lines.',
      );
    } else if (evaluatedCandidates.length === 0 && !audioSpeech) {
      reasons.push('No reference subtitle track or local audio source was available.');
    }

    if (bestAttempt) {
      const pct = Math.round(bestAttempt.alignment.confidence * 100);
      reasons.push(
        `Best automatic match (${bestAttempt.sourceLabel}) had low confidence (${pct}%).`,
      );
    }
    reasons.push('Pick a reference, browse a subtitle file, or adjust sync options below.');

    if (
      !(await isPlaybackSessionUnchanged(client, context, targetTrack, targetSlot)) ||
      !isMediaFileUnchanged()
    ) {
      deps.showMpvOsd(STALE_AUTO_RESULT_MESSAGE);
      return;
    }

    const payload = buildManualPayload(context, resolved, deps, {
      evaluatedCandidates,
      audioSpeechAvailable: audioSpeech !== null,
      guidanceMessage: reasons.join(' '),
      defaultEngine: 'builtin',
      defaultReferenceKey: bestAttempt?.referenceKey ?? null,
    });
    deps.openManualPicker(payload);
    deps.showMpvOsd('Subsync needs guidance: opened options');
  } finally {
    cleanupTemporaryFile(targetExtraction);
  }
}
