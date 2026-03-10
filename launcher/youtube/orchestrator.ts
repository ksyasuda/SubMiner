import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

import type { Args, SubtitleCandidate, YoutubeSubgenOutputs } from '../types.js';
import { log } from '../log.js';
import {
  commandExists,
  normalizeBasename,
  resolvePathMaybe,
  runExternalCommand,
  uniqueNormalizedLangCodes,
} from '../util.js';
import { state } from '../mpv.js';
import { downloadYoutubeAudio, convertAudioForWhisper } from './audio-extraction.js';
import {
  downloadManualSubtitles,
  pickBestCandidate,
  scanSubtitleCandidates,
  toYtdlpLangPattern,
} from './manual-subs.js';
import { runLoggedYoutubePhase } from './progress.js';
import { fixSubtitleWithAi } from './subtitle-fix-ai.js';
import { runWhisper } from './whisper.js';

export interface YoutubeSubtitleGenerationPlan {
  fetchManualSubtitles: true;
  fetchAutoSubtitles: false;
  publishPrimaryManualSubtitle: false;
  publishSecondaryManualSubtitle: false;
  generatePrimarySubtitle: boolean;
  generateSecondarySubtitle: boolean;
}

export function planYoutubeSubtitleGeneration(input: {
  hasPrimaryManualSubtitle: boolean;
  hasSecondaryManualSubtitle: boolean;
  secondaryCanTranslate: boolean;
}): YoutubeSubtitleGenerationPlan {
  return {
    fetchManualSubtitles: true,
    fetchAutoSubtitles: false,
    publishPrimaryManualSubtitle: false,
    publishSecondaryManualSubtitle: false,
    generatePrimarySubtitle: !input.hasPrimaryManualSubtitle,
    generateSecondarySubtitle: !input.hasSecondaryManualSubtitle && input.secondaryCanTranslate,
  };
}

function preferredLangLabel(langCodes: string[], fallback: string): string {
  return uniqueNormalizedLangCodes(langCodes)[0] || fallback;
}

function sourceTag(source: SubtitleCandidate['source']): string {
  return source;
}

export function resolveWhisperBinary(args: Args): string | null {
  const explicit = args.whisperBin.trim();
  if (explicit) return resolvePathMaybe(explicit);
  if (commandExists('whisper-cli')) return 'whisper-cli';
  return null;
}

async function maybeFixSubtitleWithAi(
  selectedPath: string,
  args: Args,
  expectedLanguage?: string,
): Promise<string> {
  if (!args.youtubeFixWithAi || args.aiConfig.enabled !== true) {
    return selectedPath;
  }
  const fixedContent = await runLoggedYoutubePhase(
    {
      startMessage: `Starting AI subtitle fix: ${path.basename(selectedPath)}`,
      finishMessage: `Finished AI subtitle fix: ${path.basename(selectedPath)}`,
      failureMessage: `AI subtitle fix failed: ${path.basename(selectedPath)}`,
      log: (level, message) => log(level, args.logLevel, message),
    },
    async () => {
      const originalContent = fs.readFileSync(selectedPath, 'utf8');
      return fixSubtitleWithAi(
        originalContent,
        args.aiConfig,
        (message) => {
          log('warn', args.logLevel, message);
        },
        expectedLanguage,
      );
    },
  );
  if (!fixedContent) {
    return selectedPath;
  }

  const fixedPath = selectedPath.replace(/\.srt$/i, '.fixed.srt');
  fs.writeFileSync(fixedPath, fixedContent, 'utf8');
  return fixedPath;
}

export async function generateYoutubeSubtitles(
  target: string,
  args: Args,
  onReady?: (lang: 'primary' | 'secondary', pathToLoad: string) => Promise<void>,
): Promise<YoutubeSubgenOutputs> {
  const outDir = path.resolve(resolvePathMaybe(args.youtubeSubgenOutDir));
  fs.mkdirSync(outDir, { recursive: true });

  const primaryLangCodes = uniqueNormalizedLangCodes(args.youtubePrimarySubLangs);
  const secondaryLangCodes = uniqueNormalizedLangCodes(args.youtubeSecondarySubLangs);
  const primaryLabel = preferredLangLabel(primaryLangCodes, 'primary');
  const secondaryLabel = preferredLangLabel(secondaryLangCodes, 'secondary');
  const secondaryCanUseWhisperTranslate =
    secondaryLangCodes.includes('en') || secondaryLangCodes.includes('eng');
  const manualLangs = toYtdlpLangPattern([...primaryLangCodes, ...secondaryLangCodes]);

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-yt-subgen-'));
  const knownFiles = new Set<string>();
  let keepTemp = args.youtubeSubgenKeepTemp;

  const publishTrack = async (
    lang: 'primary' | 'secondary',
    source: SubtitleCandidate['source'],
    selectedPath: string,
    basename: string,
  ): Promise<string> => {
    const langLabel = lang === 'primary' ? primaryLabel : secondaryLabel;
    const taggedPath = path.join(outDir, `${basename}.${langLabel}.${sourceTag(source)}.srt`);
    const aliasPath = path.join(outDir, `${basename}.${langLabel}.srt`);
    fs.copyFileSync(selectedPath, taggedPath);
    fs.copyFileSync(taggedPath, aliasPath);
    log('info', args.logLevel, `Generated subtitle (${langLabel}, ${source}) -> ${aliasPath}`);
    if (onReady) await onReady(lang, aliasPath);
    return aliasPath;
  };

  try {
    const meta = await runLoggedYoutubePhase(
      {
        startMessage: 'Starting YouTube metadata probe',
        finishMessage: 'Finished YouTube metadata probe',
        failureMessage: 'YouTube metadata probe failed',
        log: (level, message) => log(level, args.logLevel, message),
      },
      () =>
        runExternalCommand(
          'yt-dlp',
          ['--dump-single-json', '--no-warnings', target],
          {
            captureStdout: true,
            logLevel: args.logLevel,
            commandLabel: 'yt-dlp:meta',
          },
          state.youtubeSubgenChildren,
        ),
    );
    const metadata = JSON.parse(meta.stdout) as { id?: string };
    const videoId = metadata.id || `${Date.now()}`;
    const basename = normalizeBasename(videoId, videoId);

    await runLoggedYoutubePhase(
      {
        startMessage: `Starting manual subtitle probe (${manualLangs || 'requested langs'})`,
        finishMessage: 'Finished manual subtitle probe',
        failureMessage: 'Manual subtitle probe failed',
        log: (level, message) => log(level, args.logLevel, message),
      },
      () =>
        downloadManualSubtitles(
          target,
          tempDir,
          manualLangs,
          args.logLevel,
          state.youtubeSubgenChildren,
        ),
    );

    const manualSubs = scanSubtitleCandidates(
      tempDir,
      knownFiles,
      'manual',
      primaryLangCodes,
      secondaryLangCodes,
    );
    for (const sub of manualSubs) knownFiles.add(sub.path);
    const selectedPrimary = pickBestCandidate(
      manualSubs.filter((entry) => entry.lang === 'primary'),
    );
    const selectedSecondary = pickBestCandidate(
      manualSubs.filter((entry) => entry.lang === 'secondary'),
    );

    const plan = planYoutubeSubtitleGeneration({
      hasPrimaryManualSubtitle: Boolean(selectedPrimary),
      hasSecondaryManualSubtitle: Boolean(selectedSecondary),
      secondaryCanTranslate: secondaryCanUseWhisperTranslate,
    });

    let primaryAlias = '';
    let secondaryAlias = '';

    if (selectedPrimary) {
      log(
        'info',
        args.logLevel,
        `Using native YouTube subtitle track for primary (${primaryLabel}); skipping external subtitle copy.`,
      );
    }
    if (selectedSecondary) {
      log(
        'info',
        args.logLevel,
        `Using native YouTube subtitle track for secondary (${secondaryLabel}); skipping external subtitle copy.`,
      );
    }

    if (plan.generatePrimarySubtitle || plan.generateSecondarySubtitle) {
      const whisperBin = resolveWhisperBinary(args);
      const modelPath = args.whisperModel.trim()
        ? path.resolve(resolvePathMaybe(args.whisperModel.trim()))
        : '';
      const hasWhisperFallback = !!whisperBin && !!modelPath && fs.existsSync(modelPath);

      if (!hasWhisperFallback) {
        log(
          'warn',
          args.logLevel,
          'Whisper fallback is not configured; continuing with available subtitle tracks.',
        );
      } else {
        const audioPath = await runLoggedYoutubePhase(
          {
            startMessage: 'Starting fallback audio extraction for subtitle generation',
            finishMessage: 'Finished fallback audio extraction',
            failureMessage: 'Fallback audio extraction failed',
            log: (level, message) => log(level, args.logLevel, message),
          },
          () => downloadYoutubeAudio(target, args, tempDir, state.youtubeSubgenChildren),
        );
        const whisperAudioPath = await runLoggedYoutubePhase(
          {
            startMessage: 'Starting ffmpeg audio prep for whisper',
            finishMessage: 'Finished ffmpeg audio prep for whisper',
            failureMessage: 'ffmpeg audio prep for whisper failed',
            log: (level, message) => log(level, args.logLevel, message),
          },
          () => convertAudioForWhisper(audioPath, tempDir),
        );

        if (plan.generatePrimarySubtitle) {
          try {
            const primaryPrefix = path.join(tempDir, `${basename}.${primaryLabel}`);
            const primarySrt = await runLoggedYoutubePhase(
              {
                startMessage: `Starting whisper primary subtitle generation (${primaryLabel})`,
                finishMessage: `Finished whisper primary subtitle generation (${primaryLabel})`,
                failureMessage: `Whisper primary subtitle generation failed (${primaryLabel})`,
                log: (level, message) => log(level, args.logLevel, message),
              },
              () =>
                runWhisper(whisperBin!, args, {
                  modelPath,
                  audioPath: whisperAudioPath,
                  language: args.youtubeWhisperSourceLanguage,
                  translate: false,
                  outputPrefix: primaryPrefix,
                }),
            );
            const fixedPrimary = await maybeFixSubtitleWithAi(
              primarySrt,
              args,
              args.youtubeWhisperSourceLanguage,
            );
            primaryAlias = await publishTrack(
              'primary',
              fixedPrimary === primarySrt ? 'whisper' : 'whisper-fixed',
              fixedPrimary,
              basename,
            );
          } catch (error) {
            log(
              'warn',
              args.logLevel,
              `Failed to generate primary subtitle via whisper fallback: ${(error as Error).message}`,
            );
          }
        }

        if (plan.generateSecondarySubtitle) {
          try {
            const secondaryPrefix = path.join(tempDir, `${basename}.${secondaryLabel}`);
            const secondarySrt = await runLoggedYoutubePhase(
              {
                startMessage: `Starting whisper secondary subtitle generation (${secondaryLabel})`,
                finishMessage: `Finished whisper secondary subtitle generation (${secondaryLabel})`,
                failureMessage: `Whisper secondary subtitle generation failed (${secondaryLabel})`,
                log: (level, message) => log(level, args.logLevel, message),
              },
              () =>
                runWhisper(whisperBin!, args, {
                  modelPath,
                  audioPath: whisperAudioPath,
                  language: args.youtubeWhisperSourceLanguage,
                  translate: true,
                  outputPrefix: secondaryPrefix,
                }),
            );
            const fixedSecondary = await maybeFixSubtitleWithAi(secondarySrt, args);
            secondaryAlias = await publishTrack(
              'secondary',
              fixedSecondary === secondarySrt ? 'whisper-translate' : 'whisper-translate-fixed',
              fixedSecondary,
              basename,
            );
          } catch (error) {
            log(
              'warn',
              args.logLevel,
              `Failed to generate secondary subtitle via whisper fallback: ${(error as Error).message}`,
            );
          }
        }
      }
    }

    if (!secondaryCanUseWhisperTranslate && !selectedSecondary) {
      log(
        'warn',
        args.logLevel,
        `Secondary subtitle language (${secondaryLabel}) has no whisper translate fallback; relying on manual subtitles only.`,
      );
    }

    if (!primaryAlias && !secondaryAlias && !selectedPrimary && !selectedSecondary) {
      throw new Error('Failed to generate any subtitle tracks.');
    }
    if ((!primaryAlias && !selectedPrimary) || (!secondaryAlias && !selectedSecondary)) {
      log(
        'warn',
        args.logLevel,
        `Generated partial subtitle result: primary=${primaryAlias || selectedPrimary ? 'ok' : 'missing'}, secondary=${secondaryAlias || selectedSecondary ? 'ok' : 'missing'}`,
      );
    }

    return {
      basename,
      primaryPath: primaryAlias || undefined,
      secondaryPath: secondaryAlias || undefined,
      primaryNative: Boolean(selectedPrimary),
      secondaryNative: Boolean(selectedSecondary),
    };
  } catch (error) {
    keepTemp = true;
    throw error;
  } finally {
    if (keepTemp) {
      log('warn', args.logLevel, `Keeping subtitle temp dir: ${tempDir}`);
    } else {
      try {
        fs.rmSync(tempDir, { recursive: true, force: true });
      } catch {
        // ignore cleanup failures
      }
    }
  }
}
