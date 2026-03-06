import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import type { Args, SubtitleCandidate, YoutubeSubgenOutputs } from './types.js';
import { YOUTUBE_SUB_EXTENSIONS, YOUTUBE_AUDIO_EXTENSIONS } from './types.js';
import { log } from './log.js';
import {
  resolvePathMaybe,
  uniqueNormalizedLangCodes,
  escapeRegExp,
  normalizeBasename,
  runExternalCommand,
  commandExists,
} from './util.js';
import { state } from './mpv.js';

function toYtdlpLangPattern(langCodes: string[]): string {
  return langCodes.map((lang) => `${lang}.*`).join(',');
}

function filenameHasLanguageTag(filenameLower: string, langCode: string): boolean {
  const escaped = escapeRegExp(langCode);
  const pattern = new RegExp(`(^|[._-])${escaped}([._-]|$)`);
  return pattern.test(filenameLower);
}

function classifyLanguage(
  filename: string,
  primaryLangCodes: string[],
  secondaryLangCodes: string[],
): 'primary' | 'secondary' | null {
  const lower = filename.toLowerCase();
  const primary = primaryLangCodes.some((code) => filenameHasLanguageTag(lower, code));
  const secondary = secondaryLangCodes.some((code) => filenameHasLanguageTag(lower, code));
  if (primary && !secondary) return 'primary';
  if (secondary && !primary) return 'secondary';
  return null;
}

function preferredLangLabel(langCodes: string[], fallback: string): string {
  return uniqueNormalizedLangCodes(langCodes)[0] || fallback;
}

function sourceTag(source: SubtitleCandidate['source']): string {
  if (source === 'manual' || source === 'auto') return `ytdlp-${source}`;
  if (source === 'whisper-translate') return 'whisper-translate';
  return 'whisper';
}

function pickBestCandidate(candidates: SubtitleCandidate[]): SubtitleCandidate | null {
  if (candidates.length === 0) return null;
  const scored = [...candidates].sort((a, b) => {
    const sourceA = a.source === 'manual' ? 1 : 0;
    const sourceB = b.source === 'manual' ? 1 : 0;
    if (sourceA !== sourceB) return sourceB - sourceA;
    const srtA = a.ext === '.srt' ? 1 : 0;
    const srtB = b.ext === '.srt' ? 1 : 0;
    if (srtA !== srtB) return srtB - srtA;
    return b.size - a.size;
  });
  return scored[0] ?? null;
}

function scanSubtitleCandidates(
  tempDir: string,
  knownSet: Set<string>,
  source: 'manual' | 'auto',
  primaryLangCodes: string[],
  secondaryLangCodes: string[],
): SubtitleCandidate[] {
  const entries = fs.readdirSync(tempDir);
  const out: SubtitleCandidate[] = [];
  for (const name of entries) {
    const fullPath = path.join(tempDir, name);
    if (knownSet.has(fullPath)) continue;
    let stat: fs.Stats;
    try {
      stat = fs.statSync(fullPath);
    } catch {
      continue;
    }
    if (!stat.isFile()) continue;
    const ext = path.extname(fullPath).toLowerCase();
    if (!YOUTUBE_SUB_EXTENSIONS.has(ext)) continue;
    const lang = classifyLanguage(name, primaryLangCodes, secondaryLangCodes);
    if (!lang) continue;
    out.push({ path: fullPath, lang, ext, size: stat.size, source });
  }
  return out;
}

async function convertToSrt(
  inputPath: string,
  tempDir: string,
  langLabel: string,
): Promise<string> {
  if (path.extname(inputPath).toLowerCase() === '.srt') return inputPath;
  const outputPath = path.join(tempDir, `converted.${langLabel}.srt`);
  await runExternalCommand('ffmpeg', ['-y', '-loglevel', 'error', '-i', inputPath, outputPath]);
  return outputPath;
}

function findAudioFile(tempDir: string, preferredExt: string): string | null {
  const entries = fs.readdirSync(tempDir);
  const audioFiles: Array<{ path: string; ext: string; mtimeMs: number }> = [];
  for (const name of entries) {
    const fullPath = path.join(tempDir, name);
    let stat: fs.Stats;
    try {
      stat = fs.statSync(fullPath);
    } catch {
      continue;
    }
    if (!stat.isFile()) continue;
    const ext = path.extname(name).toLowerCase();
    if (!YOUTUBE_AUDIO_EXTENSIONS.has(ext)) continue;
    audioFiles.push({ path: fullPath, ext, mtimeMs: stat.mtimeMs });
  }
  if (audioFiles.length === 0) return null;
  const preferred = audioFiles.find((entry) => entry.ext === `.${preferredExt.toLowerCase()}`);
  if (preferred) return preferred.path;
  audioFiles.sort((a, b) => b.mtimeMs - a.mtimeMs);
  return audioFiles[0]?.path ?? null;
}

async function runWhisper(
  whisperBin: string,
  modelPath: string,
  audioPath: string,
  language: string,
  translate: boolean,
  outputPrefix: string,
): Promise<string> {
  const args = [
    '-m',
    modelPath,
    '-f',
    audioPath,
    '--output-srt',
    '--output-file',
    outputPrefix,
    '--language',
    language,
  ];
  if (translate) args.push('--translate');
  await runExternalCommand(whisperBin, args, {
    commandLabel: 'whisper',
    streamOutput: true,
  });
  const outputPath = `${outputPrefix}.srt`;
  if (!fs.existsSync(outputPath)) {
    throw new Error(`whisper output not found: ${outputPath}`);
  }
  return outputPath;
}

async function convertAudioForWhisper(inputPath: string, tempDir: string): Promise<string> {
  const wavPath = path.join(tempDir, 'whisper-input.wav');
  await runExternalCommand('ffmpeg', [
    '-y',
    '-loglevel',
    'error',
    '-i',
    inputPath,
    '-ar',
    '16000',
    '-ac',
    '1',
    '-c:a',
    'pcm_s16le',
    wavPath,
  ]);
  if (!fs.existsSync(wavPath)) {
    throw new Error(`Failed to prepare whisper audio input: ${wavPath}`);
  }
  return wavPath;
}

export function resolveWhisperBinary(args: Args): string | null {
  const explicit = args.whisperBin.trim();
  if (explicit) return resolvePathMaybe(explicit);
  if (commandExists('whisper-cli')) return 'whisper-cli';
  return null;
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
  const ytdlpManualLangs = toYtdlpLangPattern([...primaryLangCodes, ...secondaryLangCodes]);

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
    log('debug', args.logLevel, `YouTube subtitle temp dir: ${tempDir}`);
    const meta = await runExternalCommand(
      'yt-dlp',
      ['--dump-single-json', '--no-warnings', target],
      {
        captureStdout: true,
        logLevel: args.logLevel,
        commandLabel: 'yt-dlp:meta',
      },
      state.youtubeSubgenChildren,
    );
    const metadata = JSON.parse(meta.stdout) as { id?: string };
    const videoId = metadata.id || `${Date.now()}`;
    const basename = normalizeBasename(videoId, videoId);

    await runExternalCommand(
      'yt-dlp',
      [
        '--skip-download',
        '--no-warnings',
        '--write-subs',
        '--sub-format',
        'srt/vtt/best',
        '--sub-langs',
        ytdlpManualLangs,
        '-o',
        path.join(tempDir, '%(id)s.%(ext)s'),
        target,
      ],
      {
        allowFailure: true,
        logLevel: args.logLevel,
        commandLabel: 'yt-dlp:manual-subs',
        streamOutput: true,
      },
      state.youtubeSubgenChildren,
    );

    const manualSubs = scanSubtitleCandidates(
      tempDir,
      knownFiles,
      'manual',
      primaryLangCodes,
      secondaryLangCodes,
    );
    for (const sub of manualSubs) knownFiles.add(sub.path);
    let primaryCandidates = manualSubs.filter((entry) => entry.lang === 'primary');
    let secondaryCandidates = manualSubs.filter((entry) => entry.lang === 'secondary');

    const missingAuto: string[] = [];
    if (primaryCandidates.length === 0) missingAuto.push(toYtdlpLangPattern(primaryLangCodes));
    if (secondaryCandidates.length === 0) missingAuto.push(toYtdlpLangPattern(secondaryLangCodes));

    if (missingAuto.length > 0) {
      await runExternalCommand(
        'yt-dlp',
        [
          '--skip-download',
          '--no-warnings',
          '--write-auto-subs',
          '--sub-format',
          'srt/vtt/best',
          '--sub-langs',
          missingAuto.join(','),
          '-o',
          path.join(tempDir, '%(id)s.%(ext)s'),
          target,
        ],
        {
          allowFailure: true,
          logLevel: args.logLevel,
          commandLabel: 'yt-dlp:auto-subs',
          streamOutput: true,
        },
        state.youtubeSubgenChildren,
      );

      const autoSubs = scanSubtitleCandidates(
        tempDir,
        knownFiles,
        'auto',
        primaryLangCodes,
        secondaryLangCodes,
      );
      for (const sub of autoSubs) knownFiles.add(sub.path);
      primaryCandidates = primaryCandidates.concat(
        autoSubs.filter((entry) => entry.lang === 'primary'),
      );
      secondaryCandidates = secondaryCandidates.concat(
        autoSubs.filter((entry) => entry.lang === 'secondary'),
      );
    }

    let primaryAlias = '';
    let secondaryAlias = '';
    const selectedPrimary = pickBestCandidate(primaryCandidates);
    const selectedSecondary = pickBestCandidate(secondaryCandidates);

    if (selectedPrimary) {
      const srt = await convertToSrt(selectedPrimary.path, tempDir, primaryLabel);
      primaryAlias = await publishTrack('primary', selectedPrimary.source, srt, basename);
    }
    if (selectedSecondary) {
      const srt = await convertToSrt(selectedSecondary.path, tempDir, secondaryLabel);
      secondaryAlias = await publishTrack('secondary', selectedSecondary.source, srt, basename);
    }

    const needsPrimaryWhisper = !selectedPrimary;
    const needsSecondaryWhisper = !selectedSecondary && secondaryCanUseWhisperTranslate;
    if (needsPrimaryWhisper || needsSecondaryWhisper) {
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
        try {
          await runExternalCommand(
            'yt-dlp',
            [
              '-f',
              'bestaudio/best',
              '--extract-audio',
              '--audio-format',
              args.youtubeSubgenAudioFormat,
              '--no-warnings',
              '-o',
              path.join(tempDir, '%(id)s.%(ext)s'),
              target,
            ],
            {
              logLevel: args.logLevel,
              commandLabel: 'yt-dlp:audio',
              streamOutput: true,
            },
            state.youtubeSubgenChildren,
          );
          const audioPath = findAudioFile(tempDir, args.youtubeSubgenAudioFormat);
          if (!audioPath) {
            throw new Error('Audio extraction succeeded, but no audio file was found.');
          }
          const whisperAudioPath = await convertAudioForWhisper(audioPath, tempDir);

          if (needsPrimaryWhisper) {
            try {
              const primaryPrefix = path.join(tempDir, `${basename}.${primaryLabel}`);
              const primarySrt = await runWhisper(
                whisperBin!,
                modelPath,
                whisperAudioPath,
                args.youtubeWhisperSourceLanguage,
                false,
                primaryPrefix,
              );
              primaryAlias = await publishTrack('primary', 'whisper', primarySrt, basename);
            } catch (error) {
              log(
                'warn',
                args.logLevel,
                `Failed to generate primary subtitle via whisper fallback: ${(error as Error).message}`,
              );
            }
          }

          if (needsSecondaryWhisper) {
            try {
              const secondaryPrefix = path.join(tempDir, `${basename}.${secondaryLabel}`);
              const secondarySrt = await runWhisper(
                whisperBin!,
                modelPath,
                whisperAudioPath,
                args.youtubeWhisperSourceLanguage,
                true,
                secondaryPrefix,
              );
              secondaryAlias = await publishTrack(
                'secondary',
                'whisper-translate',
                secondarySrt,
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
        } catch (error) {
          log(
            'warn',
            args.logLevel,
            `Whisper fallback pipeline failed: ${(error as Error).message}`,
          );
        }
      }
    }

    if (!secondaryCanUseWhisperTranslate && !selectedSecondary) {
      log(
        'warn',
        args.logLevel,
        `Secondary subtitle language (${secondaryLabel}) has no whisper translate fallback; relying on yt-dlp subtitles only.`,
      );
    }

    if (!primaryAlias && !secondaryAlias) {
      throw new Error('Failed to generate any subtitle tracks.');
    }
    if (!primaryAlias || !secondaryAlias) {
      log(
        'warn',
        args.logLevel,
        `Generated partial subtitle result: primary=${primaryAlias ? 'ok' : 'missing'}, secondary=${secondaryAlias ? 'ok' : 'missing'}`,
      );
    }

    return {
      basename,
      primaryPath: primaryAlias || undefined,
      secondaryPath: secondaryAlias || undefined,
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
