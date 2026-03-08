import type { LauncherAiConfig } from '../types.js';
import { requestAiChatCompletion, resolveAiApiKey } from '../../src/ai/client.js';
import { parseSrt, stringifySrt, type SrtCue } from './srt.js';

const DEFAULT_SUBTITLE_FIX_PROMPT =
  'Fix transcription mistakes only. Preserve cue numbering, timestamps, and valid SRT formatting exactly. Return only corrected SRT.';

const SRT_BLOCK_PATTERN =
  /(?:^|\n)(\d+\n\d{2}:\d{2}:\d{2},\d{3} --> \d{2}:\d{2}:\d{2},\d{3}[\s\S]*)$/;
const CODE_FENCE_PATTERN = /^```(?:\w+)?\s*\n([\s\S]*?)\n```$/;
const JAPANESE_CHAR_PATTERN = /[\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Han}]/gu;
const LATIN_LETTER_PATTERN = /\p{Script=Latin}/gu;

export function applyFixedCueBatch(original: SrtCue[], fixed: SrtCue[]): SrtCue[] {
  if (original.length !== fixed.length) {
    throw new Error('Fixed subtitle batch must preserve cue count.');
  }

  return original.map((cue, index) => {
    const nextCue = fixed[index];
    if (!nextCue) {
      throw new Error('Missing fixed subtitle cue.');
    }
    if (cue.start !== nextCue.start || cue.end !== nextCue.end) {
      throw new Error('Fixed subtitle batch must preserve cue timestamps.');
    }
    return {
      ...cue,
      text: nextCue.text,
    };
  });
}

function chunkCues(cues: SrtCue[], size: number): SrtCue[][] {
  const chunks: SrtCue[][] = [];
  for (let index = 0; index < cues.length; index += size) {
    chunks.push(cues.slice(index, index + size));
  }
  return chunks;
}

function normalizeAiSubtitleFixCandidates(content: string): string[] {
  const trimmed = content.replace(/\r\n/g, '\n').trim();
  if (!trimmed) {
    return [];
  }

  const candidates = new Set<string>([trimmed]);
  const fenced = CODE_FENCE_PATTERN.exec(trimmed)?.[1]?.trim();
  if (fenced) {
    candidates.add(fenced);
  }

  const srtBlock = SRT_BLOCK_PATTERN.exec(trimmed)?.[1]?.trim();
  if (srtBlock) {
    candidates.add(srtBlock);
  }

  return [...candidates];
}

function parseTextOnlyCueBatch(original: SrtCue[], content: string): SrtCue[] {
  const paragraphBlocks = content
    .split(/\n{2,}/)
    .map((block) => block.trim())
    .filter((block) => block.length > 0);
  if (paragraphBlocks.length === original.length) {
    return original.map((cue, index) => ({
      ...cue,
      text: paragraphBlocks[index]!,
    }));
  }

  const lineBlocks = content
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  if (lineBlocks.length === original.length) {
    return original.map((cue, index) => ({
      ...cue,
      text: lineBlocks[index]!,
    }));
  }

  throw new Error('Fixed subtitle batch must preserve cue count.');
}

function countPatternMatches(content: string, pattern: RegExp): number {
  pattern.lastIndex = 0;
  return [...content.matchAll(pattern)].length;
}

function isJapaneseLanguageCode(language: string | undefined): boolean {
  if (!language) return false;
  const normalized = language.trim().toLowerCase();
  return normalized === 'ja' || normalized === 'jp' || normalized === 'jpn';
}

function validateExpectedLanguage(
  original: SrtCue[],
  fixed: SrtCue[],
  expectedLanguage: string | undefined,
): void {
  if (!isJapaneseLanguageCode(expectedLanguage)) return;

  const originalText = original.map((cue) => cue.text).join('\n');
  const fixedText = fixed.map((cue) => cue.text).join('\n');
  const originalJapaneseChars = countPatternMatches(originalText, JAPANESE_CHAR_PATTERN);
  if (originalJapaneseChars < 4) return;

  const fixedJapaneseChars = countPatternMatches(fixedText, JAPANESE_CHAR_PATTERN);
  const fixedLatinLetters = countPatternMatches(fixedText, LATIN_LETTER_PATTERN);
  if (fixedJapaneseChars === 0 && fixedLatinLetters >= 4) {
    throw new Error('Fixed subtitle batch changed language away from expected Japanese.');
  }
}

export function parseAiSubtitleFixResponse(
  original: SrtCue[],
  content: string,
  expectedLanguage?: string,
): SrtCue[] {
  const candidates = normalizeAiSubtitleFixCandidates(content);
  let lastError: Error | null = null;

  for (const candidate of candidates) {
    try {
      const parsed = parseSrt(candidate);
      validateExpectedLanguage(original, parsed, expectedLanguage);
      return parsed;
    } catch (error) {
      lastError = error as Error;
    }
  }

  for (const candidate of candidates) {
    try {
      const parsed = parseTextOnlyCueBatch(original, candidate);
      validateExpectedLanguage(original, parsed, expectedLanguage);
      return parsed;
    } catch (error) {
      lastError = error as Error;
    }
  }

  throw lastError ?? new Error('AI subtitle fix returned empty content.');
}

export async function fixSubtitleWithAi(
  subtitleContent: string,
  aiConfig: LauncherAiConfig,
  logWarning: (message: string) => void,
  expectedLanguage?: string,
): Promise<string | null> {
  if (aiConfig.enabled !== true) {
    return null;
  }

  const apiKey = await resolveAiApiKey(aiConfig);
  if (!apiKey) {
    return null;
  }

  const cues = parseSrt(subtitleContent);
  if (cues.length === 0) {
    return null;
  }

  const fixedChunks: SrtCue[] = [];
  for (const chunk of chunkCues(cues, 25)) {
    const fixedContent = await requestAiChatCompletion(
      {
        apiKey,
        baseUrl: aiConfig.baseUrl,
        model: aiConfig.model,
        timeoutMs: aiConfig.requestTimeoutMs,
        messages: [
          {
            role: 'system',
            content: aiConfig.systemPrompt?.trim() || DEFAULT_SUBTITLE_FIX_PROMPT,
          },
          {
            role: 'user',
            content: stringifySrt(chunk),
          },
        ],
      },
      {
        logWarning,
      },
    );
    if (!fixedContent) {
      return null;
    }

    let parsedFixed: SrtCue[];
    try {
      parsedFixed = parseAiSubtitleFixResponse(chunk, fixedContent, expectedLanguage);
    } catch (error) {
      logWarning(`AI subtitle fix returned invalid SRT: ${(error as Error).message}`);
      return null;
    }

    try {
      fixedChunks.push(...applyFixedCueBatch(chunk, parsedFixed));
    } catch (error) {
      logWarning(`AI subtitle fix validation failed: ${(error as Error).message}`);
      return null;
    }
  }

  return stringifySrt(fixedChunks);
}
