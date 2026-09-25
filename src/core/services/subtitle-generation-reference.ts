import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import type { SubtitleGenerationProgress } from '../../shared/subtitle-generation';
import { parseSrtCues } from './subtitle-cue-parser';
import { runSubtitleGenerationProcess } from './subtitle-generation-process';

export type SubtitleGenerationReference = {
  label: string;
  delaySeconds: number;
  source: { kind: 'embedded'; streamIndex: number } | { kind: 'external'; path: string };
};

// Check both titles and filenames: releases often tag only one of them.
const EXCLUDED =
  /(?:^|[^\p{L}\p{N}])(?:signs?(?:songs?)?|songs?|lyrics?|karaoke|forced|s[\s&+_-]*s|op|ed|opening|ending|generated)(?=$|[^\p{L}\p{N}])|看板|歌詞/iu;
const TEXT_CODECS = new Set(['ass', 'ssa', 'subrip', 'srt', 'webvtt', 'mov_text', 'text']);

/** Rank loaded dialogue tracks, preferring English, then explicitly full tracks. */
export function subtitleGenerationReferences(
  value: unknown,
  workingDirectory?: string,
  delays: ReadonlyMap<number, number> = new Map(),
): SubtitleGenerationReference[] {
  if (!Array.isArray(value)) return [];
  const tracks: unknown[] = value;
  return tracks
    .flatMap((track) => {
      if (typeof track !== 'object' || track === null || !('type' in track) || track.type !== 'sub')
        return [];
      const title = 'title' in track && typeof track.title === 'string' ? track.title : '';
      const filename =
        'external-filename' in track && typeof track['external-filename'] === 'string'
          ? track['external-filename']
          : '';
      const name = `${title} ${path.basename(filename)}`;
      if (('forced' in track && track.forced === true) || EXCLUDED.test(name)) return [];
      if ('codec' in track && typeof track.codec === 'string' && !TEXT_CODECS.has(track.codec))
        return [];
      let source: SubtitleGenerationReference['source'];
      if ('external' in track && track.external === true) {
        let local = filename;
        if (local.startsWith('file://')) {
          try {
            local = fileURLToPath(local);
          } catch {
            return [];
          }
        } else if (/^[a-z][a-z\d+.-]*:\/\//i.test(local)) return [];
        if (!local || (!path.isAbsolute(local) && !workingDirectory)) return [];
        if (!/\.(?:srt|ass|ssa|vtt)$/i.test(local)) return [];
        source = { kind: 'external', path: path.resolve(workingDirectory ?? '.', local) };
      } else {
        if (
          !('ff-index' in track) ||
          typeof track['ff-index'] !== 'number' ||
          !Number.isSafeInteger(track['ff-index']) ||
          track['ff-index'] < 0
        )
          return [];
        source = { kind: 'embedded', streamIndex: track['ff-index'] };
      }
      const language = 'lang' in track && typeof track.lang === 'string' ? track.lang : '';
      const english =
        /^(?:en|eng|english)(?:[-_]|$)/i.test(language) ||
        /(?:^|[\s.\[(_-])(?:en|eng|english)(?=$|[\s.\])_-])/i.test(name);
      const full = /\b(?:full|dialogue|dialog)\b/i.test(name);
      const selected = 'selected' in track && track.selected === true;
      const preferred = 'default' in track && track.default === true;
      return [
        {
          reference: {
            label:
              title ||
              path.basename(filename) ||
              `${language || 'Subtitle'} stream ${source.kind === 'embedded' ? source.streamIndex : ''}`,
            source,
            delaySeconds:
              'id' in track && typeof track.id === 'number' ? (delays.get(track.id) ?? 0) : 0,
          },
          score:
            Number(english) * 100 + Number(full) * 20 + Number(selected) * 2 + Number(preferred),
        },
      ];
    })
    .sort((a, b) => b.score - a.score)
    .map(({ reference }) => reference);
}

/** Read mpv's path base and active subtitle delays while capturing timing references. */
export async function readSubtitleGenerationReferences(
  tracks: unknown,
  requestProperty: (name: string) => Promise<unknown>,
): Promise<SubtitleGenerationReference[]> {
  const [directory, primary, secondary, primaryDelay, secondaryDelay] = await Promise.all(
    ['working-directory', 'sid', 'secondary-sid', 'sub-delay', 'secondary-sub-delay'].map((name) =>
      requestProperty(name).catch(() => null),
    ),
  );
  const delays = new Map<number, number>();
  for (const [id, delay] of [
    [primary, primaryDelay],
    [secondary, secondaryDelay],
  ]) {
    if (typeof id === 'number' && typeof delay === 'number' && Number.isFinite(delay))
      delays.set(id, delay);
  }
  return subtitleGenerationReferences(
    tracks,
    typeof directory === 'string' ? directory : undefined,
    delays,
  );
}

/** Reference timestamps are hints on the extracted audio timeline, never a coverage mask. */
export async function loadSubtitleGenerationReference(input: {
  references: readonly SubtitleGenerationReference[];
  mediaPath: string;
  ffmpegPath: string;
  directory: string;
  audioOffset: number;
  onProgress?: (progress: SubtitleGenerationProgress) => void;
  signal?: AbortSignal;
}): Promise<number[]> {
  for (const [index, reference] of input.references.entries()) {
    input.signal?.throwIfAborted();
    try {
      const output = path.join(input.directory, `reference-${index}.srt`);
      const embedded = reference.source.kind === 'embedded';
      await runSubtitleGenerationProcess({
        command: input.ffmpegPath,
        args: [
          '-nostdin',
          '-hide_banner',
          '-loglevel',
          'error',
          ...(embedded ? ['-copyts', '-start_at_zero'] : []),
          '-i',
          reference.source.kind === 'external' ? reference.source.path : input.mediaPath,
          '-map',
          reference.source.kind === 'embedded' ? `0:${reference.source.streamIndex}` : '0:s:0',
          '-c:s',
          'srt',
          output,
        ],
        signal: input.signal,
      });
      const starts = parseSrtCues(await readFile(output, 'utf8'))
        .filter((cue) => cue.text.trim() && !/[♪♫]/u.test(cue.text) && cue.endTime > cue.startTime)
        .map((cue) => cue.startTime + reference.delaySeconds - input.audioOffset)
        .filter((time) => Number.isFinite(time) && time >= 0);
      if (starts.length === 0) continue;
      input.onProgress?.({
        stage: 'extract',
        message: `Using subtitle timing reference: ${reference.label}`,
      });
      return [...new Set(starts)].sort((a, b) => a - b);
    } catch {
      input.signal?.throwIfAborted();
      // An optional reference must not prevent transcription. Try the next loaded track.
    }
  }
  if (input.references.length)
    input.onProgress?.({
      stage: 'extract',
      message: 'No readable subtitle timing reference. Using audio timing.',
    });
  return [];
}
