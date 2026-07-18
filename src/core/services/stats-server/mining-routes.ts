import type { Hono } from 'hono';
import { existsSync } from 'node:fs';
import { basename } from 'node:path';
import { AnkiConnectClient } from '../../../anki-connect.js';
import { getConfiguredWordFieldName } from '../../../anki-field-config.js';
import { resolveAnimatedImageLeadInSeconds } from '../../../anki-integration/animated-image-sync.js';
import { MediaGenerator } from '../../../media-generator.js';
import { statsJson } from '../../../types/stats-http-contract.js';
import {
  resolveRetimedSecondarySubtitleTextFromSidecar,
  resolveSecondarySubtitleTextFromSidecar,
} from '../secondary-subtitle-sidecar.js';
import {
  applyStatsWordAndSentenceCardFields,
  createStatsMiningContext,
  getStatsDirectMiningAudioFieldNames,
  getStatsWordMiningAudioFieldName,
  resolveStatsNoteFieldName,
  shouldUseStatsLapisKikuCardFields,
  statsMiningLogger,
  type StatsMiningRouteOptions,
  type StatsServerNoteInfo,
} from './mining-support.js';

export function registerStatsMiningRoutes(app: Hono, options?: StatsMiningRouteOptions): void {
  const {
    getAnkiConnectConfig,
    getEffectiveMiningDeckName,
    getSecondarySubtitleLanguages,
    getStatsMiningAlassPath,
    timeMiningPhase,
  } = createStatsMiningContext(options);
  app.post('/api/stats/mine-card', async (c) => {
    const body = await c.req.json().catch(() => null);
    const sourcePath = typeof body?.sourcePath === 'string' ? body.sourcePath.trim() : '';
    const startMs = typeof body?.startMs === 'number' ? body.startMs : NaN;
    const endMs = typeof body?.endMs === 'number' ? body.endMs : NaN;
    const sentence = typeof body?.sentence === 'string' ? body.sentence.trim() : '';
    const word = typeof body?.word === 'string' ? body.word.trim() : '';
    const bodySecondaryText =
      typeof body?.secondaryText === 'string' ? body.secondaryText.trim() : '';
    const videoTitle = typeof body?.videoTitle === 'string' ? body.videoTitle.trim() : '';
    const rawMode = c.req.query('mode');
    const mode = rawMode === 'audio' ? 'audio' : rawMode === 'word' ? 'word' : 'sentence';

    if (
      !sourcePath ||
      !sentence ||
      (mode === 'word' && !word) ||
      !Number.isFinite(startMs) ||
      !Number.isFinite(endMs)
    ) {
      return c.json(
        statsJson('mineCard', {
          error: 'sourcePath, sentence, startMs, and endMs are required',
        }),
        400,
      );
    }
    if (endMs <= startMs) {
      return c.json(statsJson('mineCard', { error: 'endMs must be greater than startMs' }), 400);
    }

    if (!existsSync(sourcePath)) {
      return c.json(statsJson('mineCard', { error: 'File not found' }), 404);
    }

    const ankiConfig = getAnkiConnectConfig();
    if (!ankiConfig) {
      return c.json(statsJson('mineCard', { error: 'AnkiConnect is not configured' }), 500);
    }
    const addYomitanNote = options?.addYomitanNote;
    if (mode === 'word' && !addYomitanNote) {
      return c.json(statsJson('mineCard', { error: 'Yomitan bridge not available' }), 500);
    }
    const secondarySubtitleLanguages = getSecondarySubtitleLanguages();
    let retimedSecondaryText = '';
    if (mode === 'sentence' && !bodySecondaryText) {
      try {
        retimedSecondaryText = await (
          options?.resolveRetimedSecondarySubtitleText ??
          resolveRetimedSecondarySubtitleTextFromSidecar
        )({
          sourcePath,
          startMs,
          endMs,
          languages: secondarySubtitleLanguages,
          alassPath: getStatsMiningAlassPath(),
        });
      } catch (error) {
        statsMiningLogger.warn(
          'Failed to resolve retimed secondary subtitle for stats mining:',
          error instanceof Error ? error.message : String(error),
        );
      }
    }
    const secondaryText =
      bodySecondaryText ||
      retimedSecondaryText ||
      resolveSecondarySubtitleTextFromSidecar({
        sourcePath,
        startMs,
        endMs,
        languages: secondarySubtitleLanguages,
      });

    const client = new AnkiConnectClient(ankiConfig.url ?? 'http://127.0.0.1:8765');
    const mediaGen = options?.createMediaGenerator?.() ?? new MediaGenerator();

    const audioPadding = ankiConfig.media?.audioPadding ?? 0;
    const normalizeAudio = ankiConfig.media?.normalizeAudio !== false;
    const maxMediaDuration = ankiConfig.media?.maxMediaDuration ?? 30;

    const startSec = startMs / 1000;
    const endSec = endMs / 1000;
    const rawDuration = endSec - startSec;
    const clampedEndSec = rawDuration > maxMediaDuration ? startSec + maxMediaDuration : endSec;

    const highlightedSentence = word
      ? sentence.replace(
          new RegExp(word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'),
          `<b>${word}</b>`,
        )
      : sentence;

    const generateAudio = ankiConfig.media?.generateAudio !== false;
    const generateImage = ankiConfig.media?.generateImage !== false && mode !== 'audio';
    const imageType = ankiConfig.media?.imageType ?? 'static';
    const syncAnimatedImageToWordAudio =
      imageType === 'avif' && ankiConfig.media?.syncAnimatedImageToWordAudio !== false;

    const audioPromise = generateAudio
      ? timeMiningPhase(mode, 'generateAudio', () =>
          mediaGen.generateAudio(
            sourcePath,
            startSec,
            clampedEndSec,
            audioPadding,
            null,
            normalizeAudio,
          ),
        )
      : Promise.resolve(null);

    const createImagePromise = (animatedLeadInSeconds = 0): Promise<Buffer | null> => {
      if (!generateImage) {
        return Promise.resolve(null);
      }

      if (imageType === 'avif') {
        return timeMiningPhase(mode, 'generateAnimatedImage', () =>
          mediaGen.generateAnimatedImage(sourcePath, startSec, clampedEndSec, audioPadding, {
            fps: ankiConfig.media?.animatedFps ?? 10,
            maxWidth: ankiConfig.media?.animatedMaxWidth ?? 640,
            maxHeight: ankiConfig.media?.animatedMaxHeight,
            crf: ankiConfig.media?.animatedCrf ?? 35,
            leadingStillDuration: animatedLeadInSeconds,
          }),
        );
      }

      const midpointSec = (startSec + clampedEndSec) / 2;
      return timeMiningPhase(mode, 'generateScreenshot', () =>
        mediaGen.generateScreenshot(sourcePath, midpointSec, {
          format: ankiConfig.media?.imageFormat ?? 'jpg',
          quality: ankiConfig.media?.imageQuality ?? 92,
          maxWidth: ankiConfig.media?.imageMaxWidth,
          maxHeight: ankiConfig.media?.imageMaxHeight,
        }),
      );
    };

    const imagePromise =
      mode === 'word' && syncAnimatedImageToWordAudio
        ? Promise.resolve<Buffer | null>(null)
        : createImagePromise();

    const errors: string[] = [];
    let noteId: number;
    let effectiveDeckNamePromise: Promise<string> | null = null;
    const getEffectiveDeckNameForRequest = (): Promise<string> => {
      effectiveDeckNamePromise ??= getEffectiveMiningDeckName(ankiConfig);
      return effectiveDeckNamePromise;
    };
    const moveNoteToConfiguredDeck = async (id: number): Promise<void> => {
      const deckName = await getEffectiveDeckNameForRequest();
      if (!deckName) {
        return;
      }
      try {
        const cardIds = await timeMiningPhase(mode, 'findCards', () =>
          client.findCards(`nid:${id}`),
        );
        await timeMiningPhase(mode, 'changeDeck', () => client.changeDeck(cardIds, deckName));
      } catch (err) {
        errors.push(`deck: ${(err as Error).message}`);
      }
    };

    if (mode === 'word') {
      const [yomitanResult, audioResult, imageResult] = await Promise.allSettled([
        timeMiningPhase(
          'word',
          'addYomitanNote',
          () => addYomitanNote!(word),
          (noteId) => (typeof noteId === 'number' ? { noteId } : {}),
        ),
        audioPromise,
        imagePromise,
      ]);

      if (yomitanResult.status === 'rejected' || !yomitanResult.value) {
        return c.json(
          statsJson('mineCard', {
            error: `Yomitan failed to create note: ${yomitanResult.status === 'rejected' ? (yomitanResult.reason as Error).message : 'no result'}`,
          }),
          502,
        );
      }

      noteId = yomitanResult.value;
      await moveNoteToConfiguredDeck(noteId);
      const audioBuffer = audioResult.status === 'fulfilled' ? audioResult.value : null;
      if (audioResult.status === 'rejected')
        errors.push(`audio: ${(audioResult.reason as Error).message}`);
      if (imageResult.status === 'rejected')
        errors.push(`image: ${(imageResult.reason as Error).message}`);

      let imageBuffer = imageResult.status === 'fulfilled' ? imageResult.value : null;
      let noteInfo: StatsServerNoteInfo | null = null;
      if (
        audioBuffer ||
        (syncAnimatedImageToWordAudio && generateImage) ||
        shouldUseStatsLapisKikuCardFields(ankiConfig)
      ) {
        try {
          const noteInfoResult = (await client.notesInfo([noteId])) as StatsServerNoteInfo[];
          noteInfo = noteInfoResult[0] ?? null;
        } catch (err) {
          if (syncAnimatedImageToWordAudio && generateImage) {
            errors.push(`image: ${(err as Error).message}`);
          }
        }
      }
      if (syncAnimatedImageToWordAudio && generateImage) {
        try {
          const animatedLeadInSeconds = noteInfo
            ? await resolveAnimatedImageLeadInSeconds({
                config: ankiConfig,
                noteInfo,
                resolveConfiguredFieldName: (candidateNoteInfo, ...preferredNames) =>
                  resolveStatsNoteFieldName(candidateNoteInfo, ...preferredNames),
                retrieveMediaFileBase64: (filename) => client.retrieveMediaFile(filename),
              })
            : 0;
          imageBuffer = await createImagePromise(animatedLeadInSeconds);
        } catch (err) {
          errors.push(`image: ${(err as Error).message}`);
        }
      }
      if (generateAudio && !audioBuffer && audioResult.status === 'fulfilled') {
        errors.push('audio: no audio generated');
      }
      if (generateImage && !imageBuffer) {
        errors.push('image: no image generated');
      }

      const mediaFields: Record<string, string> = {};
      const timestamp = Date.now();
      const sentenceFieldName = ankiConfig.fields?.sentence ?? 'Sentence';
      const audioFieldName = getStatsWordMiningAudioFieldName(ankiConfig, noteInfo);
      const imageFieldName = ankiConfig.fields?.image ?? 'Picture';

      mediaFields[sentenceFieldName] = highlightedSentence;
      applyStatsWordAndSentenceCardFields(mediaFields, noteInfo, ankiConfig);

      if (audioBuffer) {
        const audioFilename = `subminer_audio_${timestamp}_${noteId}.mp3`;
        try {
          await timeMiningPhase('word', 'uploadAudio', () =>
            client.storeMediaFile(audioFilename, audioBuffer),
          );
          mediaFields[audioFieldName] = `[sound:${audioFilename}]`;
        } catch (err) {
          errors.push(`audio upload: ${(err as Error).message}`);
        }
      }

      if (imageBuffer) {
        const imageExt = imageType === 'avif' ? 'avif' : (ankiConfig.media?.imageFormat ?? 'jpg');
        const imageFilename = `subminer_image_${timestamp}_${noteId}.${imageExt}`;
        try {
          await timeMiningPhase('word', 'uploadImage', () =>
            client.storeMediaFile(imageFilename, imageBuffer),
          );
          mediaFields[imageFieldName] = `<img src="${imageFilename}">`;
        } catch (err) {
          errors.push(`image upload: ${(err as Error).message}`);
        }
      }

      const miscInfoFieldName = ankiConfig.fields?.miscInfo ?? '';
      if (miscInfoFieldName) {
        const pattern = ankiConfig.metadata?.pattern ?? '[SubMiner] %f (%t)';
        const filenameWithExt = videoTitle || basename(sourcePath);
        const filenameWithoutExt = filenameWithExt.replace(/\.[^.]+$/, '');
        const totalMs = Math.floor(startMs);
        const totalSec2 = Math.floor(totalMs / 1000);
        const hours = String(Math.floor(totalSec2 / 3600)).padStart(2, '0');
        const minutes = String(Math.floor((totalSec2 % 3600) / 60)).padStart(2, '0');
        const secs = String(totalSec2 % 60).padStart(2, '0');
        const ms = String(totalMs % 1000).padStart(3, '0');
        mediaFields[miscInfoFieldName] = pattern
          .replace(/%f/g, filenameWithoutExt)
          .replace(/%F/g, filenameWithExt)
          .replace(/%t/g, `${hours}:${minutes}:${secs}`)
          .replace(/%T/g, `${hours}:${minutes}:${secs}:${ms}`)
          .replace(/<br>/g, '\n');
      }

      if (Object.keys(mediaFields).length > 0) {
        try {
          await timeMiningPhase('word', 'updateNoteFields', () =>
            client.updateNoteFields(noteId, mediaFields),
          );
        } catch (err) {
          errors.push(`update fields: ${(err as Error).message}`);
        }
      }

      return c.json(statsJson('mineCard', { noteId, ...(errors.length > 0 ? { errors } : {}) }));
    }

    const wordFieldName = getConfiguredWordFieldName(ankiConfig);
    const sentenceFieldName = ankiConfig.fields?.sentence ?? 'Sentence';
    const translationFieldName = ankiConfig.fields?.translation ?? 'SelectionText';
    const imageFieldName = ankiConfig.fields?.image ?? 'Picture';
    const miscInfoFieldName = ankiConfig.fields?.miscInfo ?? '';

    const fields: Record<string, string> = {
      [sentenceFieldName]: mode === 'sentence' ? sentence : highlightedSentence,
    };

    if (mode === 'sentence' && secondaryText) {
      fields[translationFieldName] = secondaryText;
    }

    if (ankiConfig.isLapis?.enabled || ankiConfig.isKiku?.enabled) {
      if (mode === 'sentence') {
        fields[wordFieldName] = sentence;
      } else if (word) {
        fields[wordFieldName] = word;
      }
      if (mode === 'sentence') {
        fields['IsSentenceCard'] = 'x';
      } else if (mode === 'audio') {
        fields['IsAudioCard'] = 'x';
      }
    }

    const model = ankiConfig.isLapis?.sentenceCardModel || 'Basic';
    const tags = ankiConfig.tags ?? ['SubMiner'];

    const addNotePromise = timeMiningPhase(
      mode,
      'addNote',
      async () =>
        client.addNote((await getEffectiveDeckNameForRequest()) || 'Default', model, fields, tags),
      (id) => ({
        noteId: id,
      }),
    );

    const [audioResult, imageResult, addNoteResult] = await Promise.allSettled([
      audioPromise,
      imagePromise,
      addNotePromise,
    ]);

    const audioBuffer = audioResult.status === 'fulfilled' ? audioResult.value : null;
    const imageBuffer = imageResult.status === 'fulfilled' ? imageResult.value : null;
    if (audioResult.status === 'rejected')
      errors.push(`audio: ${(audioResult.reason as Error).message}`);
    if (imageResult.status === 'rejected')
      errors.push(`image: ${(imageResult.reason as Error).message}`);

    if (addNoteResult.status === 'rejected') {
      return c.json(
        statsJson('mineCard', {
          error: `Failed to add note: ${(addNoteResult.reason as Error).message}`,
        }),
        502,
      );
    }
    noteId = addNoteResult.value;
    await moveNoteToConfiguredDeck(noteId);

    const mediaFields: Record<string, string> = {};
    const timestamp = Date.now();
    let noteInfo: StatsServerNoteInfo | null = null;
    if (audioBuffer) {
      try {
        const noteInfoResult = (await client.notesInfo([noteId])) as StatsServerNoteInfo[];
        noteInfo = noteInfoResult[0] ?? null;
      } catch {
        noteInfo = null;
      }
    }

    if (audioBuffer) {
      const audioFilename = `subminer_audio_${timestamp}_${noteId}.mp3`;
      try {
        await timeMiningPhase(mode, 'uploadAudio', () =>
          client.storeMediaFile(audioFilename, audioBuffer),
        );
        const audioValue = `[sound:${audioFilename}]`;
        for (const fieldName of getStatsDirectMiningAudioFieldNames(ankiConfig, noteInfo, mode)) {
          mediaFields[fieldName] = audioValue;
        }
      } catch (err) {
        errors.push(`audio upload: ${(err as Error).message}`);
      }
    }

    if (imageBuffer) {
      const imageExt = imageType === 'avif' ? 'avif' : (ankiConfig.media?.imageFormat ?? 'jpg');
      const imageFilename = `subminer_image_${timestamp}_${noteId}.${imageExt}`;
      try {
        await timeMiningPhase(mode, 'uploadImage', () =>
          client.storeMediaFile(imageFilename, imageBuffer),
        );
        mediaFields[imageFieldName] = `<img src="${imageFilename}">`;
      } catch (err) {
        errors.push(`image upload: ${(err as Error).message}`);
      }
    }

    if (miscInfoFieldName) {
      const pattern = ankiConfig.metadata?.pattern ?? '[SubMiner] %f (%t)';
      const filenameWithExt = videoTitle || basename(sourcePath);
      const filenameWithoutExt = filenameWithExt.replace(/\.[^.]+$/, '');
      const totalMs = Math.floor(startMs);
      const totalSec = Math.floor(totalMs / 1000);
      const hours = String(Math.floor(totalSec / 3600)).padStart(2, '0');
      const minutes = String(Math.floor((totalSec % 3600) / 60)).padStart(2, '0');
      const secs = String(totalSec % 60).padStart(2, '0');
      const ms = String(totalMs % 1000).padStart(3, '0');
      const miscInfo = pattern
        .replace(/%f/g, filenameWithoutExt)
        .replace(/%F/g, filenameWithExt)
        .replace(/%t/g, `${hours}:${minutes}:${secs}`)
        .replace(/%T/g, `${hours}:${minutes}:${secs}:${ms}`)
        .replace(/<br>/g, '\n');
      mediaFields[miscInfoFieldName] = miscInfo;
    }

    if (Object.keys(mediaFields).length > 0) {
      try {
        await timeMiningPhase(mode, 'updateNoteFields', () =>
          client.updateNoteFields(noteId, mediaFields),
        );
      } catch (err) {
        errors.push(`update fields: ${(err as Error).message}`);
      }
    }

    return c.json(statsJson('mineCard', { noteId, ...(errors.length > 0 ? { errors } : {}) }));
  });
}
