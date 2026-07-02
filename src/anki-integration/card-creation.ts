import { DEFAULT_ANKI_CONNECT_CONFIG } from '../config';
import {
  getConfiguredWordFieldName,
  getPreferredWordValueFromExtractedFields,
} from '../anki-field-config';
import { AnkiConnectConfig } from '../types/anki';
import { createLogger } from '../logger';
import type { MediaInput } from '../media-input';
import { SubtitleTimingTracker } from '../subtitle-timing-tracker';
import { AiConfig } from '../types/integrations';
import { MpvClient } from '../types/runtime';
import { resolveSentenceBackText } from './ai';
import {
  resolveMediaGenerationInput,
  resolveAudioStreamIndexForMediaGeneration,
  type MediaGenerationInputResolverOptions,
} from './media-source';
import { shouldMarkWordAndSentenceCard } from './note-field-utils';
import type { PendingYoutubeMediaUpdate } from './pending-youtube-media';

const log = createLogger('anki').child('integration.card-creation');

function shouldGenerateAudio(config: AnkiConnectConfig): boolean {
  return config.media?.generateAudio !== false;
}

function shouldGenerateImage(config: AnkiConnectConfig): boolean {
  return config.media?.generateImage !== false;
}

function trimToNonEmptyString(value: unknown): string | null {
  if (typeof value !== 'string') {
    return null;
  }
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

export interface CardCreationNoteInfo {
  noteId: number;
  fields: Record<string, { value: string }>;
}

type CardKind = 'sentence' | 'audio' | 'word-and-sentence';

interface CardCreationClient {
  addNote(
    deck: string,
    modelName: string,
    fields: Record<string, string>,
    tags?: string[],
  ): Promise<number>;
  addTags(noteIds: number[], tags: string[]): Promise<void>;
  notesInfo(noteIds: number[]): Promise<unknown>;
  updateNoteFields(noteId: number, fields: Record<string, string>): Promise<void>;
  storeMediaFile(filename: string, data: Buffer): Promise<void>;
  findNotes(query: string, options?: { maxRetries?: number }): Promise<number[]>;
  retrieveMediaFile(filename: string): Promise<string>;
}

interface CardCreationMediaGenerator {
  generateAudio(
    path: MediaInput,
    startTime: number,
    endTime: number,
    audioPadding?: number,
    audioStreamIndex?: number,
    normalizeAudio?: boolean,
  ): Promise<Buffer | null>;
  generateScreenshot(
    path: MediaInput,
    timestamp: number,
    options: {
      format: 'jpg' | 'png' | 'webp';
      quality?: number;
      maxWidth?: number;
      maxHeight?: number;
    },
  ): Promise<Buffer | null>;
  generateAnimatedImage(
    path: MediaInput,
    startTime: number,
    endTime: number,
    audioPadding?: number,
    options?: {
      fps?: number;
      maxWidth?: number;
      maxHeight?: number;
      crf?: number;
      leadingStillDuration?: number;
    },
  ): Promise<Buffer | null>;
}

interface CardCreationDeps {
  getConfig: () => AnkiConnectConfig;
  getAiConfig: () => AiConfig;
  getTimingTracker: () => SubtitleTimingTracker;
  getMpvClient: () => MpvClient;
  getCachedMediaPath?: MediaGenerationInputResolverOptions['getCachedMediaPath'];
  shouldRequireRemoteMediaCache?: () => boolean;
  getYoutubeMediaSourceUrl?: () => Promise<string | null | undefined> | string | null | undefined;
  queuePendingYoutubeMediaUpdate?: (job: PendingYoutubeMediaUpdate) => void;
  getDeck?: () => string | undefined;
  client: CardCreationClient;
  mediaGenerator: CardCreationMediaGenerator;
  showOsdNotification: (text: string) => void;
  showUpdateResult: (message: string, success: boolean) => void;
  showStatusNotification: (message: string) => void;
  showNotification: (noteId: number, label: string | number, errorSuffix?: string) => Promise<void>;
  beginUpdateProgress: (initialMessage: string) => void;
  endUpdateProgress: () => void;
  withUpdateProgress: <T>(initialMessage: string, action: () => Promise<T>) => Promise<T>;
  resolveConfiguredFieldName: (
    noteInfo: CardCreationNoteInfo,
    ...preferredNames: (string | undefined)[]
  ) => string | null;
  resolveNoteFieldName: (noteInfo: CardCreationNoteInfo, preferredName?: string) => string | null;
  getAnimatedImageLeadInSeconds: (noteInfo: CardCreationNoteInfo) => Promise<number>;
  extractFields: (fields: Record<string, { value: string }>) => Record<string, string>;
  processSentence: (mpvSentence: string, noteFields: Record<string, string>) => string;
  setCardTypeFields: (
    updatedFields: Record<string, string>,
    availableFieldNames: string[],
    cardKind: CardKind,
  ) => void;
  mergeFieldValue: (existing: string, newValue: string, overwrite: boolean) => string;
  formatMiscInfoPattern: (fallbackFilename: string, startTimeSeconds?: number) => string;
  getEffectiveSentenceCardConfig: () => {
    model?: string;
    sentenceField: string;
    audioField: string;
    lapisEnabled: boolean;
    kikuEnabled: boolean;
    kikuFieldGrouping: 'auto' | 'manual' | 'disabled';
    kikuDeleteDuplicateInAuto: boolean;
  };
  getFallbackDurationSeconds: () => number;
  appendKnownWordsFromNoteInfo: (noteInfo: CardCreationNoteInfo) => void;
  isUpdateInProgress: () => boolean;
  setUpdateInProgress: (value: boolean) => void;
  trackLastAddedNoteId?: (noteId: number) => void;
  trackLastAddedDuplicateNoteIds?: (noteId: number, duplicateNoteIds: number[]) => void;
  findDuplicateNoteIds?: (expression: string, noteInfo: CardCreationNoteInfo) => Promise<number[]>;
  recordCardsMinedCallback?: (count: number, noteIds?: number[]) => void;
}

export class CardCreationService {
  constructor(private readonly deps: CardCreationDeps) {}

  private getMediaResolverOptions(): MediaGenerationInputResolverOptions {
    const options: MediaGenerationInputResolverOptions = {
      logDebug: (message) => log.debug(message),
    };
    if (this.deps.getCachedMediaPath) {
      options.getCachedMediaPath = this.deps.getCachedMediaPath;
    }
    if (this.deps.shouldRequireRemoteMediaCache?.()) {
      options.remoteCacheMode = 'required';
    }
    return options;
  }

  private getConfiguredAnkiTags(): string[] {
    const tags = this.deps.getConfig().tags;
    if (!Array.isArray(tags)) {
      return [];
    }
    return [...new Set(tags.map((tag) => tag.trim()).filter((tag) => tag.length > 0))];
  }

  private async addConfiguredTagsToNote(noteId: number): Promise<void> {
    const tags = this.getConfiguredAnkiTags();
    if (tags.length === 0) {
      return;
    }
    try {
      await this.deps.client.addTags([noteId], tags);
    } catch (error) {
      log.warn('Failed to add tags to card:', (error as Error).message);
    }
  }

  async updateLastAddedFromClipboard(clipboardText: string): Promise<void> {
    try {
      if (!clipboardText || !clipboardText.trim()) {
        this.deps.showOsdNotification('Clipboard is empty');
        return;
      }

      const mpvClient = this.deps.getMpvClient();
      if (!mpvClient || !mpvClient.currentVideoPath) {
        this.deps.showOsdNotification('No video loaded');
        return;
      }

      const blocks = clipboardText
        .split(/\n\s*\n/)
        .map((block) => block.trim())
        .filter((block) => block.length > 0);

      if (blocks.length === 0) {
        this.deps.showOsdNotification('No subtitle blocks found in clipboard');
        return;
      }

      const timings: { startTime: number; endTime: number }[] = [];
      const timingTracker = this.deps.getTimingTracker();
      for (const block of blocks) {
        const timing = timingTracker.findTiming(block);
        if (timing) {
          timings.push(timing);
        }
      }

      if (timings.length === 0) {
        this.deps.showOsdNotification('Subtitle timing not found; copy again while playing');
        return;
      }

      const rangeStart = Math.min(...timings.map((entry) => entry.startTime));
      let rangeEnd = Math.max(...timings.map((entry) => entry.endTime));

      const maxMediaDuration = this.deps.getConfig().media?.maxMediaDuration ?? 30;
      if (maxMediaDuration > 0 && rangeEnd - rangeStart > maxMediaDuration) {
        log.warn(
          `Media range ${(rangeEnd - rangeStart).toFixed(1)}s exceeds cap of ${maxMediaDuration}s, clamping`,
        );
        rangeEnd = rangeStart + maxMediaDuration;
      }

      this.deps.showOsdNotification('Updating card from clipboard...');
      this.deps.beginUpdateProgress('Updating card from clipboard');
      this.deps.setUpdateInProgress(true);

      try {
        const deck = this.deps.getDeck?.() ?? this.deps.getConfig().deck;
        const query = deck ? `"deck:${deck}" added:1` : 'added:1';
        const noteIds = (await this.deps.client.findNotes(query, {
          maxRetries: 0,
        })) as number[];
        if (!noteIds || noteIds.length === 0) {
          this.deps.showOsdNotification('No recently added cards found');
          return;
        }

        const noteId = Math.max(...noteIds);
        const notesInfoResult = (await this.deps.client.notesInfo([
          noteId,
        ])) as CardCreationNoteInfo[];
        if (!notesInfoResult || notesInfoResult.length === 0) {
          this.deps.showOsdNotification('Card not found');
          return;
        }

        const noteInfo = notesInfoResult[0]!;
        const fields = this.deps.extractFields(noteInfo.fields);
        const expressionText = getPreferredWordValueFromExtractedFields(
          fields,
          this.deps.getConfig(),
        );
        const sentenceAudioField = this.getResolvedSentenceOnlyAudioFieldName(noteInfo);
        const sentenceCardConfig = this.deps.getEffectiveSentenceCardConfig();
        const sentenceField = sentenceCardConfig.sentenceField;

        const sentence = blocks.join(' ');
        const updatedFields: Record<string, string> = {};
        let updatePerformed = false;
        const errors: string[] = [];
        let miscInfoFilename: string | null = null;

        if (sentenceField) {
          const processedSentence = this.deps.processSentence(sentence, fields);
          updatedFields[sentenceField] = processedSentence;
          if (shouldMarkWordAndSentenceCard(noteInfo, sentenceCardConfig)) {
            this.deps.setCardTypeFields(
              updatedFields,
              Object.keys(noteInfo.fields),
              'word-and-sentence',
            );
          }
          updatePerformed = true;
        }

        log.info(
          `Clipboard update: timing range ${rangeStart.toFixed(2)}s - ${rangeEnd.toFixed(2)}s`,
        );

        const config = this.deps.getConfig();
        const generateAudio = shouldGenerateAudio(config);
        const generateImage = shouldGenerateImage(config);
        const mediaResolverOptions = this.getMediaResolverOptions();
        const audioSourcePath = generateAudio
          ? await resolveMediaGenerationInput(mpvClient, 'audio', mediaResolverOptions)
          : null;
        const videoPath = generateImage
          ? await resolveMediaGenerationInput(mpvClient, 'video', mediaResolverOptions)
          : null;

        if (generateAudio) {
          try {
            const audioFilename = this.generateAudioFilename();
            const audioBuffer = audioSourcePath
              ? await this.mediaGenerateAudio(audioSourcePath, rangeStart, rangeEnd)
              : null;

            if (audioBuffer) {
              await this.deps.client.storeMediaFile(audioFilename, audioBuffer);
              if (sentenceAudioField) {
                const audioValue = `[sound:${audioFilename}]`;
                const existingAudio = noteInfo.fields[sentenceAudioField]?.value || '';
                // Manual clipboard updates intentionally replace old captured sentence audio.
                updatedFields[sentenceAudioField] = this.deps.mergeFieldValue(
                  existingAudio,
                  audioValue,
                  true,
                );
              }
              miscInfoFilename = audioFilename;
              updatePerformed = true;
            }
          } catch (error) {
            log.error('Failed to generate audio:', (error as Error).message);
            errors.push('audio');
          }
        }

        if (generateImage) {
          try {
            const animatedLeadInSeconds = await this.deps.getAnimatedImageLeadInSeconds(noteInfo);
            const imageFilename = this.generateImageFilename();
            const imageBuffer = videoPath
              ? await this.generateImageBuffer(
                  videoPath,
                  rangeStart,
                  rangeEnd,
                  animatedLeadInSeconds,
                )
              : null;

            if (imageBuffer) {
              await this.deps.client.storeMediaFile(imageFilename, imageBuffer);
              const imageFieldName = this.deps.resolveConfiguredFieldName(
                noteInfo,
                this.deps.getConfig().fields?.image,
                DEFAULT_ANKI_CONNECT_CONFIG.fields.image,
              );
              if (!imageFieldName) {
                log.warn('Image field not found on note, skipping image update');
              } else {
                const existingImage = noteInfo.fields[imageFieldName]?.value || '';
                updatedFields[imageFieldName] = this.deps.mergeFieldValue(
                  existingImage,
                  `<img src="${imageFilename}">`,
                  this.deps.getConfig().behavior?.overwriteImage !== false,
                );
                miscInfoFilename = imageFilename;
                updatePerformed = true;
              }
            }
          } catch (error) {
            log.error('Failed to generate image:', (error as Error).message);
            errors.push('image');
          }
        }

        if (this.deps.getConfig().fields?.miscInfo) {
          const miscInfo = this.deps.formatMiscInfoPattern(miscInfoFilename || '', rangeStart);
          const miscInfoField = this.deps.resolveConfiguredFieldName(
            noteInfo,
            this.deps.getConfig().fields?.miscInfo,
          );
          if (miscInfo && miscInfoField) {
            updatedFields[miscInfoField] = miscInfo;
            updatePerformed = true;
          }
        }

        if (updatePerformed) {
          await this.deps.client.updateNoteFields(noteId, updatedFields);
          await this.addConfiguredTagsToNote(noteId);
          const label = expressionText || noteId;
          log.info('Updated card from clipboard:', label);
          const errorSuffix = errors.length > 0 ? `${errors.join(', ')} failed` : undefined;
          await this.deps.showNotification(noteId, label, errorSuffix);
        }
      } finally {
        this.deps.setUpdateInProgress(false);
        this.deps.endUpdateProgress();
      }
    } catch (error) {
      log.error('Error updating card from clipboard:', (error as Error).message);
      this.deps.showOsdNotification(`Update failed: ${(error as Error).message}`);
    }
  }

  async markLastCardAsAudioCard(): Promise<void> {
    if (this.deps.isUpdateInProgress()) {
      this.deps.showOsdNotification('Anki update already in progress');
      return;
    }

    try {
      const mpvClient = this.deps.getMpvClient();
      if (!mpvClient || !mpvClient.currentVideoPath) {
        this.deps.showOsdNotification('No video loaded');
        return;
      }

      if (!mpvClient.currentSubText) {
        this.deps.showOsdNotification('No current subtitle');
        return;
      }

      let startTime = mpvClient.currentSubStart;
      let endTime = mpvClient.currentSubEnd;

      if (startTime === undefined || endTime === undefined) {
        const currentTime = mpvClient.currentTimePos || 0;
        const fallback = this.deps.getFallbackDurationSeconds() / 2;
        startTime = currentTime - fallback;
        endTime = currentTime + fallback;
      }

      const maxMediaDuration = this.deps.getConfig().media?.maxMediaDuration ?? 30;
      if (maxMediaDuration > 0 && endTime - startTime > maxMediaDuration) {
        endTime = startTime + maxMediaDuration;
      }

      this.deps.showOsdNotification('Marking card as audio card...');
      await this.deps.withUpdateProgress('Marking audio card', async () => {
        const deck = this.deps.getDeck?.() ?? this.deps.getConfig().deck;
        const query = deck ? `"deck:${deck}" added:1` : 'added:1';
        const noteIds = (await this.deps.client.findNotes(query)) as number[];
        if (!noteIds || noteIds.length === 0) {
          this.deps.showOsdNotification('No recently added cards found');
          return;
        }

        const noteId = Math.max(...noteIds);
        const notesInfoResult = (await this.deps.client.notesInfo([
          noteId,
        ])) as CardCreationNoteInfo[];
        if (!notesInfoResult || notesInfoResult.length === 0) {
          this.deps.showOsdNotification('Card not found');
          return;
        }

        const noteInfo = notesInfoResult[0]!;
        const fields = this.deps.extractFields(noteInfo.fields);
        const expressionText = getPreferredWordValueFromExtractedFields(
          fields,
          this.deps.getConfig(),
        );

        const updatedFields: Record<string, string> = {};
        const errors: string[] = [];
        let miscInfoFilename: string | null = null;

        this.deps.setCardTypeFields(updatedFields, Object.keys(noteInfo.fields), 'audio');

        const sentenceField = this.deps.getConfig().fields?.sentence;
        if (sentenceField) {
          const processedSentence = this.deps.processSentence(mpvClient.currentSubText, fields);
          updatedFields[sentenceField] = processedSentence;
        }

        const sentenceCardConfig = this.deps.getEffectiveSentenceCardConfig();
        const audioFieldName = sentenceCardConfig.audioField;
        try {
          const audioFilename = this.generateAudioFilename();
          const audioBuffer = await this.mediaGenerateAudio(
            mpvClient.currentVideoPath,
            startTime,
            endTime,
          );

          if (audioBuffer) {
            await this.deps.client.storeMediaFile(audioFilename, audioBuffer);
            updatedFields[audioFieldName] = `[sound:${audioFilename}]`;
            miscInfoFilename = audioFilename;
          }
        } catch (error) {
          log.error('Failed to generate audio for audio card:', (error as Error).message);
          errors.push('audio');
        }

        if (shouldGenerateImage(this.deps.getConfig())) {
          try {
            const animatedLeadInSeconds = await this.deps.getAnimatedImageLeadInSeconds(noteInfo);
            const imageFilename = this.generateImageFilename();
            const imageBuffer = await this.generateImageBuffer(
              mpvClient.currentVideoPath,
              startTime,
              endTime,
              animatedLeadInSeconds,
            );

            const imageField = this.deps.getConfig().fields?.image;
            if (imageBuffer && imageField) {
              await this.deps.client.storeMediaFile(imageFilename, imageBuffer);
              updatedFields[imageField] = `<img src="${imageFilename}">`;
              miscInfoFilename = imageFilename;
            }
          } catch (error) {
            log.error('Failed to generate image for audio card:', (error as Error).message);
            errors.push('image');
          }
        }

        if (this.deps.getConfig().fields?.miscInfo) {
          const miscInfo = this.deps.formatMiscInfoPattern(miscInfoFilename || '', startTime);
          const miscInfoField = this.deps.resolveConfiguredFieldName(
            noteInfo,
            this.deps.getConfig().fields?.miscInfo,
          );
          if (miscInfo && miscInfoField) {
            updatedFields[miscInfoField] = miscInfo;
          }
        }

        await this.deps.client.updateNoteFields(noteId, updatedFields);
        await this.addConfiguredTagsToNote(noteId);
        const label = expressionText || noteId;
        log.info('Marked card as audio card:', label);
        const errorSuffix = errors.length > 0 ? `${errors.join(', ')} failed` : undefined;
        await this.deps.showNotification(noteId, label, errorSuffix);
      });
    } catch (error) {
      log.error('Error marking card as audio card:', (error as Error).message);
      this.deps.showOsdNotification(`Audio card failed: ${(error as Error).message}`);
    }
  }

  async createSentenceCard(
    sentence: string,
    startTime: number,
    endTime: number,
    secondarySubText?: string,
  ): Promise<boolean> {
    if (this.deps.isUpdateInProgress()) {
      this.deps.showOsdNotification('Anki update already in progress');
      return false;
    }

    const sentenceCardConfig = this.deps.getEffectiveSentenceCardConfig();
    const sentenceCardModel = sentenceCardConfig.model;
    if (!sentenceCardModel) {
      this.deps.showOsdNotification('sentenceCardModel not configured');
      return false;
    }

    const mpvClient = this.deps.getMpvClient();
    if (!mpvClient || !mpvClient.currentVideoPath) {
      this.deps.showOsdNotification('No video loaded');
      return false;
    }

    const maxMediaDuration = this.deps.getConfig().media?.maxMediaDuration ?? 30;
    if (maxMediaDuration > 0 && endTime - startTime > maxMediaDuration) {
      log.warn(
        `Sentence card media range ${(endTime - startTime).toFixed(1)}s exceeds cap of ${maxMediaDuration}s, clamping`,
      );
      endTime = startTime + maxMediaDuration;
    }

    try {
      return await this.deps.withUpdateProgress('Creating sentence card', async () => {
        const config = this.deps.getConfig();
        const generateAudio = shouldGenerateAudio(config);
        const generateImage = shouldGenerateImage(config);
        const mediaResolverOptions = this.getMediaResolverOptions();
        const videoPath = generateImage
          ? await resolveMediaGenerationInput(mpvClient, 'video', mediaResolverOptions)
          : null;
        const audioSourcePath = generateAudio
          ? await resolveMediaGenerationInput(mpvClient, 'audio', mediaResolverOptions)
          : null;
        const missingRequestedMediaInput =
          (generateImage && !videoPath) || (generateAudio && !audioSourcePath);
        const shouldQueuePendingYoutubeMedia =
          missingRequestedMediaInput &&
          this.deps.shouldRequireRemoteMediaCache?.() === true &&
          typeof this.deps.queuePendingYoutubeMediaUpdate === 'function';
        if (missingRequestedMediaInput && !shouldQueuePendingYoutubeMedia) {
          this.deps.showOsdNotification('No video loaded');
          return false;
        }
        const fields: Record<string, string> = {};
        const errors: string[] = [];
        let miscInfoFilename: string | null = null;

        const sentenceField = sentenceCardConfig.sentenceField;
        const audioFieldName = sentenceCardConfig.audioField || 'SentenceAudio';
        const translationField = config.fields?.translation || 'SelectionText';
        let resolvedMiscInfoField: string | null = null;
        let resolvedSentenceAudioField: string = audioFieldName;

        fields[sentenceField] = sentence;

        const ankiAiConfig = config.ai;
        const ankiAiEnabled =
          typeof ankiAiConfig === 'object' && ankiAiConfig !== null
            ? ankiAiConfig.enabled === true
            : ankiAiConfig === true;

        const backText = await resolveSentenceBackText(
          {
            sentence,
            secondarySubText,
            aiEnabled: ankiAiEnabled,
            aiConfig: this.deps.getAiConfig(),
          },
          {
            logWarning: (message: string) => log.warn(message),
          },
        );
        if (backText) {
          fields[translationField] = backText;
        }

        if (sentenceCardConfig.lapisEnabled || sentenceCardConfig.kikuEnabled) {
          fields.IsSentenceCard = 'x';
          fields[getConfiguredWordFieldName(config)] = sentence;
        }

        const pendingNoteInfo = this.createPendingNoteInfo(fields);
        const pendingNoteFields = Object.fromEntries(
          Object.entries(fields).map(([name, value]) => [name.toLowerCase(), value]),
        );
        const pendingExpressionText = getPreferredWordValueFromExtractedFields(
          pendingNoteFields,
          config,
        ).trim();
        let duplicateNoteIds: number[] = [];
        if (
          sentenceCardConfig.kikuEnabled &&
          sentenceCardConfig.kikuFieldGrouping !== 'disabled' &&
          pendingExpressionText &&
          this.deps.findDuplicateNoteIds
        ) {
          try {
            duplicateNoteIds = sortUniqueNoteIds(
              await this.deps.findDuplicateNoteIds(pendingExpressionText, pendingNoteInfo),
            );
          } catch (error) {
            log.warn('Failed to capture pre-add duplicate note ids:', (error as Error).message);
          }
        }

        const deck = config.deck || 'Default';
        let noteId: number;
        try {
          noteId = await this.deps.client.addNote(
            deck,
            sentenceCardModel,
            fields,
            this.getConfiguredAnkiTags(),
          );
          log.info('Created sentence card:', noteId);
        } catch (error) {
          log.error('Failed to create sentence card:', (error as Error).message);
          this.deps.showUpdateResult(`Sentence card failed: ${(error as Error).message}`, false);
          return false;
        }

        try {
          this.deps.trackLastAddedNoteId?.(noteId);
        } catch (error) {
          log.warn('Failed to track last added note:', (error as Error).message);
        }

        if (duplicateNoteIds.length > 0) {
          try {
            this.deps.trackLastAddedDuplicateNoteIds?.(noteId, duplicateNoteIds);
          } catch (error) {
            log.warn('Failed to track duplicate note ids:', (error as Error).message);
          }
        }

        try {
          this.deps.recordCardsMinedCallback?.(1, [noteId]);
        } catch (error) {
          log.warn('Failed to record mined card:', (error as Error).message);
        }

        try {
          const noteInfoResult = await this.deps.client.notesInfo([noteId]);
          const noteInfos = noteInfoResult as CardCreationNoteInfo[];
          if (noteInfos.length > 0) {
            const createdNoteInfo = noteInfos[0]!;
            this.deps.appendKnownWordsFromNoteInfo(createdNoteInfo);
            resolvedSentenceAudioField =
              this.deps.resolveNoteFieldName(createdNoteInfo, audioFieldName) || audioFieldName;
            resolvedMiscInfoField = this.deps.resolveConfiguredFieldName(
              createdNoteInfo,
              config.fields?.miscInfo,
            );

            const cardTypeFields: Record<string, string> = {};
            this.deps.setCardTypeFields(
              cardTypeFields,
              Object.keys(createdNoteInfo.fields),
              'sentence',
            );
            if (Object.keys(cardTypeFields).length > 0) {
              await this.deps.client.updateNoteFields(noteId, cardTypeFields);
            }
          }
        } catch (error) {
          log.error('Failed to normalize sentence card type fields:', (error as Error).message);
          errors.push('card type fields');
        }

        const label = sentence.length > 30 ? sentence.substring(0, 30) + '...' : sentence;
        if (shouldQueuePendingYoutubeMedia) {
          this.deps.queuePendingYoutubeMediaUpdate?.({
            sourceUrl:
              trimToNonEmptyString(await this.deps.getYoutubeMediaSourceUrl?.()) ??
              mpvClient.currentVideoPath,
            noteId,
            startTime,
            endTime,
            label,
            audioFieldName: resolvedSentenceAudioField,
            imageFieldName: config.fields?.image,
            miscInfoFieldName: resolvedMiscInfoField ?? undefined,
            generateAudio,
            generateImage,
          });
          await this.deps.showNotification(noteId, label, 'media queued');
          return true;
        }

        if (missingRequestedMediaInput) {
          return false;
        }

        const mediaFields: Record<string, string> = {};

        if (generateAudio) {
          try {
            const audioFilename = this.generateAudioFilename();
            const audioBuffer = audioSourcePath
              ? await this.mediaGenerateAudio(audioSourcePath, startTime, endTime)
              : null;

            if (audioBuffer) {
              await this.deps.client.storeMediaFile(audioFilename, audioBuffer);
              const audioValue = `[sound:${audioFilename}]`;
              mediaFields[resolvedSentenceAudioField] = audioValue;
              miscInfoFilename = audioFilename;
            }
          } catch (error) {
            log.error('Failed to generate sentence audio:', (error as Error).message);
            errors.push('audio');
          }
        }

        if (generateImage) {
          try {
            const imageFilename = this.generateImageFilename();
            const imageBuffer = await this.generateImageBuffer(videoPath!, startTime, endTime);

            const imageField = config.fields?.image;
            if (imageBuffer && imageField) {
              await this.deps.client.storeMediaFile(imageFilename, imageBuffer);
              mediaFields[imageField] = `<img src="${imageFilename}">`;
              miscInfoFilename = imageFilename;
            }
          } catch (error) {
            log.error('Failed to generate sentence image:', (error as Error).message);
            errors.push('image');
          }
        }

        if (config.fields?.miscInfo) {
          const miscInfo = this.deps.formatMiscInfoPattern(miscInfoFilename || '', startTime);
          if (miscInfo && resolvedMiscInfoField) {
            mediaFields[resolvedMiscInfoField] = miscInfo;
          }
        }

        if (Object.keys(mediaFields).length > 0) {
          try {
            await this.deps.client.updateNoteFields(noteId, mediaFields);
          } catch (error) {
            log.error('Failed to update sentence card media:', (error as Error).message);
            errors.push('media update');
          }
        }

        const errorSuffix = errors.length > 0 ? `${errors.join(', ')} failed` : undefined;
        await this.deps.showNotification(noteId, label, errorSuffix);
        return true;
      });
    } catch (error) {
      log.error('Error creating sentence card:', (error as Error).message);
      this.deps.showUpdateResult(`Sentence card failed: ${(error as Error).message}`, false);
      return false;
    }
  }

  private getResolvedSentenceAudioFieldName(noteInfo: CardCreationNoteInfo): string | null {
    return (
      this.deps.resolveNoteFieldName(
        noteInfo,
        this.deps.getEffectiveSentenceCardConfig().audioField || 'SentenceAudio',
      ) || this.deps.resolveConfiguredFieldName(noteInfo, this.deps.getConfig().fields?.audio)
    );
  }

  private getResolvedSentenceOnlyAudioFieldName(noteInfo: CardCreationNoteInfo): string | null {
    return this.deps.resolveNoteFieldName(
      noteInfo,
      this.deps.getEffectiveSentenceCardConfig().audioField || 'SentenceAudio',
    );
  }

  private createPendingNoteInfo(fields: Record<string, string>): CardCreationNoteInfo {
    return {
      noteId: -1,
      fields: Object.fromEntries(Object.entries(fields).map(([name, value]) => [name, { value }])),
    };
  }

  private async mediaGenerateAudio(
    videoPath: MediaInput,
    startTime: number,
    endTime: number,
  ): Promise<Buffer | null> {
    const mpvClient = this.deps.getMpvClient();
    if (!mpvClient) {
      return null;
    }

    return this.deps.mediaGenerator.generateAudio(
      videoPath,
      startTime,
      endTime,
      this.deps.getConfig().media?.audioPadding,
      resolveAudioStreamIndexForMediaGeneration(
        videoPath,
        mpvClient.currentAudioStreamIndex ?? undefined,
      ),
      this.deps.getConfig().media?.normalizeAudio !== false,
    );
  }

  private async generateImageBuffer(
    videoPath: MediaInput,
    startTime: number,
    endTime: number,
    animatedLeadInSeconds = 0,
  ): Promise<Buffer | null> {
    const mpvClient = this.deps.getMpvClient();
    if (!mpvClient) {
      return null;
    }

    const timestamp = mpvClient.currentTimePos || 0;

    if (this.deps.getConfig().media?.imageType === 'avif') {
      let imageStart = startTime;
      let imageEnd = endTime;

      if (!Number.isFinite(imageStart) || !Number.isFinite(imageEnd)) {
        const fallback = this.deps.getFallbackDurationSeconds() / 2;
        imageStart = timestamp - fallback;
        imageEnd = timestamp + fallback;
      }

      return this.deps.mediaGenerator.generateAnimatedImage(
        videoPath,
        imageStart,
        imageEnd,
        this.deps.getConfig().media?.audioPadding,
        {
          fps: this.deps.getConfig().media?.animatedFps,
          maxWidth: this.deps.getConfig().media?.animatedMaxWidth,
          maxHeight: this.deps.getConfig().media?.animatedMaxHeight,
          crf: this.deps.getConfig().media?.animatedCrf,
          leadingStillDuration: animatedLeadInSeconds,
        },
      );
    }

    return this.deps.mediaGenerator.generateScreenshot(videoPath, timestamp, {
      format: this.deps.getConfig().media?.imageFormat as 'jpg' | 'png' | 'webp',
      quality: this.deps.getConfig().media?.imageQuality,
      maxWidth: this.deps.getConfig().media?.imageMaxWidth,
      maxHeight: this.deps.getConfig().media?.imageMaxHeight,
    });
  }

  private generateAudioFilename(): string {
    const timestamp = Date.now();
    return `audio_${timestamp}.mp3`;
  }

  private generateImageFilename(): string {
    const timestamp = Date.now();
    const ext =
      this.deps.getConfig().media?.imageType === 'avif'
        ? 'avif'
        : this.deps.getConfig().media?.imageFormat;
    return `image_${timestamp}.${ext}`;
  }
}

function sortUniqueNoteIds(noteIds: number[]): number[] {
  return [...new Set(noteIds)].sort((left, right) => left - right);
}
