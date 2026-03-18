import { DEFAULT_ANKI_CONNECT_CONFIG } from '../config';
import {
  getConfiguredWordFieldName,
  getPreferredWordValueFromExtractedFields,
} from '../anki-field-config';
import { AiConfig, AnkiConnectConfig } from '../types';
import { createLogger } from '../logger';
import { SubtitleTimingTracker } from '../subtitle-timing-tracker';
import { MpvClient } from '../types';
import { resolveSentenceBackText } from './ai';

const log = createLogger('anki').child('integration.card-creation');

export interface CardCreationNoteInfo {
  noteId: number;
  fields: Record<string, { value: string }>;
}

type CardKind = 'sentence' | 'audio';

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
}

interface CardCreationMediaGenerator {
  generateAudio(
    path: string,
    startTime: number,
    endTime: number,
    audioPadding?: number,
    audioStreamIndex?: number,
  ): Promise<Buffer | null>;
  generateScreenshot(
    path: string,
    timestamp: number,
    options: {
      format: 'jpg' | 'png' | 'webp';
      quality?: number;
      maxWidth?: number;
      maxHeight?: number;
    },
  ): Promise<Buffer | null>;
  generateAnimatedImage(
    path: string,
    startTime: number,
    endTime: number,
    audioPadding?: number,
    options?: {
      fps?: number;
      maxWidth?: number;
      maxHeight?: number;
      crf?: number;
    },
  ): Promise<Buffer | null>;
}

interface CardCreationDeps {
  getConfig: () => AnkiConnectConfig;
  getAiConfig: () => AiConfig;
  getTimingTracker: () => SubtitleTimingTracker;
  getMpvClient: () => MpvClient;
  getDeck?: () => string | undefined;
  client: CardCreationClient;
  mediaGenerator: CardCreationMediaGenerator;
  showOsdNotification: (text: string) => void;
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
}

export class CardCreationService {
  constructor(private readonly deps: CardCreationDeps) {}

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
        const sentenceAudioField = this.getResolvedSentenceAudioFieldName(noteInfo);
        const sentenceField = this.deps.getEffectiveSentenceCardConfig().sentenceField;

        const sentence = blocks.join(' ');
        const updatedFields: Record<string, string> = {};
        let updatePerformed = false;
        const errors: string[] = [];
        let miscInfoFilename: string | null = null;

        if (sentenceField) {
          const processedSentence = this.deps.processSentence(sentence, fields);
          updatedFields[sentenceField] = processedSentence;
          updatePerformed = true;
        }

        log.info(
          `Clipboard update: timing range ${rangeStart.toFixed(2)}s - ${rangeEnd.toFixed(2)}s`,
        );

        if (this.deps.getConfig().media?.generateAudio) {
          try {
            const audioFilename = this.generateAudioFilename();
            const audioBuffer = await this.mediaGenerateAudio(
              mpvClient.currentVideoPath,
              rangeStart,
              rangeEnd,
            );

            if (audioBuffer) {
              await this.deps.client.storeMediaFile(audioFilename, audioBuffer);
              if (sentenceAudioField) {
                const existingAudio = noteInfo.fields[sentenceAudioField]?.value || '';
                updatedFields[sentenceAudioField] = this.deps.mergeFieldValue(
                  existingAudio,
                  `[sound:${audioFilename}]`,
                  this.deps.getConfig().behavior?.overwriteAudio !== false,
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

        if (this.deps.getConfig().media?.generateImage) {
          try {
            const imageFilename = this.generateImageFilename();
            const imageBuffer = await this.generateImageBuffer(
              mpvClient.currentVideoPath,
              rangeStart,
              rangeEnd,
            );

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

        if (this.deps.getConfig().media?.generateImage) {
          try {
            const imageFilename = this.generateImageFilename();
            const imageBuffer = await this.generateImageBuffer(
              mpvClient.currentVideoPath,
              startTime,
              endTime,
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

    this.deps.showOsdNotification('Creating sentence card...');
    try {
      return await this.deps.withUpdateProgress('Creating sentence card', async () => {
        const videoPath = mpvClient.currentVideoPath;
        const fields: Record<string, string> = {};
        const errors: string[] = [];
        let miscInfoFilename: string | null = null;

        const sentenceField = sentenceCardConfig.sentenceField;
        const audioFieldName = sentenceCardConfig.audioField || 'SentenceAudio';
        const translationField = this.deps.getConfig().fields?.translation || 'SelectionText';
        let resolvedMiscInfoField: string | null = null;
        let resolvedSentenceAudioField: string = audioFieldName;
        let resolvedExpressionAudioField: string | null = null;

        fields[sentenceField] = sentence;

        const ankiAiConfig = this.deps.getConfig().ai;
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
          fields[getConfiguredWordFieldName(this.deps.getConfig())] = sentence;
        }

        const deck = this.deps.getConfig().deck || 'Default';
        let noteId: number;
        try {
          noteId = await this.deps.client.addNote(
            deck,
            sentenceCardModel,
            fields,
            this.getConfiguredAnkiTags(),
          );
          log.info('Created sentence card:', noteId);
          this.deps.trackLastAddedNoteId?.(noteId);
        } catch (error) {
          log.error('Failed to create sentence card:', (error as Error).message);
          this.deps.showOsdNotification(`Sentence card failed: ${(error as Error).message}`);
          return false;
        }

        try {
          const noteInfoResult = await this.deps.client.notesInfo([noteId]);
          const noteInfos = noteInfoResult as CardCreationNoteInfo[];
          if (noteInfos.length > 0) {
            const createdNoteInfo = noteInfos[0]!;
            this.deps.appendKnownWordsFromNoteInfo(createdNoteInfo);
            resolvedSentenceAudioField =
              this.deps.resolveNoteFieldName(createdNoteInfo, audioFieldName) || audioFieldName;
            resolvedExpressionAudioField = this.deps.resolveConfiguredFieldName(
              createdNoteInfo,
              this.deps.getConfig().fields?.audio || 'ExpressionAudio',
            );
            resolvedMiscInfoField = this.deps.resolveConfiguredFieldName(
              createdNoteInfo,
              this.deps.getConfig().fields?.miscInfo,
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

        const mediaFields: Record<string, string> = {};

        try {
          const audioFilename = this.generateAudioFilename();
          const audioBuffer = await this.mediaGenerateAudio(videoPath, startTime, endTime);

          if (audioBuffer) {
            await this.deps.client.storeMediaFile(audioFilename, audioBuffer);
            const audioValue = `[sound:${audioFilename}]`;
            mediaFields[resolvedSentenceAudioField] = audioValue;
            if (
              resolvedExpressionAudioField &&
              resolvedExpressionAudioField !== resolvedSentenceAudioField
            ) {
              mediaFields[resolvedExpressionAudioField] = audioValue;
            }
            miscInfoFilename = audioFilename;
          }
        } catch (error) {
          log.error('Failed to generate sentence audio:', (error as Error).message);
          errors.push('audio');
        }

        try {
          const imageFilename = this.generateImageFilename();
          const imageBuffer = await this.generateImageBuffer(videoPath, startTime, endTime);

          const imageField = this.deps.getConfig().fields?.image;
          if (imageBuffer && imageField) {
            await this.deps.client.storeMediaFile(imageFilename, imageBuffer);
            mediaFields[imageField] = `<img src="${imageFilename}">`;
            miscInfoFilename = imageFilename;
          }
        } catch (error) {
          log.error('Failed to generate sentence image:', (error as Error).message);
          errors.push('image');
        }

        if (this.deps.getConfig().fields?.miscInfo) {
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

        const label = sentence.length > 30 ? sentence.substring(0, 30) + '...' : sentence;
        const errorSuffix = errors.length > 0 ? `${errors.join(', ')} failed` : undefined;
        await this.deps.showNotification(noteId, label, errorSuffix);
        return true;
      });
    } catch (error) {
      log.error('Error creating sentence card:', (error as Error).message);
      this.deps.showOsdNotification(`Sentence card failed: ${(error as Error).message}`);
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

  private async mediaGenerateAudio(
    videoPath: string,
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
      mpvClient.currentAudioStreamIndex ?? undefined,
    );
  }

  private async generateImageBuffer(
    videoPath: string,
    startTime: number,
    endTime: number,
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
