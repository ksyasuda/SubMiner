/*
 * SubMiner - Subtitle mining overlay for mpv
 * Copyright (C) 2024 sudacode
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

import { AnkiConnectClient } from './anki-connect';
import { SubtitleTimingTracker } from './subtitle-timing-tracker';
import { MediaGenerator } from './media-generator';
import path from 'path';
import {
  AnkiConnectConfig,
  KikuDuplicateCardInfo,
  KikuFieldGroupingChoice,
  KikuMergePreviewResponse,
  MpvClient,
  NotificationOptions,
  NPlusOneMatchMode,
} from './types';
import { DEFAULT_ANKI_CONNECT_CONFIG } from './config';
import { createLogger } from './logger';
import {
  createUiFeedbackState,
  beginUpdateProgress,
  endUpdateProgress,
  showProgressTick,
  showStatusNotification,
  withUpdateProgress,
  UiFeedbackState,
} from './anki-integration/ui-feedback';
import { KnownWordCacheManager } from './anki-integration/known-word-cache';
import { PollingRunner } from './anki-integration/polling';
import { findDuplicateNote as findDuplicateNoteForAnkiIntegration } from './anki-integration/duplicate';
import { CardCreationService } from './anki-integration/card-creation';
import { FieldGroupingService } from './anki-integration/field-grouping';

const log = createLogger('anki').child('integration');

interface NoteInfo {
  noteId: number;
  fields: Record<string, { value: string }>;
}

type CardKind = 'sentence' | 'audio';

export class AnkiIntegration {
  private client: AnkiConnectClient;
  private mediaGenerator: MediaGenerator;
  private timingTracker: SubtitleTimingTracker;
  private config: AnkiConnectConfig;
  private pollingRunner!: PollingRunner;
  private previousNoteIds = new Set<number>();
  private mpvClient: MpvClient;
  private osdCallback: ((text: string) => void) | null = null;
  private notificationCallback: ((title: string, options: NotificationOptions) => void) | null =
    null;
  private updateInProgress = false;
  private uiFeedbackState: UiFeedbackState = createUiFeedbackState();
  private parseWarningKeys = new Set<string>();
  private readonly strictGroupingFieldDefaults = new Set<string>([
    'picture',
    'sentence',
    'sentenceaudio',
    'sentencefurigana',
    'miscinfo',
  ]);
  private fieldGroupingCallback:
    | ((data: {
        original: KikuDuplicateCardInfo;
        duplicate: KikuDuplicateCardInfo;
      }) => Promise<KikuFieldGroupingChoice>)
    | null = null;
  private knownWordCache: KnownWordCacheManager;
  private cardCreationService: CardCreationService;
  private fieldGroupingService: FieldGroupingService;

  constructor(
    config: AnkiConnectConfig,
    timingTracker: SubtitleTimingTracker,
    mpvClient: MpvClient,
    osdCallback?: (text: string) => void,
    notificationCallback?: (title: string, options: NotificationOptions) => void,
    fieldGroupingCallback?: (data: {
      original: KikuDuplicateCardInfo;
      duplicate: KikuDuplicateCardInfo;
    }) => Promise<KikuFieldGroupingChoice>,
    knownWordCacheStatePath?: string,
  ) {
    this.config = {
      ...DEFAULT_ANKI_CONNECT_CONFIG,
      ...config,
      fields: {
        ...DEFAULT_ANKI_CONNECT_CONFIG.fields,
        ...(config.fields ?? {}),
      },
      ai: {
        ...DEFAULT_ANKI_CONNECT_CONFIG.ai,
        ...(config.openRouter ?? {}),
        ...(config.ai ?? {}),
      },
      media: {
        ...DEFAULT_ANKI_CONNECT_CONFIG.media,
        ...(config.media ?? {}),
      },
      behavior: {
        ...DEFAULT_ANKI_CONNECT_CONFIG.behavior,
        ...(config.behavior ?? {}),
      },
      metadata: {
        ...DEFAULT_ANKI_CONNECT_CONFIG.metadata,
        ...(config.metadata ?? {}),
      },
      isLapis: {
        ...DEFAULT_ANKI_CONNECT_CONFIG.isLapis,
        ...(config.isLapis ?? {}),
      },
      isKiku: {
        ...DEFAULT_ANKI_CONNECT_CONFIG.isKiku,
        ...(config.isKiku ?? {}),
      },
    } as AnkiConnectConfig;

    this.client = new AnkiConnectClient(this.config.url!);
    this.mediaGenerator = new MediaGenerator();
    this.timingTracker = timingTracker;
    this.mpvClient = mpvClient;
    this.osdCallback = osdCallback || null;
    this.notificationCallback = notificationCallback || null;
    this.fieldGroupingCallback = fieldGroupingCallback || null;
    this.knownWordCache = new KnownWordCacheManager({
      client: {
        findNotes: async (query, options) =>
          (await this.client.findNotes(query, options)) as unknown,
        notesInfo: async (noteIds) => (await this.client.notesInfo(noteIds)) as unknown,
      },
      getConfig: () => this.config,
      knownWordCacheStatePath,
      showStatusNotification: (message: string) => this.showStatusNotification(message),
    });
    this.pollingRunner = new PollingRunner({
      getDeck: () => this.config.deck,
      getPollingRate: () => this.config.pollingRate || DEFAULT_ANKI_CONNECT_CONFIG.pollingRate,
      findNotes: async (query, options) =>
        (await this.client.findNotes(query, options)) as number[],
      shouldAutoUpdateNewCards: () => this.config.behavior?.autoUpdateNewCards !== false,
      processNewCard: (noteId) => this.processNewCard(noteId),
      isUpdateInProgress: () => this.updateInProgress,
      setUpdateInProgress: (value) => {
        this.updateInProgress = value;
      },
      getTrackedNoteIds: () => this.previousNoteIds,
      setTrackedNoteIds: (noteIds) => {
        this.previousNoteIds = noteIds;
      },
      showStatusNotification: (message: string) => this.showStatusNotification(message),
      logDebug: (...args) => log.debug(args[0] as string, ...args.slice(1)),
      logInfo: (...args) => log.info(args[0] as string, ...args.slice(1)),
      logWarn: (...args) => log.warn(args[0] as string, ...args.slice(1)),
    });
    this.cardCreationService = new CardCreationService({
      getConfig: () => this.config,
      getTimingTracker: () => this.timingTracker,
      getMpvClient: () => this.mpvClient,
      getDeck: () => this.config.deck,
      client: {
        addNote: (deck, modelName, fields, tags) =>
          this.client.addNote(deck, modelName, fields, tags),
        addTags: (noteIds, tags) => this.client.addTags(noteIds, tags),
        notesInfo: async (noteIds) => (await this.client.notesInfo(noteIds)) as unknown,
        updateNoteFields: (noteId, fields) =>
          this.client.updateNoteFields(noteId, fields) as Promise<void>,
        storeMediaFile: (filename, data) => this.client.storeMediaFile(filename, data),
        findNotes: async (query, options) =>
          (await this.client.findNotes(query, options)) as number[],
      },
      mediaGenerator: {
        generateAudio: (videoPath, startTime, endTime, audioPadding, audioStreamIndex) =>
          this.mediaGenerator.generateAudio(
            videoPath,
            startTime,
            endTime,
            audioPadding,
            audioStreamIndex,
          ),
        generateScreenshot: (videoPath, timestamp, options) =>
          this.mediaGenerator.generateScreenshot(videoPath, timestamp, options),
        generateAnimatedImage: (videoPath, startTime, endTime, audioPadding, options) =>
          this.mediaGenerator.generateAnimatedImage(
            videoPath,
            startTime,
            endTime,
            audioPadding,
            options,
          ),
      },
      showOsdNotification: (text: string) => this.showOsdNotification(text),
      showStatusNotification: (message: string) => this.showStatusNotification(message),
      showNotification: (noteId, label, errorSuffix) =>
        this.showNotification(noteId, label, errorSuffix),
      beginUpdateProgress: (initialMessage: string) => this.beginUpdateProgress(initialMessage),
      endUpdateProgress: () => this.endUpdateProgress(),
      withUpdateProgress: <T>(initialMessage: string, action: () => Promise<T>) =>
        this.withUpdateProgress(initialMessage, action),
      resolveConfiguredFieldName: (noteInfo, ...preferredNames) =>
        this.resolveConfiguredFieldName(noteInfo, ...preferredNames),
      resolveNoteFieldName: (noteInfo, preferredName) =>
        this.resolveNoteFieldName(noteInfo, preferredName),
      extractFields: (fields) => this.extractFields(fields),
      processSentence: (mpvSentence, noteFields) => this.processSentence(mpvSentence, noteFields),
      setCardTypeFields: (updatedFields, availableFieldNames, cardKind) =>
        this.setCardTypeFields(updatedFields, availableFieldNames, cardKind),
      mergeFieldValue: (existing, newValue, overwrite) =>
        this.mergeFieldValue(existing, newValue, overwrite),
      formatMiscInfoPattern: (fallbackFilename, startTimeSeconds) =>
        this.formatMiscInfoPattern(fallbackFilename, startTimeSeconds),
      getEffectiveSentenceCardConfig: () => this.getEffectiveSentenceCardConfig(),
      getFallbackDurationSeconds: () => this.getFallbackDurationSeconds(),
      appendKnownWordsFromNoteInfo: (noteInfo) => this.appendKnownWordsFromNoteInfo(noteInfo),
      isUpdateInProgress: () => this.updateInProgress,
      setUpdateInProgress: (value) => {
        this.updateInProgress = value;
      },
      trackLastAddedNoteId: (noteId) => {
        this.previousNoteIds.add(noteId);
      },
    });
    this.fieldGroupingService = new FieldGroupingService({
      getEffectiveSentenceCardConfig: () => this.getEffectiveSentenceCardConfig(),
      isUpdateInProgress: () => this.updateInProgress,
      getDeck: () => this.config.deck,
      withUpdateProgress: <T>(initialMessage: string, action: () => Promise<T>) =>
        this.withUpdateProgress(initialMessage, action),
      showOsdNotification: (text: string) => this.showOsdNotification(text),
      findNotes: async (query, options) =>
        (await this.client.findNotes(query, options)) as number[],
      notesInfo: async (noteIds) => (await this.client.notesInfo(noteIds)) as unknown as NoteInfo[],
      extractFields: (fields) => this.extractFields(fields),
      findDuplicateNote: (expression, noteId, noteInfo) =>
        this.findDuplicateNote(expression, noteId, noteInfo),
      hasAllConfiguredFields: (noteInfo, configuredFieldNames) =>
        this.hasAllConfiguredFields(noteInfo, configuredFieldNames),
      processNewCard: (noteId, options) => this.processNewCard(noteId, options),
      getSentenceCardImageFieldName: () => this.config.fields?.image,
      resolveFieldName: (availableFieldNames, preferredName) =>
        this.resolveFieldName(availableFieldNames, preferredName),
      computeFieldGroupingMergedFields: (
        keepNoteId,
        deleteNoteId,
        keepNoteInfo,
        deleteNoteInfo,
        includeGeneratedMedia,
      ) =>
        this.computeFieldGroupingMergedFields(
          keepNoteId,
          deleteNoteId,
          keepNoteInfo,
          deleteNoteInfo,
          includeGeneratedMedia,
        ),
      getNoteFieldMap: (noteInfo) => this.getNoteFieldMap(noteInfo),
      handleFieldGroupingAuto: (originalNoteId, newNoteId, newNoteInfo, expression) =>
        this.handleFieldGroupingAuto(originalNoteId, newNoteId, newNoteInfo, expression),
      handleFieldGroupingManual: (originalNoteId, newNoteId, newNoteInfo, expression) =>
        this.handleFieldGroupingManual(originalNoteId, newNoteId, newNoteInfo, expression),
    });
  }

  isKnownWord(text: string): boolean {
    return this.knownWordCache.isKnownWord(text);
  }

  getKnownWordMatchMode(): NPlusOneMatchMode {
    return this.config.nPlusOne?.matchMode ?? DEFAULT_ANKI_CONNECT_CONFIG.nPlusOne.matchMode;
  }

  private isKnownWordCacheEnabled(): boolean {
    return this.config.nPlusOne?.highlightEnabled === true;
  }

  private startKnownWordCacheLifecycle(): void {
    this.knownWordCache.startLifecycle();
  }

  private stopKnownWordCacheLifecycle(): void {
    this.knownWordCache.stopLifecycle();
  }

  private getConfiguredAnkiTags(): string[] {
    if (!Array.isArray(this.config.tags)) {
      return [];
    }
    return [...new Set(this.config.tags.map((tag) => tag.trim()).filter((tag) => tag.length > 0))];
  }

  private async addConfiguredTagsToNote(noteId: number): Promise<void> {
    const tags = this.getConfiguredAnkiTags();
    if (tags.length === 0) {
      return;
    }
    try {
      await this.client.addTags([noteId], tags);
    } catch (error) {
      log.warn('Failed to add tags to card:', (error as Error).message);
    }
  }

  async refreshKnownWordCache(): Promise<void> {
    return this.knownWordCache.refresh(true);
  }

  private appendKnownWordsFromNoteInfo(noteInfo: NoteInfo): void {
    if (!this.isKnownWordCacheEnabled()) {
      return;
    }

    this.knownWordCache.appendFromNoteInfo({
      noteId: noteInfo.noteId,
      fields: noteInfo.fields,
    });
  }

  private getLapisConfig(): {
    enabled: boolean;
    sentenceCardModel?: string;
    sentenceCardSentenceField?: string;
    sentenceCardAudioField?: string;
  } {
    const lapis = this.config.isLapis;
    return {
      enabled: lapis?.enabled === true,
      sentenceCardModel: lapis?.sentenceCardModel,
      sentenceCardSentenceField: lapis?.sentenceCardSentenceField,
      sentenceCardAudioField: lapis?.sentenceCardAudioField,
    };
  }

  private getKikuConfig(): {
    enabled: boolean;
    fieldGrouping?: 'auto' | 'manual' | 'disabled';
    deleteDuplicateInAuto?: boolean;
  } {
    const kiku = this.config.isKiku;
    return {
      enabled: kiku?.enabled === true,
      fieldGrouping: kiku?.fieldGrouping,
      deleteDuplicateInAuto: kiku?.deleteDuplicateInAuto,
    };
  }

  private getEffectiveSentenceCardConfig(): {
    model?: string;
    sentenceField: string;
    audioField: string;
    lapisEnabled: boolean;
    kikuEnabled: boolean;
    kikuFieldGrouping: 'auto' | 'manual' | 'disabled';
    kikuDeleteDuplicateInAuto: boolean;
  } {
    const lapis = this.getLapisConfig();
    const kiku = this.getKikuConfig();

    return {
      model: lapis.sentenceCardModel,
      sentenceField: lapis.sentenceCardSentenceField || 'Sentence',
      audioField: lapis.sentenceCardAudioField || 'SentenceAudio',
      lapisEnabled: lapis.enabled,
      kikuEnabled: kiku.enabled,
      kikuFieldGrouping: (kiku.fieldGrouping || 'disabled') as 'auto' | 'manual' | 'disabled',
      kikuDeleteDuplicateInAuto: kiku.deleteDuplicateInAuto !== false,
    };
  }

  start(): void {
    if (this.pollingRunner.isRunning) {
      this.stop();
    }

    log.info('Starting AnkiConnect integration with polling rate:', this.config.pollingRate);
    this.startKnownWordCacheLifecycle();
    this.pollingRunner.start();
  }

  stop(): void {
    this.pollingRunner.stop();
    this.stopKnownWordCacheLifecycle();
    log.info('Stopped AnkiConnect integration');
  }

  private poll(): void {
    void this.pollingRunner.poll();
  }

  private async processNewCard(
    noteId: number,
    options?: { skipKikuFieldGrouping?: boolean },
  ): Promise<void> {
    this.beginUpdateProgress('Updating card');
    try {
      const notesInfoResult = await this.client.notesInfo([noteId]);
      const notesInfo = notesInfoResult as unknown as NoteInfo[];
      if (!notesInfo || notesInfo.length === 0) {
        log.warn('Card not found:', noteId);
        return;
      }

      const noteInfo = notesInfo[0];
      this.appendKnownWordsFromNoteInfo(noteInfo);
      const fields = this.extractFields(noteInfo.fields);

      const expressionText = fields.expression || fields.word || '';
      if (!expressionText) {
        log.warn('No expression/word field found in card:', noteId);
        return;
      }

      const sentenceCardConfig = this.getEffectiveSentenceCardConfig();
      if (
        !options?.skipKikuFieldGrouping &&
        sentenceCardConfig.kikuEnabled &&
        sentenceCardConfig.kikuFieldGrouping !== 'disabled'
      ) {
        const duplicateNoteId = await this.findDuplicateNote(expressionText, noteId, noteInfo);
        if (duplicateNoteId !== null) {
          if (sentenceCardConfig.kikuFieldGrouping === 'auto') {
            await this.handleFieldGroupingAuto(duplicateNoteId, noteId, noteInfo, expressionText);
            return;
          } else if (sentenceCardConfig.kikuFieldGrouping === 'manual') {
            const handled = await this.handleFieldGroupingManual(
              duplicateNoteId,
              noteId,
              noteInfo,
              expressionText,
            );
            if (handled) return;
          }
        }
      }

      const updatedFields: Record<string, string> = {};
      let updatePerformed = false;
      let miscInfoFilename: string | null = null;
      const sentenceField = sentenceCardConfig.sentenceField;

      if (sentenceField && this.mpvClient.currentSubText) {
        const processedSentence = this.processSentence(this.mpvClient.currentSubText, fields);
        updatedFields[sentenceField] = processedSentence;
        updatePerformed = true;
      }

      if (this.config.media?.generateAudio && this.mpvClient) {
        try {
          const audioFilename = this.generateAudioFilename();
          const audioBuffer = await this.generateAudio();

          if (audioBuffer) {
            await this.client.storeMediaFile(audioFilename, audioBuffer);
            const sentenceAudioField = this.getResolvedSentenceAudioFieldName(noteInfo);
            if (sentenceAudioField) {
              const existingAudio = noteInfo.fields[sentenceAudioField]?.value || '';
              updatedFields[sentenceAudioField] = this.mergeFieldValue(
                existingAudio,
                `[sound:${audioFilename}]`,
                this.config.behavior?.overwriteAudio !== false,
              );
            }
            miscInfoFilename = audioFilename;
            updatePerformed = true;
          }
        } catch (error) {
          log.error('Failed to generate audio:', (error as Error).message);
          this.showOsdNotification(`Audio generation failed: ${(error as Error).message}`);
        }
      }

      let imageBuffer: Buffer | null = null;
      if (this.config.media?.generateImage && this.mpvClient) {
        try {
          const imageFilename = this.generateImageFilename();
          imageBuffer = await this.generateImage();

          if (imageBuffer) {
            await this.client.storeMediaFile(imageFilename, imageBuffer);
            const imageFieldName = this.resolveConfiguredFieldName(
              noteInfo,
              this.config.fields?.image,
              DEFAULT_ANKI_CONNECT_CONFIG.fields.image,
            );
            if (!imageFieldName) {
              log.warn('Image field not found on note, skipping image update');
            } else {
              const existingImage = noteInfo.fields[imageFieldName]?.value || '';
              updatedFields[imageFieldName] = this.mergeFieldValue(
                existingImage,
                `<img src="${imageFilename}">`,
                this.config.behavior?.overwriteImage !== false,
              );
              miscInfoFilename = imageFilename;
              updatePerformed = true;
            }
          }
        } catch (error) {
          log.error('Failed to generate image:', (error as Error).message);
          this.showOsdNotification(`Image generation failed: ${(error as Error).message}`);
        }
      }

      if (this.config.fields?.miscInfo) {
        const miscInfo = this.formatMiscInfoPattern(
          miscInfoFilename || '',
          this.mpvClient.currentSubStart,
        );
        const miscInfoField = this.resolveConfiguredFieldName(
          noteInfo,
          this.config.fields?.miscInfo,
        );
        if (miscInfo && miscInfoField) {
          updatedFields[miscInfoField] = miscInfo;
          updatePerformed = true;
        }
      }

      if (updatePerformed) {
        await this.client.updateNoteFields(noteId, updatedFields);
        await this.addConfiguredTagsToNote(noteId);
        log.info('Updated card fields for:', expressionText);
        await this.showNotification(noteId, expressionText);
      }
    } catch (error) {
      if ((error as Error).message.includes('note was not found')) {
        log.warn('Card was deleted before update:', noteId);
      } else {
        log.error('Error processing new card:', (error as Error).message);
      }
    } finally {
      this.endUpdateProgress();
    }
  }

  private extractFields(fields: Record<string, { value: string }>): Record<string, string> {
    const result: Record<string, string> = {};
    for (const [key, value] of Object.entries(fields)) {
      result[key.toLowerCase()] = value.value || '';
    }
    return result;
  }

  private processSentence(mpvSentence: string, noteFields: Record<string, string>): string {
    if (this.config.behavior?.highlightWord === false) {
      return mpvSentence;
    }

    const sentenceFieldName = this.config.fields?.sentence?.toLowerCase() || 'sentence';
    const existingSentence = noteFields[sentenceFieldName] || '';

    const highlightMatch = existingSentence.match(/<b>(.*?)<\/b>/);
    if (!highlightMatch || !highlightMatch[1]) {
      return mpvSentence;
    }

    const highlightedText = highlightMatch[1];
    const index = mpvSentence.indexOf(highlightedText);

    if (index === -1) {
      return mpvSentence;
    }

    const prefix = mpvSentence.substring(0, index);
    const suffix = mpvSentence.substring(index + highlightedText.length);
    return `${prefix}<b>${highlightedText}</b>${suffix}`;
  }

  private async generateAudio(): Promise<Buffer | null> {
    const mpvClient = this.mpvClient;
    if (!mpvClient || !mpvClient.currentVideoPath) {
      return null;
    }

    const videoPath = mpvClient.currentVideoPath;
    let startTime = mpvClient.currentSubStart;
    let endTime = mpvClient.currentSubEnd;

    if (startTime === undefined || endTime === undefined) {
      const currentTime = mpvClient.currentTimePos || 0;
      const fallback = this.getFallbackDurationSeconds() / 2;
      startTime = currentTime - fallback;
      endTime = currentTime + fallback;
    }

    return this.mediaGenerator.generateAudio(
      videoPath,
      startTime,
      endTime,
      this.config.media?.audioPadding,
      this.mpvClient.currentAudioStreamIndex,
    );
  }

  private async generateImage(): Promise<Buffer | null> {
    if (!this.mpvClient || !this.mpvClient.currentVideoPath) {
      return null;
    }

    const videoPath = this.mpvClient.currentVideoPath;
    const timestamp = this.mpvClient.currentTimePos || 0;

    if (this.config.media?.imageType === 'avif') {
      let startTime = this.mpvClient.currentSubStart;
      let endTime = this.mpvClient.currentSubEnd;

      if (startTime === undefined || endTime === undefined) {
        const fallback = this.getFallbackDurationSeconds() / 2;
        startTime = timestamp - fallback;
        endTime = timestamp + fallback;
      }

      return this.mediaGenerator.generateAnimatedImage(
        videoPath,
        startTime,
        endTime,
        this.config.media?.audioPadding,
        {
          fps: this.config.media?.animatedFps,
          maxWidth: this.config.media?.animatedMaxWidth,
          maxHeight: this.config.media?.animatedMaxHeight,
          crf: this.config.media?.animatedCrf,
        },
      );
    } else {
      return this.mediaGenerator.generateScreenshot(videoPath, timestamp, {
        format: this.config.media?.imageFormat as 'jpg' | 'png' | 'webp',
        quality: this.config.media?.imageQuality,
        maxWidth: this.config.media?.imageMaxWidth,
        maxHeight: this.config.media?.imageMaxHeight,
      });
    }
  }

  private formatMiscInfoPattern(fallbackFilename: string, startTimeSeconds?: number): string {
    if (!this.config.metadata?.pattern) {
      return '';
    }

    const currentVideoPath = this.mpvClient.currentVideoPath || '';
    const videoFilename = currentVideoPath ? path.basename(currentVideoPath) : '';
    const filenameWithExt = videoFilename || fallbackFilename;
    const filenameWithoutExt = filenameWithExt.replace(/\.[^.]+$/, '');

    const currentTimePos =
      typeof startTimeSeconds === 'number' && Number.isFinite(startTimeSeconds)
        ? startTimeSeconds
        : this.mpvClient.currentTimePos;
    let totalMilliseconds = 0;
    if (Number.isFinite(currentTimePos) && currentTimePos >= 0) {
      totalMilliseconds = Math.floor(currentTimePos * 1000);
    } else {
      const now = new Date();
      totalMilliseconds =
        now.getHours() * 3600000 +
        now.getMinutes() * 60000 +
        now.getSeconds() * 1000 +
        now.getMilliseconds();
    }

    const totalSeconds = Math.floor(totalMilliseconds / 1000);
    const hours = String(Math.floor(totalSeconds / 3600)).padStart(2, '0');
    const minutes = String(Math.floor((totalSeconds % 3600) / 60)).padStart(2, '0');
    const seconds = String(totalSeconds % 60).padStart(2, '0');
    const milliseconds = String(totalMilliseconds % 1000).padStart(3, '0');

    let result = this.config.metadata?.pattern
      .replace(/%f/g, filenameWithoutExt)
      .replace(/%F/g, filenameWithExt)
      .replace(/%t/g, `${hours}:${minutes}:${seconds}`)
      .replace(/%T/g, `${hours}:${minutes}:${seconds}:${milliseconds}`)
      .replace(/<br>/g, '\n');

    return result;
  }

  private getFallbackDurationSeconds(): number {
    const configured = this.config.media?.fallbackDuration;
    if (typeof configured === 'number' && Number.isFinite(configured) && configured > 0) {
      return configured;
    }
    return DEFAULT_ANKI_CONNECT_CONFIG.media.fallbackDuration;
  }

  private generateAudioFilename(): string {
    const timestamp = Date.now();
    return `audio_${timestamp}.mp3`;
  }

  private generateImageFilename(): string {
    const timestamp = Date.now();
    const ext = this.config.media?.imageType === 'avif' ? 'avif' : this.config.media?.imageFormat;
    return `image_${timestamp}.${ext}`;
  }

  private showStatusNotification(message: string): void {
    showStatusNotification(message, {
      getNotificationType: () => this.config.behavior?.notificationType,
      showOsd: (text: string) => {
        this.showOsdNotification(text);
      },
      showSystemNotification: (title: string, options: NotificationOptions) => {
        if (this.notificationCallback) {
          this.notificationCallback(title, options);
        }
      },
    });
  }

  private beginUpdateProgress(initialMessage: string): void {
    beginUpdateProgress(this.uiFeedbackState, initialMessage, (text: string) => {
      this.showOsdNotification(text);
    });
  }

  private endUpdateProgress(): void {
    endUpdateProgress(this.uiFeedbackState, (timer) => {
      clearInterval(timer);
    });
  }

  private showProgressTick(): void {
    showProgressTick(this.uiFeedbackState, (text: string) => {
      this.showOsdNotification(text);
    });
  }

  private async withUpdateProgress<T>(
    initialMessage: string,
    action: () => Promise<T>,
  ): Promise<T> {
    return withUpdateProgress(
      this.uiFeedbackState,
      {
        setUpdateInProgress: (value: boolean) => {
          this.updateInProgress = value;
        },
        showOsdNotification: (text: string) => {
          this.showOsdNotification(text);
        },
      },
      initialMessage,
      action,
    );
  }

  private showOsdNotification(text: string): void {
    if (this.osdCallback) {
      this.osdCallback(text);
    } else if (this.mpvClient && this.mpvClient.send) {
      this.mpvClient.send({
        command: ['show-text', text, '3000'],
      });
    }
  }

  private resolveFieldName(availableFieldNames: string[], preferredName: string): string | null {
    const exact = availableFieldNames.find((name) => name === preferredName);
    if (exact) return exact;

    const lower = preferredName.toLowerCase();
    const ci = availableFieldNames.find((name) => name.toLowerCase() === lower);
    return ci || null;
  }

  private resolveNoteFieldName(noteInfo: NoteInfo, preferredName?: string): string | null {
    if (!preferredName) return null;
    return this.resolveFieldName(Object.keys(noteInfo.fields), preferredName);
  }

  private resolveConfiguredFieldName(
    noteInfo: NoteInfo,
    ...preferredNames: (string | undefined)[]
  ): string | null {
    for (const preferredName of preferredNames) {
      const resolved = this.resolveNoteFieldName(noteInfo, preferredName);
      if (resolved) return resolved;
    }
    return null;
  }

  private warnFieldParseOnce(fieldName: string, reason: string, detail?: string): void {
    const key = `${fieldName.toLowerCase()}::${reason}`;
    if (this.parseWarningKeys.has(key)) return;
    this.parseWarningKeys.add(key);
    const suffix = detail ? ` (${detail})` : '';
    log.warn(`Field grouping parse warning [${fieldName}] ${reason}${suffix}`);
  }

  private setCardTypeFields(
    updatedFields: Record<string, string>,
    availableFieldNames: string[],
    cardKind: CardKind,
  ): void {
    const audioFlagNames = ['IsAudioCard'];

    if (cardKind === 'sentence') {
      const sentenceFlag = this.resolveFieldName(availableFieldNames, 'IsSentenceCard');
      if (sentenceFlag) {
        updatedFields[sentenceFlag] = 'x';
      }

      for (const audioFlagName of audioFlagNames) {
        const resolved = this.resolveFieldName(availableFieldNames, audioFlagName);
        if (resolved && resolved !== sentenceFlag) {
          updatedFields[resolved] = '';
        }
      }

      const wordAndSentenceFlag = this.resolveFieldName(
        availableFieldNames,
        'IsWordAndSentenceCard',
      );
      if (wordAndSentenceFlag && wordAndSentenceFlag !== sentenceFlag) {
        updatedFields[wordAndSentenceFlag] = '';
      }
      return;
    }

    const resolvedAudioFlags = Array.from(
      new Set(
        audioFlagNames
          .map((name) => this.resolveFieldName(availableFieldNames, name))
          .filter((name): name is string => Boolean(name)),
      ),
    );
    const audioFlagName = resolvedAudioFlags[0] || null;
    if (audioFlagName) {
      updatedFields[audioFlagName] = 'x';
    }
    for (const extraAudioFlag of resolvedAudioFlags.slice(1)) {
      updatedFields[extraAudioFlag] = '';
    }

    const sentenceFlag = this.resolveFieldName(availableFieldNames, 'IsSentenceCard');
    if (sentenceFlag && sentenceFlag !== audioFlagName) {
      updatedFields[sentenceFlag] = '';
    }

    const wordAndSentenceFlag = this.resolveFieldName(availableFieldNames, 'IsWordAndSentenceCard');
    if (wordAndSentenceFlag && wordAndSentenceFlag !== audioFlagName) {
      updatedFields[wordAndSentenceFlag] = '';
    }
  }

  private async showNotification(
    noteId: number,
    label: string | number,
    errorSuffix?: string,
  ): Promise<void> {
    const message = errorSuffix
      ? `Updated card: ${label} (${errorSuffix})`
      : `Updated card: ${label}`;

    const type = this.config.behavior?.notificationType || 'osd';

    if (type === 'osd' || type === 'both') {
      this.showOsdNotification(message);
    }

    if ((type === 'system' || type === 'both') && this.notificationCallback) {
      let notificationIconPath: string | undefined;

      if (this.mpvClient && this.mpvClient.currentVideoPath) {
        try {
          const timestamp = this.mpvClient.currentTimePos || 0;
          const iconBuffer = await this.mediaGenerator.generateNotificationIcon(
            this.mpvClient.currentVideoPath,
            timestamp,
          );
          if (iconBuffer && iconBuffer.length > 0) {
            notificationIconPath = this.mediaGenerator.writeNotificationIconToFile(
              iconBuffer,
              noteId,
            );
          }
        } catch (err) {
          log.warn('Failed to generate notification icon:', (err as Error).message);
        }
      }

      this.notificationCallback('Anki Card Updated', {
        body: message,
        icon: notificationIconPath,
      });

      if (notificationIconPath) {
        this.mediaGenerator.scheduleNotificationIconCleanup(notificationIconPath);
      }
    }
  }

  private mergeFieldValue(existing: string, newValue: string, overwrite: boolean): string {
    if (overwrite || !existing.trim()) {
      return newValue;
    }
    if (this.config.behavior?.mediaInsertMode === 'prepend') {
      return newValue + existing;
    }
    return existing + newValue;
  }

  /**
   * Update the last added Anki card using subtitle blocks from clipboard.
   * This is the manual update flow (animecards-style) when auto-update is disabled.
   */
  async updateLastAddedFromClipboard(clipboardText: string): Promise<void> {
    return this.cardCreationService.updateLastAddedFromClipboard(clipboardText);
  }

  async triggerFieldGroupingForLastAddedCard(): Promise<void> {
    return this.fieldGroupingService.triggerFieldGroupingForLastAddedCard();
  }

  async markLastCardAsAudioCard(): Promise<void> {
    return this.cardCreationService.markLastCardAsAudioCard();
  }

  async createSentenceCard(
    sentence: string,
    startTime: number,
    endTime: number,
    secondarySubText?: string,
  ): Promise<boolean> {
    return this.cardCreationService.createSentenceCard(
      sentence,
      startTime,
      endTime,
      secondarySubText,
    );
  }

  private async findDuplicateNote(
    expression: string,
    excludeNoteId: number,
    noteInfo: NoteInfo,
  ): Promise<number | null> {
    return findDuplicateNoteForAnkiIntegration(expression, excludeNoteId, noteInfo, {
      findNotes: async (query, options) => (await this.client.findNotes(query, options)) as unknown,
      notesInfo: async (noteIds) => (await this.client.notesInfo(noteIds)) as unknown,
      getDeck: () => this.config.deck,
      resolveFieldName: (info, preferredName) => this.resolveNoteFieldName(info, preferredName),
      logWarn: (message, error) => {
        log.warn(message, (error as Error).message);
      },
    });
  }

  private getGroupableFieldNames(): string[] {
    const fields: string[] = [];
    fields.push('Sentence');
    fields.push('SentenceAudio');
    fields.push('Picture');
    if (this.config.fields?.image) fields.push(this.config.fields?.image);
    if (this.config.fields?.sentence) fields.push(this.config.fields?.sentence);
    if (
      this.config.fields?.audio &&
      this.config.fields?.audio.toLowerCase() !== 'expressionaudio'
    ) {
      fields.push(this.config.fields?.audio);
    }
    const sentenceCardConfig = this.getEffectiveSentenceCardConfig();
    const sentenceAudioField = sentenceCardConfig.audioField;
    if (!fields.includes(sentenceAudioField)) fields.push(sentenceAudioField);
    if (this.config.fields?.miscInfo) fields.push(this.config.fields?.miscInfo);
    fields.push('SentenceFurigana');
    return fields;
  }

  private getPreferredSentenceAudioFieldName(): string {
    const sentenceCardConfig = this.getEffectiveSentenceCardConfig();
    return sentenceCardConfig.audioField || 'SentenceAudio';
  }

  private getResolvedSentenceAudioFieldName(noteInfo: NoteInfo): string | null {
    return (
      this.resolveNoteFieldName(noteInfo, this.getPreferredSentenceAudioFieldName()) ||
      this.resolveConfiguredFieldName(noteInfo, this.config.fields?.audio)
    );
  }

  private extractUngroupedValue(value: string): string {
    const groupedSpanRegex = /<span\s+data-group-id="[^"]*">[\s\S]*?<\/span>/gi;
    const ungrouped = value.replace(groupedSpanRegex, '').trim();
    if (ungrouped) return ungrouped;
    return value.trim();
  }

  private extractLastSoundTag(value: string): string {
    const matches = value.match(/\[sound:[^\]]+\]/g);
    if (!matches || matches.length === 0) return '';
    return matches[matches.length - 1];
  }

  private extractLastImageTag(value: string): string {
    const matches = value.match(/<img\b[^>]*>/gi);
    if (!matches || matches.length === 0) return '';
    return matches[matches.length - 1];
  }

  private extractImageTags(value: string): string[] {
    const matches = value.match(/<img\b[^>]*>/gi);
    return matches || [];
  }

  private ensureImageGroupId(imageTag: string, groupId: number): string {
    if (!imageTag) return '';
    if (/data-group-id=/i.test(imageTag)) {
      return imageTag.replace(/data-group-id="[^"]*"/i, `data-group-id="${groupId}"`);
    }
    return imageTag.replace(/<img\b/i, `<img data-group-id="${groupId}"`);
  }

  private extractSpanEntries(
    value: string,
    fieldName: string,
  ): { groupId: number; content: string }[] {
    const entries: { groupId: number; content: string }[] = [];
    const malformedIdRegex = /<span\s+[^>]*data-group-id="([^"]*)"[^>]*>/gi;
    let malformed;
    while ((malformed = malformedIdRegex.exec(value)) !== null) {
      const rawId = malformed[1];
      const groupId = Number(rawId);
      if (!Number.isFinite(groupId) || groupId <= 0) {
        this.warnFieldParseOnce(fieldName, 'invalid-group-id', rawId);
      }
    }

    const spanRegex = /<span\s+data-group-id="(\d+)"[^>]*>([\s\S]*?)<\/span>/gi;
    let match;
    while ((match = spanRegex.exec(value)) !== null) {
      const groupId = Number(match[1]);
      if (!Number.isFinite(groupId) || groupId <= 0) continue;
      const content = this.normalizeStrictGroupedValue(match[2] || '', fieldName);
      if (!content) {
        this.warnFieldParseOnce(fieldName, 'empty-group-content');
        log.debug('Skipping span with empty normalized content', {
          fieldName,
          rawContent: (match[2] || '').slice(0, 120),
        });
        continue;
      }
      entries.push({ groupId, content });
    }
    if (entries.length === 0 && /<span\b/i.test(value)) {
      this.warnFieldParseOnce(fieldName, 'no-usable-span-entries');
    }
    return entries;
  }

  private parseStrictEntries(
    value: string,
    fallbackGroupId: number,
    fieldName: string,
  ): { groupId: number; content: string }[] {
    const entries = this.extractSpanEntries(value, fieldName);
    if (entries.length === 0) {
      const ungrouped = this.normalizeStrictGroupedValue(
        this.extractUngroupedValue(value),
        fieldName,
      );
      if (ungrouped) {
        entries.push({ groupId: fallbackGroupId, content: ungrouped });
      }
    }

    const unique: { groupId: number; content: string }[] = [];
    const seen = new Set<string>();
    for (const entry of entries) {
      const key = `${entry.groupId}::${entry.content}`;
      if (seen.has(key)) continue;
      seen.add(key);
      unique.push(entry);
    }
    return unique;
  }

  private parsePictureEntries(
    value: string,
    fallbackGroupId: number,
  ): { groupId: number; tag: string }[] {
    const tags = this.extractImageTags(value);
    const result: { groupId: number; tag: string }[] = [];
    for (const tag of tags) {
      const idMatch = tag.match(/data-group-id="(\d+)"/i);
      let groupId = fallbackGroupId;
      if (idMatch) {
        const parsed = Number(idMatch[1]);
        if (!Number.isFinite(parsed) || parsed <= 0) {
          this.warnFieldParseOnce('Picture', 'invalid-group-id', idMatch[1]);
        } else {
          groupId = parsed;
        }
      }
      const normalizedTag = this.ensureImageGroupId(tag, groupId);
      if (!normalizedTag) {
        this.warnFieldParseOnce('Picture', 'empty-image-tag');
        continue;
      }
      result.push({ groupId, tag: normalizedTag });
    }
    return result;
  }

  private normalizeStrictGroupedValue(value: string, fieldName: string): string {
    const ungrouped = this.extractUngroupedValue(value);
    if (!ungrouped) return '';

    const normalizedField = fieldName.toLowerCase();
    if (normalizedField === 'sentenceaudio' || normalizedField === 'expressionaudio') {
      const lastSoundTag = this.extractLastSoundTag(ungrouped);
      if (!lastSoundTag) {
        this.warnFieldParseOnce(fieldName, 'missing-sound-tag');
      }
      return lastSoundTag || ungrouped;
    }

    if (normalizedField === 'picture') {
      const lastImageTag = this.extractLastImageTag(ungrouped);
      if (!lastImageTag) {
        this.warnFieldParseOnce(fieldName, 'missing-image-tag');
      }
      return lastImageTag || ungrouped;
    }

    return ungrouped;
  }

  private getStrictSpanGroupingFields(): Set<string> {
    const strictFields = new Set(this.strictGroupingFieldDefaults);
    const sentenceCardConfig = this.getEffectiveSentenceCardConfig();
    strictFields.add((sentenceCardConfig.sentenceField || 'sentence').toLowerCase());
    strictFields.add((sentenceCardConfig.audioField || 'sentenceaudio').toLowerCase());
    if (this.config.fields?.image) strictFields.add(this.config.fields.image.toLowerCase());
    if (this.config.fields?.miscInfo) strictFields.add(this.config.fields.miscInfo.toLowerCase());
    return strictFields;
  }

  private shouldUseStrictSpanGrouping(fieldName: string): boolean {
    const normalized = fieldName.toLowerCase();
    return this.getStrictSpanGroupingFields().has(normalized);
  }

  private applyFieldGrouping(
    existingValue: string,
    newValue: string,
    keepGroupId: number,
    sourceGroupId: number,
    fieldName: string,
  ): string {
    if (this.shouldUseStrictSpanGrouping(fieldName)) {
      if (fieldName.toLowerCase() === 'picture') {
        const keepEntries = this.parsePictureEntries(existingValue, keepGroupId);
        const sourceEntries = this.parsePictureEntries(newValue, sourceGroupId);
        if (keepEntries.length === 0 && sourceEntries.length === 0) {
          return existingValue || newValue;
        }
        const mergedTags = keepEntries.map((entry) =>
          this.ensureImageGroupId(entry.tag, entry.groupId),
        );
        const seen = new Set(mergedTags);
        for (const entry of sourceEntries) {
          const normalized = this.ensureImageGroupId(entry.tag, entry.groupId);
          if (seen.has(normalized)) continue;
          seen.add(normalized);
          mergedTags.push(normalized);
        }
        return mergedTags.join('');
      }

      const keepEntries = this.parseStrictEntries(existingValue, keepGroupId, fieldName);
      const sourceEntries = this.parseStrictEntries(newValue, sourceGroupId, fieldName);
      if (keepEntries.length === 0 && sourceEntries.length === 0) {
        return existingValue || newValue;
      }
      if (sourceEntries.length === 0) {
        return keepEntries
          .map((entry) => `<span data-group-id="${entry.groupId}">${entry.content}</span>`)
          .join('');
      }
      const merged = [...keepEntries];
      const seen = new Set(keepEntries.map((entry) => `${entry.groupId}::${entry.content}`));
      for (const entry of sourceEntries) {
        const key = `${entry.groupId}::${entry.content}`;
        if (seen.has(key)) continue;
        seen.add(key);
        merged.push(entry);
      }
      if (merged.length === 0) return existingValue;
      return merged
        .map((entry) => `<span data-group-id="${entry.groupId}">${entry.content}</span>`)
        .join('');
    }

    if (!existingValue.trim()) return newValue;
    if (!newValue.trim()) return existingValue;

    const hasGroups = /data-group-id/.test(existingValue);

    if (!hasGroups) {
      return `<span data-group-id="${keepGroupId}">${existingValue}</span>\n` + newValue;
    }

    const groupedSpanRegex = /<span\s+data-group-id="[^"]*">[\s\S]*?<\/span>/g;
    let lastEnd = 0;
    let result = '';
    let match;

    while ((match = groupedSpanRegex.exec(existingValue)) !== null) {
      const before = existingValue.slice(lastEnd, match.index);
      if (before.trim()) {
        result += `<span data-group-id="${keepGroupId}">${before.trim()}</span>\n`;
      }
      result += match[0] + '\n';
      lastEnd = match.index + match[0].length;
    }

    const after = existingValue.slice(lastEnd);
    if (after.trim()) {
      result += `\n<span data-group-id="${keepGroupId}">${after.trim()}</span>`;
    }

    return result + '\n' + newValue;
  }

  private async generateMediaForMerge(): Promise<{
    audioField?: string;
    audioValue?: string;
    imageField?: string;
    imageValue?: string;
    miscInfoValue?: string;
  }> {
    const result: {
      audioField?: string;
      audioValue?: string;
      imageField?: string;
      imageValue?: string;
      miscInfoValue?: string;
    } = {};

    if (this.config.media?.generateAudio && this.mpvClient?.currentVideoPath) {
      try {
        const audioFilename = this.generateAudioFilename();
        const audioBuffer = await this.generateAudio();
        if (audioBuffer) {
          await this.client.storeMediaFile(audioFilename, audioBuffer);
          result.audioField = this.getPreferredSentenceAudioFieldName();
          result.audioValue = `[sound:${audioFilename}]`;
          if (this.config.fields?.miscInfo) {
            result.miscInfoValue = this.formatMiscInfoPattern(
              audioFilename,
              this.mpvClient.currentSubStart,
            );
          }
        }
      } catch (error) {
        log.error('Failed to generate audio for merge:', (error as Error).message);
      }
    }

    if (this.config.media?.generateImage && this.mpvClient?.currentVideoPath) {
      try {
        const imageFilename = this.generateImageFilename();
        const imageBuffer = await this.generateImage();
        if (imageBuffer) {
          await this.client.storeMediaFile(imageFilename, imageBuffer);
          result.imageField = this.config.fields?.image || DEFAULT_ANKI_CONNECT_CONFIG.fields.image;
          result.imageValue = `<img src="${imageFilename}">`;
          if (this.config.fields?.miscInfo && !result.miscInfoValue) {
            result.miscInfoValue = this.formatMiscInfoPattern(
              imageFilename,
              this.mpvClient.currentSubStart,
            );
          }
        }
      } catch (error) {
        log.error('Failed to generate image for merge:', (error as Error).message);
      }
    }

    return result;
  }

  private getResolvedFieldValue(noteInfo: NoteInfo, preferredFieldName?: string): string {
    if (!preferredFieldName) return '';
    const resolved = this.resolveNoteFieldName(noteInfo, preferredFieldName);
    if (!resolved) return '';
    return noteInfo.fields[resolved]?.value || '';
  }

  private async computeFieldGroupingMergedFields(
    keepNoteId: number,
    deleteNoteId: number,
    keepNoteInfo: NoteInfo,
    deleteNoteInfo: NoteInfo,
    includeGeneratedMedia: boolean,
  ): Promise<Record<string, string>> {
    const groupableFields = this.getGroupableFieldNames();
    const keepFieldNames = Object.keys(keepNoteInfo.fields);
    const sourceFields: Record<string, string> = {};
    const resolvedKeepFieldByPreferred = new Map<string, string>();
    for (const preferredFieldName of groupableFields) {
      sourceFields[preferredFieldName] = this.getResolvedFieldValue(
        deleteNoteInfo,
        preferredFieldName,
      );
      const keepResolved = this.resolveFieldName(keepFieldNames, preferredFieldName);
      if (keepResolved) {
        resolvedKeepFieldByPreferred.set(preferredFieldName, keepResolved);
      }
    }

    if (!sourceFields['SentenceFurigana'] && sourceFields['Sentence']) {
      sourceFields['SentenceFurigana'] = sourceFields['Sentence'];
    }
    if (!sourceFields['Sentence'] && sourceFields['SentenceFurigana']) {
      sourceFields['Sentence'] = sourceFields['SentenceFurigana'];
    }
    if (!sourceFields['Expression'] && sourceFields['Word']) {
      sourceFields['Expression'] = sourceFields['Word'];
    }
    if (!sourceFields['Word'] && sourceFields['Expression']) {
      sourceFields['Word'] = sourceFields['Expression'];
    }
    if (!sourceFields['SentenceAudio'] && sourceFields['ExpressionAudio']) {
      sourceFields['SentenceAudio'] = sourceFields['ExpressionAudio'];
    }
    if (!sourceFields['ExpressionAudio'] && sourceFields['SentenceAudio']) {
      sourceFields['ExpressionAudio'] = sourceFields['SentenceAudio'];
    }

    if (
      this.config.fields?.sentence &&
      !sourceFields[this.config.fields?.sentence] &&
      this.mpvClient.currentSubText
    ) {
      const deleteFields = this.extractFields(deleteNoteInfo.fields);
      sourceFields[this.config.fields?.sentence] = this.processSentence(
        this.mpvClient.currentSubText,
        deleteFields,
      );
    }

    if (includeGeneratedMedia) {
      const media = await this.generateMediaForMerge();
      if (media.audioField && media.audioValue && !sourceFields[media.audioField]) {
        sourceFields[media.audioField] = media.audioValue;
      }
      if (media.imageField && media.imageValue && !sourceFields[media.imageField]) {
        sourceFields[media.imageField] = media.imageValue;
      }
      if (
        this.config.fields?.miscInfo &&
        media.miscInfoValue &&
        !sourceFields[this.config.fields?.miscInfo]
      ) {
        sourceFields[this.config.fields?.miscInfo] = media.miscInfoValue;
      }
    }

    const mergedFields: Record<string, string> = {};
    for (const preferredFieldName of groupableFields) {
      const keepFieldName = resolvedKeepFieldByPreferred.get(preferredFieldName);
      if (!keepFieldName) continue;

      const keepFieldNormalized = keepFieldName.toLowerCase();
      if (
        keepFieldNormalized === 'expression' ||
        keepFieldNormalized === 'expressionfurigana' ||
        keepFieldNormalized === 'expressionreading' ||
        keepFieldNormalized === 'expressionaudio'
      ) {
        continue;
      }

      const existingValue = keepNoteInfo.fields[keepFieldName]?.value || '';
      const newValue = sourceFields[preferredFieldName] || '';
      const isStrictField = this.shouldUseStrictSpanGrouping(keepFieldName);
      if (!existingValue.trim() && !newValue.trim()) continue;

      if (isStrictField) {
        mergedFields[keepFieldName] = this.applyFieldGrouping(
          existingValue,
          newValue,
          keepNoteId,
          deleteNoteId,
          keepFieldName,
        );
      } else if (existingValue.trim() && newValue.trim()) {
        mergedFields[keepFieldName] = this.applyFieldGrouping(
          existingValue,
          newValue,
          keepNoteId,
          deleteNoteId,
          keepFieldName,
        );
      } else {
        if (!newValue.trim()) continue;
        mergedFields[keepFieldName] = newValue;
      }
    }

    // Keep sentence/expression audio fields aligned after grouping. Otherwise a
    // kept note can retain stale ExpressionAudio while SentenceAudio is merged.
    const sentenceCardConfig = this.getEffectiveSentenceCardConfig();
    const resolvedSentenceAudioField = this.resolveFieldName(
      keepFieldNames,
      sentenceCardConfig.audioField || 'SentenceAudio',
    );
    const resolvedExpressionAudioField = this.resolveFieldName(
      keepFieldNames,
      this.config.fields?.audio || 'ExpressionAudio',
    );
    if (
      resolvedSentenceAudioField &&
      resolvedExpressionAudioField &&
      resolvedExpressionAudioField !== resolvedSentenceAudioField
    ) {
      const mergedSentenceAudioValue =
        mergedFields[resolvedSentenceAudioField] ||
        keepNoteInfo.fields[resolvedSentenceAudioField]?.value ||
        '';
      if (mergedSentenceAudioValue.trim()) {
        mergedFields[resolvedExpressionAudioField] = mergedSentenceAudioValue;
      }
    }

    return mergedFields;
  }

  private getNoteFieldMap(noteInfo: NoteInfo): Record<string, string> {
    const fields: Record<string, string> = {};
    for (const [name, field] of Object.entries(noteInfo.fields)) {
      fields[name] = field?.value || '';
    }
    return fields;
  }

  async buildFieldGroupingPreview(
    keepNoteId: number,
    deleteNoteId: number,
    deleteDuplicate: boolean,
  ): Promise<KikuMergePreviewResponse> {
    return this.fieldGroupingService.buildFieldGroupingPreview(
      keepNoteId,
      deleteNoteId,
      deleteDuplicate,
    );
  }

  private async performFieldGroupingMerge(
    keepNoteId: number,
    deleteNoteId: number,
    deleteNoteInfo: NoteInfo,
    expression: string,
    deleteDuplicate = true,
  ): Promise<void> {
    const keepNotesInfoResult = await this.client.notesInfo([keepNoteId]);
    const keepNotesInfo = keepNotesInfoResult as unknown as NoteInfo[];
    if (!keepNotesInfo || keepNotesInfo.length === 0) {
      log.warn('Keep note not found:', keepNoteId);
      return;
    }
    const keepNoteInfo = keepNotesInfo[0];
    const mergedFields = await this.computeFieldGroupingMergedFields(
      keepNoteId,
      deleteNoteId,
      keepNoteInfo,
      deleteNoteInfo,
      true,
    );

    if (Object.keys(mergedFields).length > 0) {
      await this.client.updateNoteFields(keepNoteId, mergedFields);
      await this.addConfiguredTagsToNote(keepNoteId);
    }

    if (deleteDuplicate) {
      await this.client.deleteNotes([deleteNoteId]);
      this.previousNoteIds.delete(deleteNoteId);
    }

    log.info('Merged duplicate card:', expression, 'into note:', keepNoteId);
    this.showStatusNotification(
      deleteDuplicate
        ? `Merged duplicate: ${expression}`
        : `Grouped duplicate (kept both): ${expression}`,
    );
    await this.showNotification(keepNoteId, expression);
  }

  private async handleFieldGroupingAuto(
    originalNoteId: number,
    newNoteId: number,
    newNoteInfo: NoteInfo,
    expression: string,
  ): Promise<void> {
    try {
      const sentenceCardConfig = this.getEffectiveSentenceCardConfig();
      await this.performFieldGroupingMerge(
        originalNoteId,
        newNoteId,
        newNoteInfo,
        expression,
        sentenceCardConfig.kikuDeleteDuplicateInAuto,
      );
    } catch (error) {
      log.error('Field grouping auto merge failed:', (error as Error).message);
      this.showOsdNotification(`Field grouping failed: ${(error as Error).message}`);
    }
  }

  private async handleFieldGroupingManual(
    originalNoteId: number,
    newNoteId: number,
    newNoteInfo: NoteInfo,
    expression: string,
  ): Promise<boolean> {
    if (!this.fieldGroupingCallback) {
      log.warn('No field grouping callback registered, skipping manual mode');
      this.showOsdNotification('Field grouping UI unavailable');
      return false;
    }

    try {
      const originalNotesInfoResult = await this.client.notesInfo([originalNoteId]);
      const originalNotesInfo = originalNotesInfoResult as unknown as NoteInfo[];
      if (!originalNotesInfo || originalNotesInfo.length === 0) {
        return false;
      }
      const originalNoteInfo = originalNotesInfo[0];
      const sentenceCardConfig = this.getEffectiveSentenceCardConfig();

      const originalFields = this.extractFields(originalNoteInfo.fields);
      const newFields = this.extractFields(newNoteInfo.fields);

      const originalCard: KikuDuplicateCardInfo = {
        noteId: originalNoteId,
        expression: originalFields.expression || originalFields.word || expression,
        sentencePreview: this.truncateSentence(
          originalFields[(sentenceCardConfig.sentenceField || 'sentence').toLowerCase()] || '',
        ),
        hasAudio:
          this.hasFieldValue(originalNoteInfo, this.config.fields?.audio) ||
          this.hasFieldValue(originalNoteInfo, sentenceCardConfig.audioField),
        hasImage: this.hasFieldValue(originalNoteInfo, this.config.fields?.image),
        isOriginal: true,
      };

      const newCard: KikuDuplicateCardInfo = {
        noteId: newNoteId,
        expression: newFields.expression || newFields.word || expression,
        sentencePreview: this.truncateSentence(
          newFields[(sentenceCardConfig.sentenceField || 'sentence').toLowerCase()] ||
            this.mpvClient.currentSubText ||
            '',
        ),
        hasAudio:
          this.hasFieldValue(newNoteInfo, this.config.fields?.audio) ||
          this.hasFieldValue(newNoteInfo, sentenceCardConfig.audioField),
        hasImage: this.hasFieldValue(newNoteInfo, this.config.fields?.image),
        isOriginal: false,
      };

      const choice = await this.fieldGroupingCallback({
        original: originalCard,
        duplicate: newCard,
      });

      if (choice.cancelled) {
        this.showOsdNotification('Field grouping cancelled');
        return false;
      }

      const keepNoteId = choice.keepNoteId;
      const deleteNoteId = choice.deleteNoteId;
      const deleteNoteInfo = deleteNoteId === newNoteId ? newNoteInfo : originalNoteInfo;

      await this.performFieldGroupingMerge(
        keepNoteId,
        deleteNoteId,
        deleteNoteInfo,
        expression,
        choice.deleteDuplicate,
      );
      return true;
    } catch (error) {
      log.error('Field grouping manual merge failed:', (error as Error).message);
      this.showOsdNotification(`Field grouping failed: ${(error as Error).message}`);
      return false;
    }
  }

  private truncateSentence(sentence: string): string {
    const clean = sentence.replace(/<[^>]*>/g, '').trim();
    if (clean.length <= 100) return clean;
    return clean.substring(0, 100) + '...';
  }

  private hasFieldValue(noteInfo: NoteInfo, preferredFieldName?: string): boolean {
    const resolved = this.resolveNoteFieldName(noteInfo, preferredFieldName);
    if (!resolved) return false;
    return Boolean(noteInfo.fields[resolved]?.value);
  }

  private hasAllConfiguredFields(
    noteInfo: NoteInfo,
    configuredFieldNames: (string | undefined)[],
  ): boolean {
    const requiredFields = configuredFieldNames.filter((fieldName): fieldName is string =>
      Boolean(fieldName),
    );
    if (requiredFields.length === 0) return true;
    return requiredFields.every((fieldName) => this.hasFieldValue(noteInfo, fieldName));
  }

  private async refreshMiscInfoField(noteId: number, noteInfo: NoteInfo): Promise<void> {
    if (!this.config.fields?.miscInfo || !this.config.metadata?.pattern) return;

    const resolvedMiscField = this.resolveNoteFieldName(noteInfo, this.config.fields?.miscInfo);
    if (!resolvedMiscField) return;

    const nextValue = this.formatMiscInfoPattern('', this.mpvClient.currentSubStart);
    if (!nextValue) return;

    const currentValue = noteInfo.fields[resolvedMiscField]?.value || '';
    if (currentValue === nextValue) return;

    await this.client.updateNoteFields(noteId, {
      [resolvedMiscField]: nextValue,
    });
    await this.addConfiguredTagsToNote(noteId);
  }

  applyRuntimeConfigPatch(patch: Partial<AnkiConnectConfig>): void {
    const wasEnabled = this.config.nPlusOne?.highlightEnabled === true;
    const previousPollingRate = this.config.pollingRate;
    this.config = {
      ...this.config,
      ...patch,
      nPlusOne:
        patch.nPlusOne !== undefined
          ? {
              ...(this.config.nPlusOne ?? DEFAULT_ANKI_CONNECT_CONFIG.nPlusOne),
              ...patch.nPlusOne,
            }
          : this.config.nPlusOne,
      fields:
        patch.fields !== undefined
          ? { ...this.config.fields, ...patch.fields }
          : this.config.fields,
      media:
        patch.media !== undefined ? { ...this.config.media, ...patch.media } : this.config.media,
      behavior:
        patch.behavior !== undefined
          ? { ...this.config.behavior, ...patch.behavior }
          : this.config.behavior,
      metadata:
        patch.metadata !== undefined
          ? { ...this.config.metadata, ...patch.metadata }
          : this.config.metadata,
      isLapis:
        patch.isLapis !== undefined
          ? { ...this.config.isLapis, ...patch.isLapis }
          : this.config.isLapis,
      isKiku:
        patch.isKiku !== undefined
          ? { ...this.config.isKiku, ...patch.isKiku }
          : this.config.isKiku,
    };

    if (wasEnabled && this.config.nPlusOne?.highlightEnabled === false) {
      this.stopKnownWordCacheLifecycle();
      this.knownWordCache.clearKnownWordCacheState();
    } else {
      this.startKnownWordCacheLifecycle();
    }

    if (
      patch.pollingRate !== undefined &&
      previousPollingRate !== this.config.pollingRate &&
      this.pollingRunner.isRunning
    ) {
      this.pollingRunner.start();
    }
  }

  destroy(): void {
    this.stop();
    this.mediaGenerator.cleanup();
  }
}
