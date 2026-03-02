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
  showStatusNotification,
  withUpdateProgress,
  UiFeedbackState,
} from './anki-integration/ui-feedback';
import { KnownWordCacheManager } from './anki-integration/known-word-cache';
import { PollingRunner } from './anki-integration/polling';
import type { AnkiConnectProxyServer } from './anki-integration/anki-connect-proxy';
import { findDuplicateNote as findDuplicateNoteForAnkiIntegration } from './anki-integration/duplicate';
import { CardCreationService } from './anki-integration/card-creation';
import { FieldGroupingService } from './anki-integration/field-grouping';
import { FieldGroupingMergeCollaborator } from './anki-integration/field-grouping-merge';
import { NoteUpdateWorkflow } from './anki-integration/note-update-workflow';
import { FieldGroupingWorkflow } from './anki-integration/field-grouping-workflow';

const log = createLogger('anki').child('integration');

interface NoteInfo {
  noteId: number;
  fields: Record<string, { value: string }>;
}

type CardKind = 'sentence' | 'audio';

function trimToNonEmptyString(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function decodeURIComponentSafe(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

function extractFilenameFromMediaPath(rawPath: string): string {
  const trimmedPath = rawPath.trim();
  if (!trimmedPath) return '';

  if (/^[a-zA-Z][a-zA-Z\d+\-.]*:\/\//.test(trimmedPath)) {
    try {
      const parsed = new URL(trimmedPath);
      return decodeURIComponentSafe(path.basename(parsed.pathname));
    } catch {
      // Fall through to separator-based handling below.
    }
  }

  const separatorIndex = trimmedPath.search(/[?#]/);
  const pathWithoutQuery =
    separatorIndex >= 0 ? trimmedPath.slice(0, separatorIndex) : trimmedPath;
  return decodeURIComponentSafe(path.basename(pathWithoutQuery));
}

function shouldPreferMediaTitleForMiscInfo(rawPath: string, filename: string): boolean {
  const loweredPath = rawPath.toLowerCase();
  const loweredFilename = filename.toLowerCase();
  if (loweredPath.includes('api_key=')) {
    return true;
  }
  if (loweredPath.startsWith('http://') || loweredPath.startsWith('https://')) {
    return true;
  }
  return (
    loweredFilename === 'stream' ||
    loweredFilename === 'master.m3u8' ||
    loweredFilename === 'index.m3u8' ||
    loweredFilename === 'playlist.m3u8'
  );
}

export class AnkiIntegration {
  private client: AnkiConnectClient;
  private mediaGenerator: MediaGenerator;
  private timingTracker: SubtitleTimingTracker;
  private config: AnkiConnectConfig;
  private pollingRunner!: PollingRunner;
  private proxyServer: AnkiConnectProxyServer | null = null;
  private started = false;
  private previousNoteIds = new Set<number>();
  private mpvClient: MpvClient;
  private osdCallback: ((text: string) => void) | null = null;
  private notificationCallback: ((title: string, options: NotificationOptions) => void) | null =
    null;
  private updateInProgress = false;
  private uiFeedbackState: UiFeedbackState = createUiFeedbackState();
  private parseWarningKeys = new Set<string>();
  private fieldGroupingCallback:
    | ((data: {
        original: KikuDuplicateCardInfo;
        duplicate: KikuDuplicateCardInfo;
      }) => Promise<KikuFieldGroupingChoice>)
    | null = null;
  private knownWordCache: KnownWordCacheManager;
  private cardCreationService: CardCreationService;
  private fieldGroupingMergeCollaborator: FieldGroupingMergeCollaborator;
  private fieldGroupingService: FieldGroupingService;
  private noteUpdateWorkflow: NoteUpdateWorkflow;
  private fieldGroupingWorkflow: FieldGroupingWorkflow;

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
    this.config = this.normalizeConfig(config);
    this.client = new AnkiConnectClient(this.config.url!);
    this.mediaGenerator = new MediaGenerator();
    this.timingTracker = timingTracker;
    this.mpvClient = mpvClient;
    this.osdCallback = osdCallback || null;
    this.notificationCallback = notificationCallback || null;
    this.fieldGroupingCallback = fieldGroupingCallback || null;
    this.knownWordCache = this.createKnownWordCache(knownWordCacheStatePath);
    this.pollingRunner = this.createPollingRunner();
    this.cardCreationService = this.createCardCreationService();
    this.fieldGroupingMergeCollaborator = this.createFieldGroupingMergeCollaborator();
    this.fieldGroupingService = this.createFieldGroupingService();
    this.noteUpdateWorkflow = this.createNoteUpdateWorkflow();
    this.fieldGroupingWorkflow = this.createFieldGroupingWorkflow();
  }

  private createFieldGroupingMergeCollaborator(): FieldGroupingMergeCollaborator {
    return new FieldGroupingMergeCollaborator({
      getConfig: () => this.config,
      getEffectiveSentenceCardConfig: () => this.getEffectiveSentenceCardConfig(),
      getCurrentSubtitleText: () => this.mpvClient.currentSubText,
      resolveFieldName: (availableFieldNames, preferredName) =>
        this.resolveFieldName(availableFieldNames, preferredName),
      resolveNoteFieldName: (noteInfo, preferredName) =>
        this.resolveNoteFieldName(noteInfo, preferredName),
      extractFields: (fields) => this.extractFields(fields),
      processSentence: (mpvSentence, noteFields) => this.processSentence(mpvSentence, noteFields),
      generateMediaForMerge: () => this.generateMediaForMerge(),
      warnFieldParseOnce: (fieldName, reason, detail) =>
        this.warnFieldParseOnce(fieldName, reason, detail),
    });
  }

  private normalizeConfig(config: AnkiConnectConfig): AnkiConnectConfig {
    const resolvedUrl =
      typeof config.url === 'string' && config.url.trim().length > 0
        ? config.url.trim()
        : DEFAULT_ANKI_CONNECT_CONFIG.url;
    const proxySource =
      config.proxy && typeof config.proxy === 'object'
        ? (config.proxy as NonNullable<AnkiConnectConfig['proxy']>)
        : {};
    const normalizedProxyPort =
      typeof proxySource.port === 'number' &&
      Number.isInteger(proxySource.port) &&
      proxySource.port >= 1 &&
      proxySource.port <= 65535
        ? proxySource.port
        : DEFAULT_ANKI_CONNECT_CONFIG.proxy?.port;
    const normalizedProxyHost =
      typeof proxySource.host === 'string' && proxySource.host.trim().length > 0
        ? proxySource.host.trim()
        : DEFAULT_ANKI_CONNECT_CONFIG.proxy?.host;
    const normalizedProxyUpstreamUrl =
      typeof proxySource.upstreamUrl === 'string' && proxySource.upstreamUrl.trim().length > 0
        ? proxySource.upstreamUrl.trim()
        : resolvedUrl;

    return {
      ...DEFAULT_ANKI_CONNECT_CONFIG,
      ...config,
      url: resolvedUrl,
      fields: {
        ...DEFAULT_ANKI_CONNECT_CONFIG.fields,
        ...(config.fields ?? {}),
      },
      proxy: {
        ...DEFAULT_ANKI_CONNECT_CONFIG.proxy,
        ...(config.proxy ?? {}),
        enabled: proxySource.enabled === true,
        host: normalizedProxyHost,
        port: normalizedProxyPort,
        upstreamUrl: normalizedProxyUpstreamUrl,
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
  }

  private createKnownWordCache(knownWordCacheStatePath?: string): KnownWordCacheManager {
    return new KnownWordCacheManager({
      client: {
        findNotes: async (query, options) =>
          (await this.client.findNotes(query, options)) as unknown,
        notesInfo: async (noteIds) => (await this.client.notesInfo(noteIds)) as unknown,
      },
      getConfig: () => this.config,
      knownWordCacheStatePath,
      showStatusNotification: (message: string) => this.showStatusNotification(message),
    });
  }

  private createPollingRunner(): PollingRunner {
    return new PollingRunner({
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
  }

  private createProxyServer(): AnkiConnectProxyServer {
    const { AnkiConnectProxyServer } =
      require('./anki-integration/anki-connect-proxy') as typeof import('./anki-integration/anki-connect-proxy');
    return new AnkiConnectProxyServer({
      shouldAutoUpdateNewCards: () => this.config.behavior?.autoUpdateNewCards !== false,
      processNewCard: (noteId: number) => this.processNewCard(noteId),
      getDeck: () => this.config.deck,
      findNotes: async (query, options) =>
        (await this.client.findNotes(query, options)) as number[],
      logInfo: (message, ...args) => log.info(message, ...args),
      logWarn: (message, ...args) => log.warn(message, ...args),
      logError: (message, ...args) => log.error(message, ...args),
    });
  }

  private getOrCreateProxyServer(): AnkiConnectProxyServer {
    if (!this.proxyServer) {
      this.proxyServer = this.createProxyServer();
    }
    return this.proxyServer;
  }

  private createCardCreationService(): CardCreationService {
    return new CardCreationService({
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
  }

  private createFieldGroupingService(): FieldGroupingService {
    return new FieldGroupingService({
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
        this.fieldGroupingMergeCollaborator.computeFieldGroupingMergedFields(
          keepNoteId,
          deleteNoteId,
          keepNoteInfo,
          deleteNoteInfo,
          includeGeneratedMedia,
        ),
      getNoteFieldMap: (noteInfo) => this.fieldGroupingMergeCollaborator.getNoteFieldMap(noteInfo),
      handleFieldGroupingAuto: (originalNoteId, newNoteId, newNoteInfo, expression) =>
        this.handleFieldGroupingAuto(originalNoteId, newNoteId, newNoteInfo, expression),
      handleFieldGroupingManual: (originalNoteId, newNoteId, newNoteInfo, expression) =>
        this.handleFieldGroupingManual(originalNoteId, newNoteId, newNoteInfo, expression),
    });
  }

  private createNoteUpdateWorkflow(): NoteUpdateWorkflow {
    return new NoteUpdateWorkflow({
      client: {
        notesInfo: async (noteIds) => (await this.client.notesInfo(noteIds)) as unknown,
        updateNoteFields: (noteId, fields) => this.client.updateNoteFields(noteId, fields),
        storeMediaFile: (filename, data) => this.client.storeMediaFile(filename, data),
      },
      getConfig: () => this.config,
      getCurrentSubtitleText: () => this.mpvClient.currentSubText,
      getCurrentSubtitleStart: () => this.mpvClient.currentSubStart,
      getEffectiveSentenceCardConfig: () => this.getEffectiveSentenceCardConfig(),
      appendKnownWordsFromNoteInfo: (noteInfo) => this.appendKnownWordsFromNoteInfo(noteInfo),
      extractFields: (fields) => this.extractFields(fields),
      findDuplicateNote: (expression, excludeNoteId, noteInfo) =>
        this.findDuplicateNote(expression, excludeNoteId, noteInfo),
      handleFieldGroupingAuto: (originalNoteId, newNoteId, newNoteInfo, expression) =>
        this.handleFieldGroupingAuto(originalNoteId, newNoteId, newNoteInfo, expression),
      handleFieldGroupingManual: (originalNoteId, newNoteId, newNoteInfo, expression) =>
        this.handleFieldGroupingManual(originalNoteId, newNoteId, newNoteInfo, expression),
      processSentence: (mpvSentence, noteFields) => this.processSentence(mpvSentence, noteFields),
      resolveConfiguredFieldName: (noteInfo, ...preferredNames) =>
        this.resolveConfiguredFieldName(noteInfo, ...preferredNames),
      getResolvedSentenceAudioFieldName: (noteInfo) =>
        this.getResolvedSentenceAudioFieldName(noteInfo),
      mergeFieldValue: (existing, newValue, overwrite) =>
        this.mergeFieldValue(existing, newValue, overwrite),
      generateAudioFilename: () => this.generateAudioFilename(),
      generateAudio: () => this.generateAudio(),
      generateImageFilename: () => this.generateImageFilename(),
      generateImage: () => this.generateImage(),
      formatMiscInfoPattern: (fallbackFilename, startTimeSeconds) =>
        this.formatMiscInfoPattern(fallbackFilename, startTimeSeconds),
      addConfiguredTagsToNote: (noteId) => this.addConfiguredTagsToNote(noteId),
      showNotification: (noteId, label) => this.showNotification(noteId, label),
      showOsdNotification: (message) => this.showOsdNotification(message),
      beginUpdateProgress: (initialMessage) => this.beginUpdateProgress(initialMessage),
      endUpdateProgress: () => this.endUpdateProgress(),
      logWarn: (...args) => log.warn(args[0] as string, ...args.slice(1)),
      logInfo: (...args) => log.info(args[0] as string, ...args.slice(1)),
      logError: (...args) => log.error(args[0] as string, ...args.slice(1)),
    });
  }

  private createFieldGroupingWorkflow(): FieldGroupingWorkflow {
    return new FieldGroupingWorkflow({
      client: {
        notesInfo: async (noteIds) => (await this.client.notesInfo(noteIds)) as unknown,
        updateNoteFields: (noteId, fields) => this.client.updateNoteFields(noteId, fields),
        deleteNotes: (noteIds) => this.client.deleteNotes(noteIds),
      },
      getConfig: () => this.config,
      getEffectiveSentenceCardConfig: () => this.getEffectiveSentenceCardConfig(),
      getCurrentSubtitleText: () => this.mpvClient.currentSubText,
      getFieldGroupingCallback: () => this.fieldGroupingCallback,
      computeFieldGroupingMergedFields: (
        keepNoteId,
        deleteNoteId,
        keepNoteInfo,
        deleteNoteInfo,
        includeGeneratedMedia,
      ) =>
        this.fieldGroupingMergeCollaborator.computeFieldGroupingMergedFields(
          keepNoteId,
          deleteNoteId,
          keepNoteInfo,
          deleteNoteInfo,
          includeGeneratedMedia,
        ),
      extractFields: (fields) => this.extractFields(fields),
      hasFieldValue: (noteInfo, preferredFieldName) =>
        this.hasFieldValue(noteInfo, preferredFieldName),
      addConfiguredTagsToNote: (noteId) => this.addConfiguredTagsToNote(noteId),
      removeTrackedNoteId: (noteId) => {
        this.previousNoteIds.delete(noteId);
      },
      showStatusNotification: (message) => this.showStatusNotification(message),
      showNotification: (noteId, label) => this.showNotification(noteId, label),
      showOsdNotification: (message) => this.showOsdNotification(message),
      logError: (...args) => log.error(args[0] as string, ...args.slice(1)),
      logInfo: (...args) => log.info(args[0] as string, ...args.slice(1)),
      truncateSentence: (sentence) => this.truncateSentence(sentence),
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
  } {
    const lapis = this.config.isLapis;
    return {
      enabled: lapis?.enabled === true,
      sentenceCardModel: lapis?.sentenceCardModel,
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
      sentenceField: 'Sentence',
      audioField: 'SentenceAudio',
      lapisEnabled: lapis.enabled,
      kikuEnabled: kiku.enabled,
      kikuFieldGrouping: (kiku.fieldGrouping || 'disabled') as 'auto' | 'manual' | 'disabled',
      kikuDeleteDuplicateInAuto: kiku.deleteDuplicateInAuto !== false,
    };
  }

  private isProxyTransportEnabled(config: AnkiConnectConfig = this.config): boolean {
    return config.proxy?.enabled === true;
  }

  private getTransportConfigKey(config: AnkiConnectConfig = this.config): string {
    if (this.isProxyTransportEnabled(config)) {
      return [
        'proxy',
        config.proxy?.host ?? '',
        String(config.proxy?.port ?? ''),
        config.proxy?.upstreamUrl ?? '',
      ].join(':');
    }
    return ['polling', String(config.pollingRate ?? DEFAULT_ANKI_CONNECT_CONFIG.pollingRate)].join(
      ':',
    );
  }

  private startTransport(): void {
    if (this.isProxyTransportEnabled()) {
      const proxyHost = this.config.proxy?.host ?? '127.0.0.1';
      const proxyPort = this.config.proxy?.port ?? 8766;
      const upstreamUrl = this.config.proxy?.upstreamUrl ?? this.config.url ?? '';
      this.getOrCreateProxyServer().start({
        host: proxyHost,
        port: proxyPort,
        upstreamUrl,
      });
      log.info(
        `Starting AnkiConnect integration with local proxy: http://${proxyHost}:${proxyPort} -> ${upstreamUrl}`,
      );
      return;
    }

    log.info('Starting AnkiConnect integration with polling rate:', this.config.pollingRate);
    this.pollingRunner.start();
  }

  private stopTransport(): void {
    this.pollingRunner.stop();
    this.proxyServer?.stop();
  }

  start(): void {
    if (this.started) {
      this.stop();
    }

    this.startKnownWordCacheLifecycle();
    this.startTransport();
    this.started = true;
  }

  stop(): void {
    this.stopTransport();
    this.stopKnownWordCacheLifecycle();
    this.started = false;
    log.info('Stopped AnkiConnect integration');
  }

  private async processNewCard(
    noteId: number,
    options?: { skipKikuFieldGrouping?: boolean },
  ): Promise<void> {
    await this.noteUpdateWorkflow.execute(noteId, options);
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
    const videoFilename = extractFilenameFromMediaPath(currentVideoPath);
    const mediaTitle = trimToNonEmptyString(this.mpvClient.currentMediaTitle);
    const filenameWithExt =
      (shouldPreferMediaTitleForMiscInfo(currentVideoPath, videoFilename)
        ? mediaTitle || videoFilename
        : videoFilename || mediaTitle) || fallbackFilename;
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
      logInfo: (message) => {
        log.info(message);
      },
      logDebug: (message) => {
        log.debug(message);
      },
      logWarn: (message, error) => {
        log.warn(message, (error as Error).message);
      },
    });
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

  private async handleFieldGroupingAuto(
    originalNoteId: number,
    newNoteId: number,
    newNoteInfo: NoteInfo,
    expression: string,
  ): Promise<void> {
    void expression;
    await this.fieldGroupingWorkflow.handleAuto(originalNoteId, newNoteId, newNoteInfo);
  }

  private async handleFieldGroupingManual(
    originalNoteId: number,
    newNoteId: number,
    newNoteInfo: NoteInfo,
    expression: string,
  ): Promise<boolean> {
    void expression;
    return this.fieldGroupingWorkflow.handleManual(originalNoteId, newNoteId, newNoteInfo);
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

  applyRuntimeConfigPatch(patch: Partial<AnkiConnectConfig>): void {
    const wasEnabled = this.config.nPlusOne?.highlightEnabled === true;
    const previousTransportKey = this.getTransportConfigKey(this.config);

    const mergedConfig: AnkiConnectConfig = {
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
      proxy:
        patch.proxy !== undefined ? { ...this.config.proxy, ...patch.proxy } : this.config.proxy,
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
    this.config = this.normalizeConfig(mergedConfig);

    if (wasEnabled && this.config.nPlusOne?.highlightEnabled === false) {
      this.stopKnownWordCacheLifecycle();
      this.knownWordCache.clearKnownWordCacheState();
    } else {
      this.startKnownWordCacheLifecycle();
    }

    const nextTransportKey = this.getTransportConfigKey(this.config);
    if (this.started && previousTransportKey !== nextTransportKey) {
      this.stopTransport();
      this.startTransport();
    }
  }

  destroy(): void {
    this.stop();
    this.mediaGenerator.cleanup();
  }
}
