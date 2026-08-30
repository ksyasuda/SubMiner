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
  type CardKind,
  KikuDuplicateCardInfo,
  KikuFieldGroupingChoice,
  KikuMergePreviewResponse,
  NotificationOptions,
  type WordCardKind,
  type MediaTimingReviewDecision,
  type MediaTimingReviewRequest,
} from './types/anki';
import { AiConfig } from './types/integrations';
import type { KnownWordMaturityTier } from './types/subtitle';
import { MpvClient } from './types/runtime';
import { OPEN_ANKI_CARD_ACTION_ID } from './types/notification';
import type { NotificationType, OverlayNotificationPayload } from './types/notification';
import type { NPlusOneMatchMode, SubtitleMiningContext } from './types/subtitle';
import { DEFAULT_ANKI_CONNECT_CONFIG } from './config';
import {
  getConfiguredWordFieldCandidates,
  getConfiguredWordFieldName,
  getPreferredWordValueFromExtractedFields,
} from './anki-field-config';
import { createLogger } from './logger';
import { captureLiveSubtitleMiningContext } from './core/services/mining';
import {
  createUiFeedbackState,
  beginUpdateProgress,
  clearUpdateProgress,
  endUpdateProgress,
  showStatusNotification,
  showUpdateResult,
  withUpdateProgress,
  UiFeedbackState,
} from './anki-integration/ui-feedback';
import { applyCardKindFlagFields, resolveWordCardKindSetting } from './anki-integration/card-kinds';
import { KnownWordCacheManager } from './anki-integration/known-word-cache';
import { PollingRunner } from './anki-integration/polling';
import type { AnkiConnectProxyServer } from './anki-integration/anki-connect-proxy';
import { findDuplicateNote as findDuplicateNoteForAnkiIntegration } from './anki-integration/duplicate';
import { findDuplicateNoteIds as findDuplicateNoteIdsForAnkiIntegration } from './anki-integration/duplicate';
import { CardCreationService } from './anki-integration/card-creation';
import { FieldGroupingService } from './anki-integration/field-grouping';
import { FieldGroupingMergeCollaborator } from './anki-integration/field-grouping-merge';
import { NoteUpdateWorkflow } from './anki-integration/note-update-workflow';
import { FieldGroupingWorkflow } from './anki-integration/field-grouping-workflow';
import { resolveAnimatedImageLeadInSeconds } from './anki-integration/animated-image-sync';
import { AnkiIntegrationRuntime, normalizeAnkiIntegrationConfig } from './anki-integration/runtime';
import {
  resolveAudioStreamIndexForMediaGeneration,
  resolveMediaGenerationInput,
  resolveMediaGenerationInputPath,
  type MediaGenerationInputResolverOptions,
} from './anki-integration/media-source';
import type { PendingYoutubeMediaUpdate } from './anki-integration/pending-youtube-media';
import { PendingYoutubeMediaQueue } from './anki-integration/pending-youtube-media-queue';
import type {
  PendingYoutubeMediaQueueFailedOptions,
  PendingYoutubeMediaQueueReadyOptions,
} from './anki-integration/pending-youtube-media-queue';
import { resolveMpvVolumeScale } from './anki-integration/mpv-volume';

const log = createLogger('anki').child('integration');

interface NoteInfo {
  noteId: number;
  fields: Record<string, { value: string }>;
}

function trimToNonEmptyString(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function stripRubyReadingText(value: string): string {
  return value
    .replace(/<rt\b[^>]*>[\s\S]*?<\/rt>/gi, '')
    .replace(/<rp\b[^>]*>[\s\S]*?<\/rp>/gi, '');
}

function stripHtmlTags(value: string): string {
  return value.replace(/<[^>]+>/g, '');
}

function getVisibleFuriganaText(value: string): string {
  return stripHtmlTags(stripRubyReadingText(value));
}

function boldMatchingFuriganaTerms(sentenceFurigana: string, highlightedText: string): string {
  if (!sentenceFurigana || !highlightedText || /<b\b/i.test(sentenceFurigana)) {
    return sentenceFurigana;
  }

  const spanRegex = /<span\b[^>]*>[\s\S]*?<\/span>/gi;
  const spans: Array<{ start: number; end: number; visibleStart: number; visibleEnd: number }> = [];
  let visibleSentence = '';
  let match: RegExpExecArray | null;
  while ((match = spanRegex.exec(sentenceFurigana)) !== null) {
    const visibleStart = visibleSentence.length;
    visibleSentence += getVisibleFuriganaText(match[0] || '');
    spans.push({
      start: match.index,
      end: match.index + match[0].length,
      visibleStart,
      visibleEnd: visibleSentence.length,
    });
  }

  if (spans.length === 0) {
    return sentenceFurigana.replace(highlightedText, `<b>${highlightedText}</b>`);
  }

  const highlightStart = visibleSentence.indexOf(highlightedText);
  if (highlightStart === -1) {
    return sentenceFurigana;
  }
  const highlightEnd = highlightStart + highlightedText.length;
  const matchingSpans = spans.filter(
    (span) => span.visibleEnd > highlightStart && span.visibleStart < highlightEnd,
  );
  if (matchingSpans.length === 0) {
    return sentenceFurigana;
  }

  let result = sentenceFurigana;
  for (const span of [...matchingSpans].reverse()) {
    result = `${result.slice(0, span.start)}<b>${result.slice(
      span.start,
      span.end,
    )}</b>${result.slice(span.end)}`;
  }
  return result;
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
  const pathWithoutQuery = separatorIndex >= 0 ? trimmedPath.slice(0, separatorIndex) : trimmedPath;
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

function toOverlayNotificationImageSource(iconBuffer: Buffer): string {
  return `data:image/png;base64,${iconBuffer.toString('base64')}`;
}

interface NotificationIcon {
  filePath?: string;
  overlayImageSource: string;
}

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
  private overlayNotificationCallback: ((payload: OverlayNotificationPayload) => void) | null =
    null;
  private overlayNotificationDismissCallback: ((id: string) => void) | null = null;
  private overlayUpdateProgressActive = false;
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
  private runtime: AnkiIntegrationRuntime;
  private aiConfig: AiConfig;
  private recordCardsMinedCallback: ((count: number, noteIds?: number[]) => void) | null = null;
  private knownWordCacheUpdatedCallback: (() => void) | null = null;
  private consumeSubtitleMiningContextCallback: (() => SubtitleMiningContext | null) | null = null;
  private mediaTimingReviewCallback:
    | ((request: MediaTimingReviewRequest) => Promise<MediaTimingReviewDecision>)
    | null = null;
  private noteIdRedirects = new Map<number, number>();
  private trackedDuplicateNoteIds = new Map<number, number[]>();
  private getCachedMediaPath: MediaGenerationInputResolverOptions['getCachedMediaPath'] | null =
    null;
  private shouldRequireRemoteMediaCache: (() => boolean) | null = null;
  private getYoutubeMediaSourceUrl:
    | (() => Promise<string | null | undefined> | string | null | undefined)
    | null = null;
  private pendingYoutubeMediaQueue: PendingYoutubeMediaQueue;

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
    aiConfig: AiConfig = {},
    recordCardsMined?: (count: number, noteIds?: number[]) => void,
    overlayNotificationCallback?: (payload: OverlayNotificationPayload) => void,
    getCachedMediaPath?: MediaGenerationInputResolverOptions['getCachedMediaPath'],
    shouldRequireRemoteMediaCache?: () => boolean,
    getYoutubeMediaSourceUrl?: () => Promise<string | null | undefined> | string | null | undefined,
    overlayNotificationDismissCallback?: (id: string) => void,
  ) {
    this.config = normalizeAnkiIntegrationConfig(config);
    this.aiConfig = { ...aiConfig };
    this.client = new AnkiConnectClient(this.config.url!);
    this.mediaGenerator = new MediaGenerator();
    this.timingTracker = timingTracker;
    this.mpvClient = mpvClient;
    this.osdCallback = osdCallback || null;
    this.notificationCallback = notificationCallback || null;
    this.overlayNotificationCallback = overlayNotificationCallback || null;
    this.fieldGroupingCallback = fieldGroupingCallback || null;
    this.recordCardsMinedCallback = recordCardsMined ?? null;
    this.getCachedMediaPath = getCachedMediaPath ?? null;
    this.shouldRequireRemoteMediaCache = shouldRequireRemoteMediaCache ?? null;
    this.getYoutubeMediaSourceUrl = getYoutubeMediaSourceUrl ?? null;
    this.overlayNotificationDismissCallback = overlayNotificationDismissCallback ?? null;
    this.pendingYoutubeMediaQueue = this.createPendingYoutubeMediaQueue();
    this.knownWordCache = this.createKnownWordCache(knownWordCacheStatePath);
    this.pollingRunner = this.createPollingRunner();
    this.cardCreationService = this.createCardCreationService();
    this.fieldGroupingMergeCollaborator = this.createFieldGroupingMergeCollaborator();
    this.fieldGroupingService = this.createFieldGroupingService();
    this.noteUpdateWorkflow = this.createNoteUpdateWorkflow();
    this.fieldGroupingWorkflow = this.createFieldGroupingWorkflow();
    this.runtime = this.createRuntime(config);
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
      generateMediaForMerge: (noteInfo) => this.generateMediaForMerge(noteInfo),
      warnFieldParseOnce: (fieldName, reason, detail) =>
        this.warnFieldParseOnce(fieldName, reason, detail),
    });
  }

  private recordCardsMinedSafely(
    count: number,
    noteIds: number[] | undefined,
    source: string,
  ): void {
    if (!this.recordCardsMinedCallback) {
      return;
    }

    try {
      this.recordCardsMinedCallback(count, noteIds);
    } catch (error) {
      log.warn(`recordCardsMined callback failed during ${source}:`, (error as Error).message);
    }
  }

  private getMediaResolverOptions(): MediaGenerationInputResolverOptions {
    const options: MediaGenerationInputResolverOptions = {
      logDebug: (message) => log.debug(message),
    };
    if (this.getCachedMediaPath) {
      options.getCachedMediaPath = this.getCachedMediaPath;
    }
    if (this.shouldRequireRemoteMediaCache?.()) {
      options.remoteCacheMode = 'required';
    }
    return options;
  }

  private createPendingYoutubeMediaQueue(): PendingYoutubeMediaQueue {
    return new PendingYoutubeMediaQueue({
      client: {
        notesInfo: async (noteIds) => (await this.client.notesInfo(noteIds)) as unknown,
        updateNoteFields: (noteId, fields) => this.client.updateNoteFields(noteId, fields),
        storeMediaFile: (filename, data) => this.client.storeMediaFile(filename, data),
      },
      mediaGenerator: {
        generateAudio: (
          videoPath,
          startTime,
          endTime,
          audioPadding,
          audioStreamIndex,
          normalizeAudio,
          volumeScale,
        ) =>
          this.mediaGenerator.generateAudio(
            videoPath,
            startTime,
            endTime,
            audioPadding,
            audioStreamIndex,
            normalizeAudio,
            volumeScale,
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
      getConfig: () => this.config,
      getCurrentVideoPath: () => this.getCurrentYoutubeMediaSourceUrl(),
      getCachedMediaPath: this.getCachedMediaPath,
      shouldRequireRemoteMediaCache: () => this.shouldRequireRemoteMediaCache?.() === true,
      getSubtitleMediaRange: (context) => this.getSubtitleMediaRange(context),
      resolveConfiguredFieldName: (noteInfo, ...preferredNames) =>
        this.resolveConfiguredFieldName(noteInfo, ...preferredNames),
      mergeFieldValue: (existing, newValue, overwrite) =>
        this.mergeFieldValue(existing, newValue, overwrite),
      getAnimatedImageLeadInSeconds: (noteInfo) => this.getAnimatedImageLeadInSeconds(noteInfo),
      getMpvVolumeScale: () => this.getMpvVolumeScale(),
      generateAudioFilename: () => this.generateAudioFilename(),
      generateImageFilename: () => this.generateImageFilename(),
      formatMiscInfoPatternForMediaPath: (
        fallbackFilename,
        startTimeSeconds,
        mediaPath,
        mediaTitle,
      ) =>
        this.formatMiscInfoPatternForMediaPath(
          fallbackFilename,
          startTimeSeconds,
          mediaPath,
          mediaTitle,
        ),
      showStatusNotification: (message) => this.showStatusNotification(message),
      showNotification: (noteId, label, errorSuffix) =>
        this.showNotification(noteId, label, errorSuffix),
      logInfo: (...args) => log.info(args[0] as string, ...args.slice(1)),
      logWarn: (...args) => log.warn(args[0] as string, ...args.slice(1)),
      logError: (...args) => log.error(args[0] as string, ...args.slice(1)),
    });
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
      recordCardsAdded: (count, noteIds) => {
        this.recordCardsMinedSafely(count, noteIds, 'polling');
      },
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
      recordCardsAdded: (count, noteIds) => {
        this.recordCardsMinedSafely(count, noteIds, 'proxy');
      },
      trackAddedDuplicateNoteIds: (noteId, duplicateNoteIds) => {
        this.trackDuplicateNoteIdsForNote(noteId, duplicateNoteIds);
      },
      getDeck: () => this.config.deck,
      findNotes: async (query, options) =>
        (await this.client.findNotes(query, options)) as number[],
      notifyUnavailable: (message) => this.showStatusNotification(message),
      logInfo: (message, ...args) => log.info(message, ...args),
      logWarn: (message, ...args) => log.warn(message, ...args),
      logError: (message, ...args) => log.error(message, ...args),
    });
  }

  private createRuntime(initialConfig: AnkiConnectConfig): AnkiIntegrationRuntime {
    return new AnkiIntegrationRuntime({
      initialConfig,
      pollingRunner: this.pollingRunner,
      knownWordCache: this.knownWordCache,
      proxyServerFactory: () => this.createProxyServer(),
      logInfo: (message, ...args) => log.info(message, ...args),
      logWarn: (message, ...args) => log.warn(message, ...args),
      logError: (message, ...args) => log.error(message, ...args),
      onConfigChanged: (nextConfig) => {
        this.config = nextConfig;
        this.client = new AnkiConnectClient(nextConfig.url!);
      },
    });
  }

  private createCardCreationService(): CardCreationService {
    return new CardCreationService({
      getConfig: () => this.config,
      getAiConfig: () => this.aiConfig,
      getTimingTracker: () => this.timingTracker,
      getMpvClient: () => this.mpvClient,
      ...(this.getCachedMediaPath ? { getCachedMediaPath: this.getCachedMediaPath } : {}),
      shouldRequireRemoteMediaCache: () => this.shouldRequireRemoteMediaCache?.() === true,
      getYoutubeMediaSourceUrl: () => this.getCurrentYoutubeMediaSourceUrl(),
      queuePendingYoutubeMediaUpdate: (job) => this.queuePendingYoutubeMediaUpdate(job),
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
        retrieveMediaFile: (filename) => this.client.retrieveMediaFile(filename),
        deleteNotes: (noteIds) => this.client.deleteNotes(noteIds),
      },
      mediaGenerator: {
        generateAudio: (
          videoPath,
          startTime,
          endTime,
          audioPadding,
          audioStreamIndex,
          normalizeAudio,
          volumeScale,
        ) =>
          this.mediaGenerator.generateAudio(
            videoPath,
            startTime,
            endTime,
            audioPadding,
            audioStreamIndex,
            normalizeAudio,
            volumeScale,
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
      showOsdNotification: (text: string) => this.showStatusNotification(text),
      showUpdateResult: (message: string, success: boolean) =>
        this.showUpdateResult(message, success),
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
      getAnimatedImageLeadInSeconds: (noteInfo) => this.getAnimatedImageLeadInSeconds(noteInfo),
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
      removeKnownWordNote: (noteId) => this.removeKnownWordNote(noteId),
      isUpdateInProgress: () => this.updateInProgress,
      setUpdateInProgress: (value) => {
        this.updateInProgress = value;
      },
      trackLastAddedNoteId: (noteId) => {
        this.previousNoteIds.add(noteId);
      },
      trackLastAddedDuplicateNoteIds: (noteId, duplicateNoteIds) => {
        this.trackedDuplicateNoteIds.set(noteId, [...duplicateNoteIds]);
      },
      findDuplicateNoteIds: (expression, noteInfo) =>
        this.findDuplicateNoteIds(expression, noteInfo),
      recordCardsMinedCallback: (count, noteIds) => {
        this.recordCardsMinedSafely(count, noteIds, 'card creation');
      },
      reviewMediaTiming: (request) => this.reviewMediaTiming(request),
    });
  }

  private createFieldGroupingService(): FieldGroupingService {
    return new FieldGroupingService({
      getEffectiveSentenceCardConfig: () => this.getEffectiveSentenceCardConfig(),
      getConfig: () => this.config,
      isUpdateInProgress: () => this.updateInProgress,
      getDeck: () => this.config.deck,
      withUpdateProgress: <T>(initialMessage: string, action: () => Promise<T>) =>
        this.withUpdateProgress(initialMessage, action),
      showOsdNotification: (text: string) => this.showStatusNotification(text),
      findNotes: async (query, options) =>
        (await this.client.findNotes(query, options)) as number[],
      notesInfo: async (noteIds) => (await this.client.notesInfo(noteIds)) as unknown as NoteInfo[],
      extractFields: (fields) => this.extractFields(fields),
      findDuplicateNote: (expression, noteId, noteInfo) =>
        this.findDuplicateNote(expression, noteId, noteInfo),
      getTrackedDuplicateNoteIds: (noteId) =>
        this.trackedDuplicateNoteIds.has(noteId)
          ? [...(this.trackedDuplicateNoteIds.get(noteId) ?? [])]
          : null,
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
        deleteNotes: (noteIds) => this.client.deleteNotes(noteIds),
      },
      getConfig: () => this.config,
      getCurrentSubtitleText: () => this.mpvClient.currentSubText,
      getCurrentSubtitleStart: () => this.mpvClient.currentSubStart,
      getEffectiveSentenceCardConfig: () => this.getEffectiveSentenceCardConfig(),
      appendKnownWordsFromNoteInfo: (noteInfo) => this.appendKnownWordsFromNoteInfo(noteInfo),
      removeKnownWordNote: (noteId) => this.removeKnownWordNote(noteId),
      extractFields: (fields) => this.extractFields(fields),
      findDuplicateNote: (expression, excludeNoteId, noteInfo) =>
        this.findDuplicateNote(expression, excludeNoteId, noteInfo),
      handleFieldGroupingAuto: (originalNoteId, newNoteId, newNoteInfo, expression) =>
        this.handleFieldGroupingAuto(originalNoteId, newNoteId, newNoteInfo, expression),
      handleFieldGroupingManual: (originalNoteId, newNoteId, newNoteInfo, expression) =>
        this.handleFieldGroupingManual(originalNoteId, newNoteId, newNoteInfo, expression),
      processSentence: (mpvSentence, noteFields) => this.processSentence(mpvSentence, noteFields),
      processSentenceFurigana: (sentenceFurigana, noteFields) =>
        this.processSentenceFurigana(sentenceFurigana, noteFields),
      setCardTypeFields: (updatedFields, availableFieldNames, cardKind) =>
        this.setCardTypeFields(updatedFields, availableFieldNames, cardKind),
      resolveConfiguredFieldName: (noteInfo, ...preferredNames) =>
        this.resolveConfiguredFieldName(noteInfo, ...preferredNames),
      getAnimatedImageLeadInSeconds: (noteInfo) => this.getAnimatedImageLeadInSeconds(noteInfo),
      mergeFieldValue: (existing, newValue, overwrite) =>
        this.mergeFieldValue(existing, newValue, overwrite),
      generateAudioFilename: () => this.generateAudioFilename(),
      generateAudio: (context) => this.generateAudio(context),
      generateImageFilename: () => this.generateImageFilename(),
      generateImage: (animatedLeadInSeconds, context) =>
        this.generateImage(animatedLeadInSeconds, context),
      formatMiscInfoPattern: (fallbackFilename, startTimeSeconds) =>
        this.formatMiscInfoPattern(fallbackFilename, startTimeSeconds),
      consumeSubtitleMiningContext: () => this.consumeSubtitleMiningContext(),
      captureSubtitleMediaContext: () => captureLiveSubtitleMiningContext(this.mpvClient),
      queuePendingYoutubeMediaUpdate: (job) => this.queuePendingYoutubeMediaUpdateForNote(job),
      addConfiguredTagsToNote: (noteId) => this.addConfiguredTagsToNote(noteId),
      showNotification: (noteId, label) => this.showNotification(noteId, label),
      showOsdNotification: (message) => this.showStatusNotification(message),
      beginUpdateProgress: (initialMessage) => this.beginUpdateProgress(initialMessage),
      endUpdateProgress: () => this.endUpdateProgress(),
      logWarn: (...args) => log.warn(args[0] as string, ...args.slice(1)),
      logInfo: (...args) => log.info(args[0] as string, ...args.slice(1)),
      logError: (...args) => log.error(args[0] as string, ...args.slice(1)),
      reviewMediaTiming: (request) => this.reviewMediaTiming(request),
    });
  }

  private createFieldGroupingWorkflow(): FieldGroupingWorkflow {
    return new FieldGroupingWorkflow({
      client: {
        notesInfo: async (noteIds) => (await this.client.notesInfo(noteIds)) as unknown,
        updateNoteFields: (noteId, fields) => this.client.updateNoteFields(noteId, fields),
        addTags: (noteIds, tags) => this.client.addTags(noteIds, tags),
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
      rememberMergedNoteIds: (deletedNoteId, keptNoteId) => {
        this.rememberMergedNoteIds(deletedNoteId, keptNoteId);
      },
      showStatusNotification: (message) => this.showStatusNotification(message),
      showNotification: (noteId, label) => this.showNotification(noteId, label),
      showOsdNotification: (message) => this.showStatusNotification(message),
      logError: (...args) => log.error(args[0] as string, ...args.slice(1)),
      logInfo: (...args) => log.info(args[0] as string, ...args.slice(1)),
      truncateSentence: (sentence) => this.truncateSentence(sentence),
    });
  }

  isKnownWord(
    text: string,
    reading?: string,
    options?: { allowReadingOnlyMatch?: boolean },
  ): boolean {
    return this.knownWordCache.isKnownWord(text, reading, options);
  }

  getKnownWordTier(
    text: string,
    reading?: string,
    options?: { allowReadingOnlyMatch?: boolean },
  ): KnownWordMaturityTier | null {
    return this.knownWordCache.getKnownWordTier(text, reading, options);
  }

  getKnownWordMatchMode(): NPlusOneMatchMode {
    return this.config.knownWords?.matchMode ?? DEFAULT_ANKI_CONNECT_CONFIG.knownWords.matchMode;
  }

  async openNoteInAnki(noteId: number): Promise<void> {
    await this.client.openNoteInBrowser(noteId);
  }

  private isKnownWordCacheEnabled(): boolean {
    return (
      this.config.knownWords?.highlightEnabled === true || this.config.nPlusOne?.enabled === true
    );
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
    const shouldNotify = this.isKnownWordCacheEnabled();
    await this.knownWordCache.refresh(true);
    if (shouldNotify) {
      this.notifyKnownWordCacheUpdated();
    }
  }

  private appendKnownWordsFromNoteInfo(noteInfo: NoteInfo): void {
    if (!this.isKnownWordCacheEnabled()) {
      return;
    }

    const changed = this.knownWordCache.appendFromNoteInfo({
      noteId: noteInfo.noteId,
      fields: noteInfo.fields,
    });
    if (changed) {
      this.notifyKnownWordCacheUpdated();
    }
  }

  private removeKnownWordNote(noteId: number): void {
    if (this.knownWordCache.removeNote(noteId)) {
      this.notifyKnownWordCacheUpdated();
    }
  }

  private notifyKnownWordCacheUpdated(): void {
    if (!this.knownWordCacheUpdatedCallback) {
      return;
    }

    try {
      this.knownWordCacheUpdatedCallback();
    } catch (error) {
      log.warn('Known-word cache update callback failed:', (error as Error).message);
    }
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
    wordCardKind: WordCardKind;
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
      wordCardKind: resolveWordCardKindSetting(this.config.lapisKiku?.wordCardKind),
    };
  }

  start(): void {
    this.runtime.start();
  }

  waitUntilReady(): Promise<void> {
    return this.runtime.waitUntilReady();
  }

  stop(): void {
    this.runtime.stop();
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

  private getSentenceHighlightText(noteFields: Record<string, string>): string {
    const sentenceFieldName = this.config.fields?.sentence?.toLowerCase() || 'sentence';
    const existingSentence = noteFields[sentenceFieldName] || '';
    return (
      existingSentence.match(/<b>(.*?)<\/b>/)?.[1] ||
      getPreferredWordValueFromExtractedFields(noteFields, this.config).trim()
    );
  }

  private processSentence(mpvSentence: string, noteFields: Record<string, string>): string {
    if (this.config.behavior?.highlightWord === false) {
      return mpvSentence;
    }

    const highlightedText = this.getSentenceHighlightText(noteFields);
    if (!highlightedText) {
      return mpvSentence;
    }

    const index = mpvSentence.indexOf(highlightedText);

    if (index === -1) {
      return mpvSentence;
    }

    const prefix = mpvSentence.substring(0, index);
    const suffix = mpvSentence.substring(index + highlightedText.length);
    return `${prefix}<b>${highlightedText}</b>${suffix}`;
  }

  private processSentenceFurigana(
    sentenceFurigana: string,
    noteFields: Record<string, string>,
  ): string {
    if (this.config.behavior?.highlightWord === false) {
      return sentenceFurigana;
    }

    const highlightedText = this.getSentenceHighlightText(noteFields);
    return highlightedText
      ? boldMatchingFuriganaTerms(sentenceFurigana, highlightedText)
      : sentenceFurigana;
  }

  private consumeSubtitleMiningContext(): SubtitleMiningContext | null {
    if (!this.consumeSubtitleMiningContextCallback) {
      return null;
    }

    try {
      return this.consumeSubtitleMiningContextCallback();
    } catch (error) {
      log.warn('Subtitle mining context callback failed:', (error as Error).message);
      return null;
    }
  }

  private getSubtitleMediaRange(context?: SubtitleMiningContext): {
    startTime: number;
    endTime: number;
  } {
    if (
      context &&
      Number.isFinite(context.startTime) &&
      Number.isFinite(context.endTime) &&
      context.endTime > context.startTime
    ) {
      return {
        startTime: context.startTime,
        endTime: context.endTime,
      };
    }

    if (
      Number.isFinite(this.mpvClient.currentSubStart) &&
      Number.isFinite(this.mpvClient.currentSubEnd) &&
      this.mpvClient.currentSubEnd > this.mpvClient.currentSubStart
    ) {
      return {
        startTime: this.mpvClient.currentSubStart,
        endTime: this.mpvClient.currentSubEnd,
      };
    }

    const currentTime = this.mpvClient.currentTimePos || 0;
    const fallback = this.getFallbackDurationSeconds() / 2;
    return {
      startTime: currentTime - fallback,
      endTime: currentTime + fallback,
    };
  }

  private queuePendingYoutubeMediaUpdate(job: PendingYoutubeMediaUpdate): void {
    this.pendingYoutubeMediaQueue.enqueue(job);
  }

  private async queuePendingYoutubeMediaUpdateForNote(job: {
    noteId: number;
    noteInfo: NoteInfo;
    context?: SubtitleMiningContext;
    label: string | number;
  }): Promise<boolean> {
    return this.pendingYoutubeMediaQueue.queueFromNote(job);
  }

  private async getCurrentYoutubeMediaSourceUrl(): Promise<string> {
    return (
      trimToNonEmptyString(await this.getYoutubeMediaSourceUrl?.()) ??
      trimToNonEmptyString(this.mpvClient.currentVideoPath) ??
      ''
    );
  }

  private async getMpvVolumeScale(): Promise<number | undefined> {
    return resolveMpvVolumeScale(this.mpvClient, this.config.media?.mirrorMpvVolume !== false);
  }

  async handleYoutubeMediaCacheReady(
    sourceUrl: string,
    cachedPath: string,
    options?: PendingYoutubeMediaQueueReadyOptions,
  ): Promise<void> {
    await this.pendingYoutubeMediaQueue.handleReady(sourceUrl, cachedPath, options);
  }

  async handleYoutubeMediaCacheFailed(
    sourceUrl: string,
    options?: PendingYoutubeMediaQueueFailedOptions,
  ): Promise<void> {
    await this.pendingYoutubeMediaQueue.handleFailed(sourceUrl, options);
  }

  private async generateAudio(context?: SubtitleMiningContext): Promise<Buffer | null> {
    const mpvClient = this.mpvClient;
    if (!mpvClient || !mpvClient.currentVideoPath) {
      return null;
    }

    const videoPath = await resolveMediaGenerationInput(
      mpvClient,
      'audio',
      this.getMediaResolverOptions(),
    );
    if (!videoPath) {
      return null;
    }
    const { startTime, endTime } = this.getSubtitleMediaRange(context);

    return this.mediaGenerator.generateAudio(
      videoPath,
      startTime,
      endTime,
      context?.mediaPaddingSeconds ?? this.config.media?.audioPadding,
      resolveAudioStreamIndexForMediaGeneration(videoPath, this.mpvClient.currentAudioStreamIndex),
      this.config.media?.normalizeAudio !== false,
      await this.getMpvVolumeScale(),
    );
  }

  private async generateImage(
    animatedLeadInSeconds = 0,
    context?: SubtitleMiningContext,
  ): Promise<Buffer | null> {
    if (!this.mpvClient || !this.mpvClient.currentVideoPath) {
      return null;
    }

    const videoPath = await resolveMediaGenerationInput(
      this.mpvClient,
      'video',
      this.getMediaResolverOptions(),
    );
    if (!videoPath) {
      return null;
    }
    const mediaRange = this.getSubtitleMediaRange(context);
    const timestamp = context
      ? mediaRange.startTime + (mediaRange.endTime - mediaRange.startTime) / 2
      : this.mpvClient.currentTimePos || 0;

    if (this.config.media?.imageType === 'avif') {
      return this.mediaGenerator.generateAnimatedImage(
        videoPath,
        mediaRange.startTime,
        mediaRange.endTime,
        context?.mediaPaddingSeconds ?? this.config.media?.audioPadding,
        {
          fps: this.config.media?.animatedFps,
          maxWidth: this.config.media?.animatedMaxWidth,
          maxHeight: this.config.media?.animatedMaxHeight,
          crf: this.config.media?.animatedCrf,
          leadingStillDuration: animatedLeadInSeconds,
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
    return this.formatMiscInfoPatternForMediaPath(
      fallbackFilename,
      startTimeSeconds,
      this.mpvClient.currentVideoPath || '',
      this.mpvClient.currentMediaTitle ?? undefined,
    );
  }

  private formatMiscInfoPatternForMediaPath(
    fallbackFilename: string,
    startTimeSeconds: number | undefined,
    mediaPath: string,
    mediaTitle?: string,
  ): string {
    if (!this.config.metadata?.pattern) {
      return '';
    }

    const videoFilename = extractFilenameFromMediaPath(mediaPath);
    const resolvedMediaTitle = trimToNonEmptyString(mediaTitle);
    const filenameWithExt =
      (shouldPreferMediaTitleForMiscInfo(mediaPath, videoFilename)
        ? resolvedMediaTitle || videoFilename
        : videoFilename || resolvedMediaTitle) || fallbackFilename;
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
      getNotificationType: () => this.getNotificationType(),
      showOsd: (text: string) => {
        this.showOsdNotification(text);
      },
      showOverlayNotification: (payload) => {
        this.overlayNotificationCallback?.(payload);
      },
      showSystemNotification: (title: string, options: NotificationOptions) => {
        if (this.notificationCallback) {
          this.notificationCallback(title, options);
        }
      },
    });
  }

  private getNotificationType(): NotificationType {
    return this.config.behavior?.notificationType ?? 'osd';
  }

  private shouldUseOsdNotifications(): boolean {
    const type = this.getNotificationType();
    return type === 'osd' || type === 'osd-system';
  }

  private shouldUseOverlayNotifications(): boolean {
    const type = this.getNotificationType();
    return type === 'overlay' || type === 'both';
  }

  private beginUpdateProgress(initialMessage: string): void {
    if (!this.shouldUseOsdNotifications()) {
      if (this.shouldUseOverlayNotifications()) {
        this.overlayUpdateProgressActive = true;
        this.overlayNotificationCallback?.({
          id: 'anki-update-progress',
          title: 'Anki update',
          body: initialMessage,
          variant: 'progress',
          persistent: true,
        });
      }
      return;
    }
    beginUpdateProgress(this.uiFeedbackState, initialMessage, (text: string) => {
      this.showOsdNotification(text);
    });
  }

  private endUpdateProgress(): void {
    if (this.overlayUpdateProgressActive) {
      this.overlayUpdateProgressActive = false;
      this.overlayNotificationDismissCallback?.('anki-update-progress');
    }
    if (!this.shouldUseOsdNotifications()) {
      return;
    }
    endUpdateProgress(this.uiFeedbackState, (timer) => {
      clearInterval(timer);
    });
  }

  private clearUpdateProgress(): void {
    if (!this.shouldUseOsdNotifications()) {
      return;
    }
    clearUpdateProgress(this.uiFeedbackState, (timer) => {
      clearInterval(timer);
    });
  }

  private async withUpdateProgress<T>(
    initialMessage: string,
    action: () => Promise<T>,
  ): Promise<T> {
    if (!this.shouldUseOsdNotifications()) {
      this.updateInProgress = true;
      if (this.shouldUseOverlayNotifications()) {
        this.overlayUpdateProgressActive = true;
        this.overlayNotificationCallback?.({
          id: 'anki-update-progress',
          title: 'Anki update',
          body: initialMessage,
          variant: 'progress',
          persistent: true,
        });
      }
      try {
        return await action();
      } finally {
        this.updateInProgress = false;
        this.endUpdateProgress();
      }
    }
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
    applyCardKindFlagFields(updatedFields, cardKind, (preferredName) =>
      this.resolveFieldName(availableFieldNames, preferredName),
    );
  }

  private async showNotification(
    noteId: number,
    label: string | number,
    errorSuffix?: string,
  ): Promise<void> {
    const message = errorSuffix
      ? `Updated card: ${label} (${errorSuffix})`
      : `Updated card: ${label}`;

    const type = this.getNotificationType();

    if (type === 'osd' || type === 'osd-system') {
      this.showUpdateResult(message, errorSuffix === undefined);
    } else {
      this.clearUpdateProgress();
    }

    const shouldShowOverlayNotification =
      (type === 'overlay' || type === 'both') && this.overlayNotificationCallback !== null;
    const shouldShowSystemNotification =
      (type === 'system' || type === 'both' || type === 'osd-system') &&
      this.notificationCallback !== null;
    const notificationIcon =
      shouldShowOverlayNotification || shouldShowSystemNotification
        ? await this.generateNotificationIcon(noteId, shouldShowSystemNotification)
        : undefined;

    if (shouldShowOverlayNotification && this.overlayNotificationCallback) {
      this.overlayUpdateProgressActive = false;
      this.overlayNotificationCallback({
        id: 'anki-update-progress',
        title: 'Anki Card Updated',
        body: message,
        ...(notificationIcon ? { image: notificationIcon.overlayImageSource } : {}),
        variant: errorSuffix === undefined ? 'success' : 'error',
        persistent: false,
        actions: [{ id: OPEN_ANKI_CARD_ACTION_ID, label: 'Open in Anki', noteId }],
      });
    }

    if (shouldShowSystemNotification && this.notificationCallback) {
      this.notificationCallback('Anki Card Updated', {
        body: message,
        icon: notificationIcon?.filePath,
      });
    }

    if (notificationIcon) {
      if (notificationIcon.filePath) {
        this.mediaGenerator.scheduleNotificationIconCleanup(notificationIcon.filePath);
      }
    }
  }

  private async generateNotificationIcon(
    noteId: number,
    shouldWriteToFile: boolean,
  ): Promise<NotificationIcon | undefined> {
    if (!this.mpvClient?.currentVideoPath) {
      return undefined;
    }

    try {
      const timestamp = this.mpvClient.currentTimePos || 0;
      const notificationIconSource = await resolveMediaGenerationInputPath(this.mpvClient, 'video');
      if (!notificationIconSource) {
        throw new Error('No media source available for notification icon');
      }
      const iconBuffer = await this.mediaGenerator.generateNotificationIcon(
        notificationIconSource,
        timestamp,
      );
      if (iconBuffer && iconBuffer.length > 0) {
        const notificationIcon: NotificationIcon = {
          overlayImageSource: toOverlayNotificationImageSource(iconBuffer),
        };
        if (shouldWriteToFile) {
          try {
            notificationIcon.filePath = this.mediaGenerator.writeNotificationIconToFile(
              iconBuffer,
              noteId,
            );
          } catch (err) {
            log.warn('Failed to write notification icon:', (err as Error).message);
          }
        }
        return notificationIcon;
      }
    } catch (err) {
      log.warn('Failed to generate notification icon:', (err as Error).message);
    }

    return undefined;
  }

  private showUpdateResult(message: string, success: boolean): void {
    showUpdateResult(
      this.uiFeedbackState,
      {
        clearProgressTimer: (timer) => {
          clearInterval(timer);
        },
        showOsdNotification: (text) => {
          this.showOsdNotification(text);
        },
      },
      { message, success },
    );
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
    const trackedDuplicateNoteIdsBeforeCreate = new Set(this.trackedDuplicateNoteIds.keys());
    const created = await this.cardCreationService.createSentenceCard(
      sentence,
      startTime,
      endTime,
      secondarySubText,
    );
    if (
      created &&
      this.shouldTriggerFieldGroupingAfterLocalSentenceCardCreate(
        trackedDuplicateNoteIdsBeforeCreate,
      )
    ) {
      try {
        await this.fieldGroupingService.triggerFieldGroupingForLastAddedCard();
      } catch (error) {
        log.warn('Failed to trigger field grouping after sentence card creation:', error);
      }
    }
    return created;
  }

  trackDuplicateNoteIdsForNote(noteId: number, duplicateNoteIds: number[]): void {
    this.trackedDuplicateNoteIds.set(noteId, [...duplicateNoteIds]);
  }

  private shouldTriggerFieldGroupingAfterLocalSentenceCardCreate(
    trackedDuplicateNoteIdsBeforeCreate: Set<number>,
  ): boolean {
    const sentenceCardConfig = this.getEffectiveSentenceCardConfig();
    if (!sentenceCardConfig.kikuEnabled || sentenceCardConfig.kikuFieldGrouping === 'disabled') {
      return false;
    }

    for (const noteId of this.trackedDuplicateNoteIds.keys()) {
      if (!trackedDuplicateNoteIdsBeforeCreate.has(noteId)) {
        return true;
      }
    }
    return false;
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
      getWordFieldCandidates: () => this.getConfiguredWordFieldCandidates(),
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

  private async findDuplicateNoteIds(expression: string, noteInfo: NoteInfo): Promise<number[]> {
    return findDuplicateNoteIdsForAnkiIntegration(expression, -1, noteInfo, {
      findNotes: async (query, options) => (await this.client.findNotes(query, options)) as unknown,
      notesInfo: async (noteIds) => (await this.client.notesInfo(noteIds)) as unknown,
      getDeck: () => this.config.deck,
      getWordFieldCandidates: () => this.getConfiguredWordFieldCandidates(),
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

  private getConfiguredWordFieldName(): string {
    return getConfiguredWordFieldName(this.config);
  }

  private getConfiguredWordFieldCandidates(): string[] {
    return getConfiguredWordFieldCandidates(this.config);
  }

  private async getAnimatedImageLeadInSeconds(noteInfo: NoteInfo): Promise<number> {
    return resolveAnimatedImageLeadInSeconds({
      config: this.config,
      noteInfo,
      resolveConfiguredFieldName: (candidateNoteInfo, ...preferredNames) =>
        this.resolveConfiguredFieldName(candidateNoteInfo, ...preferredNames),
      retrieveMediaFileBase64: (filename) => this.client.retrieveMediaFile(filename),
      logWarn: (message, ...args) => log.warn(message, ...args),
    });
  }

  private async generateMediaForMerge(noteInfo?: NoteInfo): Promise<{
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

    if (this.config.media?.generateAudio !== false && this.mpvClient?.currentVideoPath) {
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

    if (this.config.media?.generateImage !== false && this.mpvClient?.currentVideoPath) {
      try {
        const animatedLeadInSeconds = noteInfo
          ? await this.getAnimatedImageLeadInSeconds(noteInfo)
          : 0;
        const imageFilename = this.generateImageFilename();
        const imageBuffer = await this.generateImage(animatedLeadInSeconds);
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

  applyRuntimeConfigPatch(patch: Partial<AnkiConnectConfig>, aiConfig?: AiConfig): void {
    if (aiConfig) {
      this.aiConfig = { ...aiConfig };
    }
    this.runtime.applyRuntimeConfigPatch(patch);
  }

  destroy(): void {
    this.stop();
    this.mediaGenerator.cleanup();
  }

  setRecordCardsMinedCallback(
    callback: ((count: number, noteIds?: number[]) => void) | null,
  ): void {
    this.recordCardsMinedCallback = callback;
  }

  setKnownWordCacheUpdatedCallback(callback: (() => void) | null): void {
    this.knownWordCacheUpdatedCallback = callback;
  }

  setSubtitleMiningContextConsumer(callback: (() => SubtitleMiningContext | null) | null): void {
    this.consumeSubtitleMiningContextCallback = callback;
  }

  setMediaTimingReviewCallback(
    callback: ((request: MediaTimingReviewRequest) => Promise<MediaTimingReviewDecision>) | null,
  ): void {
    this.mediaTimingReviewCallback = callback;
  }

  private async reviewMediaTiming(
    request: Omit<MediaTimingReviewRequest, 'audioPadding' | 'maxMediaDuration'>,
  ): Promise<MediaTimingReviewDecision> {
    if (this.config.media?.reviewTiming !== true || !this.mediaTimingReviewCallback) {
      return { action: 'use-original' };
    }
    return await this.mediaTimingReviewCallback({
      ...request,
      audioPadding: Math.max(0, this.config.media.audioPadding ?? 0),
      maxMediaDuration: Math.max(0, this.config.media.maxMediaDuration ?? 30),
    });
  }

  resolveCurrentNoteId(noteId: number): number {
    let resolved = noteId;
    const seen = new Set<number>();
    while (this.noteIdRedirects.has(resolved) && !seen.has(resolved)) {
      seen.add(resolved);
      resolved = this.noteIdRedirects.get(resolved)!;
    }
    return resolved;
  }

  private rememberMergedNoteIds(deletedNoteId: number, keptNoteId: number): void {
    const resolvedKeepNoteId = this.resolveCurrentNoteId(keptNoteId);
    const visited = new Set<number>([deletedNoteId]);
    let current = deletedNoteId;

    while (true) {
      this.noteIdRedirects.set(current, resolvedKeepNoteId);
      const next = Array.from(this.noteIdRedirects.entries()).find(
        ([, targetNoteId]) => targetNoteId === current,
      )?.[0];
      if (next === undefined || visited.has(next)) {
        break;
      }
      visited.add(next);
      current = next;
    }
  }
}
