import { DEFAULT_ANKI_CONNECT_CONFIG } from '../config';
import type { MediaInput } from '../media-input';
import type { MediaGenerator } from '../media-generator';
import type { AnkiConnectConfig } from '../types/anki';
import type { SubtitleMiningContext } from '../types/subtitle';
import { youtubeMediaUrlsMatch, type PendingYoutubeMediaUpdate } from './pending-youtube-media';
import type { MediaGenerationInputResolverOptions } from './media-source';

type PendingYoutubeMediaUpdateResult = 'updated' | 'partial' | 'failed';

export interface PendingYoutubeMediaQueueReadyOptions {
  notifyNoQueued?: boolean;
}

export interface PendingYoutubeMediaQueueFailedOptions {
  notifyStatus?: boolean;
}

export interface PendingYoutubeMediaNoteInfo {
  noteId: number;
  fields: Record<string, { value: string }>;
}

export interface PendingYoutubeMediaQueueDeps {
  client: {
    notesInfo(noteIds: number[]): Promise<unknown>;
    updateNoteFields(noteId: number, fields: Record<string, string>): Promise<void>;
    storeMediaFile(filename: string, data: Buffer): Promise<void>;
  };
  mediaGenerator: Pick<
    MediaGenerator,
    'generateAudio' | 'generateScreenshot' | 'generateAnimatedImage'
  >;
  getConfig: () => AnkiConnectConfig;
  getCurrentVideoPath: () => string | undefined;
  getCachedMediaPath: MediaGenerationInputResolverOptions['getCachedMediaPath'] | null;
  shouldRequireRemoteMediaCache: () => boolean;
  getSubtitleMediaRange: (context?: SubtitleMiningContext) => {
    startTime: number;
    endTime: number;
  };
  getResolvedSentenceAudioFieldName: (noteInfo: PendingYoutubeMediaNoteInfo) => string | null;
  resolveConfiguredFieldName: (
    noteInfo: PendingYoutubeMediaNoteInfo,
    ...preferredNames: (string | undefined)[]
  ) => string | null;
  mergeFieldValue: (existing: string, newValue: string, overwrite: boolean) => string;
  getAnimatedImageLeadInSeconds: (noteInfo: PendingYoutubeMediaNoteInfo) => Promise<number>;
  generateAudioFilename: () => string;
  generateImageFilename: () => string;
  formatMiscInfoPatternForMediaPath: (
    fallbackFilename: string,
    startTimeSeconds: number | undefined,
    mediaPath: string,
    mediaTitle?: string,
  ) => string;
  showStatusNotification: (message: string) => void;
  showNotification: (noteId: number, label: string | number, errorSuffix?: string) => Promise<void>;
  logInfo: (message: string, ...args: unknown[]) => void;
  logWarn: (message: string, ...args: unknown[]) => void;
  logError: (message: string, ...args: unknown[]) => void;
}

function trimToNonEmptyString(value: unknown): string | null {
  if (typeof value !== 'string') {
    return null;
  }
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function shouldGenerateAudio(config: AnkiConnectConfig): boolean {
  return config.media?.generateAudio !== false;
}

function shouldGenerateImage(config: AnkiConnectConfig): boolean {
  return config.media?.generateImage !== false;
}

export class PendingYoutubeMediaQueue {
  private updates: PendingYoutubeMediaUpdate[] = [];

  constructor(private readonly deps: PendingYoutubeMediaQueueDeps) {}

  enqueue(job: PendingYoutubeMediaUpdate): void {
    if (!job.generateAudio && !job.generateImage) {
      return;
    }
    const isFirstQueuedForSource = !this.updates.some((existing) =>
      youtubeMediaUrlsMatch(existing.sourceUrl, job.sourceUrl),
    );
    this.updates.push(job);
    this.deps.logInfo('Queued YouTube media update for note:', job.noteId);
    if (isFirstQueuedForSource) {
      this.deps.showStatusNotification(
        'YouTube media cache is still downloading. Card media will be added when the cache is ready.',
      );
    }
  }

  async queueFromNote(job: {
    noteId: number;
    noteInfo: PendingYoutubeMediaNoteInfo;
    context?: SubtitleMiningContext;
    label: string | number;
  }): Promise<boolean> {
    const sourceUrl = trimToNonEmptyString(this.deps.getCurrentVideoPath());
    const getCachedMediaPath = this.deps.getCachedMediaPath;
    if (!sourceUrl || this.deps.shouldRequireRemoteMediaCache() !== true || !getCachedMediaPath) {
      return false;
    }

    let cachedPath: string | null = null;
    try {
      cachedPath = await getCachedMediaPath(sourceUrl, 'video');
    } catch (error) {
      this.deps.logWarn(
        'Failed to read YouTube cache state; falling back to immediate media generation:',
        error instanceof Error ? error.message : String(error),
      );
      return false;
    }
    if (cachedPath) {
      return false;
    }

    const config = this.deps.getConfig();
    const mediaRange = this.deps.getSubtitleMediaRange(job.context);
    this.enqueue({
      sourceUrl,
      noteId: job.noteId,
      startTime: mediaRange.startTime,
      endTime: mediaRange.endTime,
      label: job.label,
      audioFieldName: this.deps.getResolvedSentenceAudioFieldName(job.noteInfo) ?? undefined,
      imageFieldName:
        this.deps.resolveConfiguredFieldName(
          job.noteInfo,
          config.fields?.image,
          DEFAULT_ANKI_CONNECT_CONFIG.fields.image,
        ) ?? undefined,
      miscInfoFieldName:
        this.deps.resolveConfiguredFieldName(job.noteInfo, config.fields?.miscInfo) ?? undefined,
      generateAudio: shouldGenerateAudio(config),
      generateImage: shouldGenerateImage(config),
    });
    return true;
  }

  async handleReady(
    sourceUrl: string,
    cachedPath: string,
    options: PendingYoutubeMediaQueueReadyOptions = {},
  ): Promise<void> {
    const jobs = this.takeMatchingUpdates(sourceUrl);
    if (jobs.length === 0) {
      if (options.notifyNoQueued !== false) {
        this.deps.showStatusNotification('YouTube media cache ready.');
      }
      return;
    }

    this.deps.showStatusNotification(
      `YouTube media cache ready. Adding media to ${jobs.length} queued card${
        jobs.length === 1 ? '' : 's'
      }.`,
    );

    let updatedCount = 0;
    let partialCount = 0;
    let failedCount = 0;
    for (const job of jobs) {
      try {
        const result = await this.applyUpdate(job, cachedPath);
        if (result === 'updated') {
          updatedCount += 1;
        } else if (result === 'partial') {
          partialCount += 1;
        } else {
          failedCount += 1;
        }
      } catch (error) {
        failedCount += 1;
        this.deps.logError(
          'Failed to apply queued YouTube media update:',
          error instanceof Error ? error.message : String(error),
        );
      }
    }

    if (partialCount > 0 || failedCount > 0) {
      this.deps.showStatusNotification(
        `Queued YouTube media finished with ${updatedCount} updated, ${partialCount} partial, and ${failedCount} failed.`,
      );
    }
  }

  async handleFailed(
    sourceUrl: string,
    options: PendingYoutubeMediaQueueFailedOptions = {},
  ): Promise<void> {
    const jobs = this.takeMatchingUpdates(sourceUrl);
    if (jobs.length === 0) {
      if (options.notifyStatus !== false) {
        this.deps.showStatusNotification('YouTube media cache failed.');
      }
      return;
    }

    if (options.notifyStatus !== false) {
      this.deps.showStatusNotification(
        `YouTube media cache failed. Media was not added to ${jobs.length} queued card${
          jobs.length === 1 ? '' : 's'
        }.`,
      );
    }
    this.deps.logWarn('Discarding queued YouTube media updates after cache failure:', jobs.length);

    for (const job of jobs) {
      try {
        await this.deps.showNotification(job.noteId, job.label, 'media cache failed');
      } catch (error) {
        this.deps.logWarn(
          'Failed to show queued YouTube media failure notification:',
          error instanceof Error ? error.message : String(error),
        );
      }
    }
  }

  private takeMatchingUpdates(sourceUrl: string): PendingYoutubeMediaUpdate[] {
    const matched: PendingYoutubeMediaUpdate[] = [];
    const remaining: PendingYoutubeMediaUpdate[] = [];
    for (const job of this.updates) {
      if (youtubeMediaUrlsMatch(job.sourceUrl, sourceUrl)) {
        matched.push(job);
      } else {
        remaining.push(job);
      }
    }
    this.updates = remaining;
    return matched;
  }

  private async applyUpdate(
    job: PendingYoutubeMediaUpdate,
    cachedPath: string,
  ): Promise<PendingYoutubeMediaUpdateResult> {
    const notesInfoResult = await this.deps.client.notesInfo([job.noteId]);
    const notesInfo = notesInfoResult as unknown as PendingYoutubeMediaNoteInfo[];
    const noteInfo = notesInfo[0];
    if (!noteInfo) {
      this.deps.logWarn('Queued YouTube media target note not found:', job.noteId);
      return 'failed';
    }

    const config = this.deps.getConfig();
    const mediaFields: Record<string, string> = {};
    const errors: string[] = [];
    let miscInfoFilename: string | null = null;
    const cachedMediaInput: MediaInput = {
      path: cachedPath,
      source: 'youtube-cache',
    };

    if (job.generateAudio) {
      try {
        const audioFilename = this.deps.generateAudioFilename();
        const audioBuffer = await this.deps.mediaGenerator.generateAudio(
          cachedMediaInput,
          job.startTime,
          job.endTime,
          config.media?.audioPadding,
          undefined,
        );
        if (audioBuffer) {
          await this.deps.client.storeMediaFile(audioFilename, audioBuffer);
          const audioField =
            job.audioFieldName || this.deps.getResolvedSentenceAudioFieldName(noteInfo) || null;
          if (audioField) {
            const existingAudio = noteInfo.fields[audioField]?.value || '';
            mediaFields[audioField] = this.deps.mergeFieldValue(
              existingAudio,
              `[sound:${audioFilename}]`,
              config.behavior?.overwriteAudio !== false,
            );
          }
          miscInfoFilename = audioFilename;
        }
      } catch (error) {
        errors.push('audio');
        this.deps.logError('Failed to generate queued YouTube audio:', (error as Error).message);
      }
    }

    if (job.generateImage) {
      try {
        const imageFilename = this.deps.generateImageFilename();
        const animatedLeadInSeconds = await this.deps.getAnimatedImageLeadInSeconds(noteInfo);
        const imageBuffer = await this.generateImageFromInput(
          cachedMediaInput,
          job.startTime,
          job.endTime,
          animatedLeadInSeconds,
        );
        if (imageBuffer) {
          await this.deps.client.storeMediaFile(imageFilename, imageBuffer);
          const imageField =
            job.imageFieldName ||
            this.deps.resolveConfiguredFieldName(
              noteInfo,
              config.fields?.image,
              DEFAULT_ANKI_CONNECT_CONFIG.fields.image,
            );
          if (imageField) {
            const existingImage = noteInfo.fields[imageField]?.value || '';
            mediaFields[imageField] = this.deps.mergeFieldValue(
              existingImage,
              `<img src="${imageFilename}">`,
              config.behavior?.overwriteImage !== false,
            );
          } else {
            this.deps.logWarn(
              'Image field not found on queued YouTube media note, skipping image update',
            );
          }
          miscInfoFilename = imageFilename;
        }
      } catch (error) {
        errors.push('image');
        this.deps.logError('Failed to generate queued YouTube image:', (error as Error).message);
      }
    }

    if (config.fields?.miscInfo && miscInfoFilename) {
      const miscInfoField =
        job.miscInfoFieldName ||
        this.deps.resolveConfiguredFieldName(noteInfo, config.fields.miscInfo);
      const miscInfo = this.deps.formatMiscInfoPatternForMediaPath(
        miscInfoFilename,
        job.startTime,
        job.sourceUrl,
      );
      if (miscInfoField && miscInfo) {
        mediaFields[miscInfoField] = miscInfo;
      }
    }

    if (Object.keys(mediaFields).length === 0) {
      return 'failed';
    }

    await this.deps.client.updateNoteFields(job.noteId, mediaFields);
    const errorSuffix = errors.length > 0 ? `${errors.join(', ')} failed` : undefined;
    await this.deps.showNotification(job.noteId, job.label, errorSuffix);
    this.deps.logInfo('Applied queued YouTube media update for note:', job.noteId);
    return errors.length === 0 ? 'updated' : 'partial';
  }

  private async generateImageFromInput(
    videoPath: MediaInput,
    startTime: number,
    endTime: number,
    animatedLeadInSeconds = 0,
  ): Promise<Buffer | null> {
    const config = this.deps.getConfig();
    if (config.media?.imageType === 'avif') {
      return this.deps.mediaGenerator.generateAnimatedImage(
        videoPath,
        startTime,
        endTime,
        config.media?.audioPadding,
        {
          fps: config.media?.animatedFps,
          maxWidth: config.media?.animatedMaxWidth,
          maxHeight: config.media?.animatedMaxHeight,
          crf: config.media?.animatedCrf,
          leadingStillDuration: animatedLeadInSeconds,
        },
      );
    }

    const timestamp = startTime + (endTime - startTime) / 2;
    return this.deps.mediaGenerator.generateScreenshot(videoPath, timestamp, {
      format: config.media?.imageFormat as 'jpg' | 'png' | 'webp',
      quality: config.media?.imageQuality,
      maxWidth: config.media?.imageMaxWidth,
      maxHeight: config.media?.imageMaxHeight,
    });
  }
}
