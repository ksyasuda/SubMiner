import { DEFAULT_ANKI_CONNECT_CONFIG } from '../config';
import { getPreferredWordValueFromExtractedFields } from '../anki-field-config';
import type { SubtitleMiningContext } from '../types/subtitle';
import { shouldMarkWordAndSentenceCard } from './note-field-utils';

export interface NoteUpdateWorkflowNoteInfo {
  noteId: number;
  fields: Record<string, { value: string }>;
}

export interface NoteUpdateWorkflowDeps {
  client: {
    notesInfo(noteIds: number[]): Promise<unknown>;
    updateNoteFields(noteId: number, fields: Record<string, string>): Promise<void>;
    storeMediaFile(filename: string, data: Buffer): Promise<void>;
  };
  getConfig: () => {
    fields?: {
      word?: string;
      sentence?: string;
      image?: string;
      miscInfo?: string;
    };
    media?: {
      generateAudio?: boolean;
      generateImage?: boolean;
      imageType?: 'static' | 'avif';
      syncAnimatedImageToWordAudio?: boolean;
    };
    behavior?: {
      overwriteAudio?: boolean;
      overwriteImage?: boolean;
    };
  };
  getCurrentSubtitleText: () => string | undefined;
  getCurrentSubtitleStart: () => number | undefined;
  getEffectiveSentenceCardConfig: () => {
    sentenceField: string;
    lapisEnabled: boolean;
    kikuEnabled: boolean;
    kikuFieldGrouping: 'auto' | 'manual' | 'disabled';
  };
  appendKnownWordsFromNoteInfo: (noteInfo: NoteUpdateWorkflowNoteInfo) => void;
  extractFields: (fields: Record<string, { value: string }>) => Record<string, string>;
  findDuplicateNote: (
    expression: string,
    excludeNoteId: number,
    noteInfo: NoteUpdateWorkflowNoteInfo,
  ) => Promise<number | null>;
  handleFieldGroupingAuto: (
    originalNoteId: number,
    newNoteId: number,
    newNoteInfo: NoteUpdateWorkflowNoteInfo,
    expression: string,
  ) => Promise<void>;
  handleFieldGroupingManual: (
    originalNoteId: number,
    newNoteId: number,
    newNoteInfo: NoteUpdateWorkflowNoteInfo,
    expression: string,
  ) => Promise<boolean>;
  processSentence: (mpvSentence: string, noteFields: Record<string, string>) => string;
  processSentenceFurigana?: (
    sentenceFurigana: string,
    noteFields: Record<string, string>,
  ) => string;
  setCardTypeFields: (
    updatedFields: Record<string, string>,
    availableFieldNames: string[],
    cardKind: 'word-and-sentence',
  ) => void;
  resolveConfiguredFieldName: (
    noteInfo: NoteUpdateWorkflowNoteInfo,
    ...preferredNames: (string | undefined)[]
  ) => string | null;
  getResolvedSentenceAudioFieldName: (noteInfo: NoteUpdateWorkflowNoteInfo) => string | null;
  getAnimatedImageLeadInSeconds: (noteInfo: NoteUpdateWorkflowNoteInfo) => Promise<number>;
  mergeFieldValue: (existing: string, newValue: string, overwrite: boolean) => string;
  generateAudioFilename: () => string;
  generateAudio: (context?: SubtitleMiningContext) => Promise<Buffer | null>;
  generateImageFilename: () => string;
  generateImage: (
    animatedLeadInSeconds?: number,
    context?: SubtitleMiningContext,
  ) => Promise<Buffer | null>;
  formatMiscInfoPattern: (fallbackFilename: string, startTimeSeconds?: number) => string;
  consumeSubtitleMiningContext?: () => SubtitleMiningContext | null;
  queuePendingYoutubeMediaUpdate?: (job: {
    noteId: number;
    noteInfo: NoteUpdateWorkflowNoteInfo;
    context?: SubtitleMiningContext;
    label: string | number;
  }) => Promise<boolean>;
  addConfiguredTagsToNote: (noteId: number) => Promise<void>;
  showNotification: (noteId: number, label: string | number) => Promise<void>;
  showOsdNotification: (message: string) => void;
  beginUpdateProgress: (initialMessage: string) => void;
  endUpdateProgress: () => void;
  logWarn: (message: string, ...args: unknown[]) => void;
  logInfo: (message: string, ...args: unknown[]) => void;
  logError: (message: string, ...args: unknown[]) => void;
}

function normalizeSubtitleContextText(text: string): string {
  return text
    .replace(/<[^>]*>/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function hasUsableSubtitleContextTiming(context: SubtitleMiningContext): boolean {
  return (
    Number.isFinite(context.startTime) &&
    Number.isFinite(context.endTime) &&
    context.endTime > context.startTime
  );
}

function subtitleContextMatchesSentence(contextText: string, noteSentence: string): boolean {
  const normalizedContext = normalizeSubtitleContextText(contextText);
  const normalizedSentence = normalizeSubtitleContextText(noteSentence);
  if (!normalizedContext || !normalizedSentence) {
    return false;
  }
  return (
    normalizedContext === normalizedSentence ||
    normalizedContext.includes(normalizedSentence) ||
    normalizedSentence.includes(normalizedContext)
  );
}

export class NoteUpdateWorkflow {
  constructor(private readonly deps: NoteUpdateWorkflowDeps) {}

  private consumeMatchingSubtitleMiningContext(
    fields: Record<string, string>,
    sentenceField: string,
    configuredSentenceField?: string,
  ): SubtitleMiningContext | null {
    const context = this.deps.consumeSubtitleMiningContext?.() ?? null;
    if (!context || !hasUsableSubtitleContextTiming(context)) {
      return null;
    }

    const candidateFields = [
      sentenceField,
      configuredSentenceField,
      DEFAULT_ANKI_CONNECT_CONFIG.fields.sentence,
    ];
    const noteSentence = candidateFields
      .map((fieldName) => (fieldName ? fields[fieldName.toLowerCase()] : undefined))
      .find((value): value is string => typeof value === 'string' && value.trim().length > 0);

    if (!noteSentence || subtitleContextMatchesSentence(context.text, noteSentence)) {
      return context;
    }
    return null;
  }

  async execute(noteId: number, options?: { skipKikuFieldGrouping?: boolean }): Promise<void> {
    this.deps.beginUpdateProgress('Updating card');
    try {
      const notesInfoResult = await this.deps.client.notesInfo([noteId]);
      const notesInfo = notesInfoResult as NoteUpdateWorkflowNoteInfo[];
      if (!notesInfo || notesInfo.length === 0) {
        this.deps.logWarn('Card not found:', noteId);
        return;
      }

      const noteInfo = notesInfo[0]!;
      this.deps.appendKnownWordsFromNoteInfo(noteInfo);
      const fields = this.deps.extractFields(noteInfo.fields);
      const config = this.deps.getConfig();

      const expressionText = getPreferredWordValueFromExtractedFields(fields, config).trim();
      const hasExpressionText = expressionText.length > 0;
      if (!hasExpressionText) {
        // Some note types omit Expression/Word; still run enrichment updates and skip duplicate checks.
        this.deps.logWarn(
          'No expression/word field found in card; skipping duplicate checks but continuing update:',
          noteId,
        );
      }

      const sentenceCardConfig = this.deps.getEffectiveSentenceCardConfig();
      const shouldRunFieldGrouping =
        !options?.skipKikuFieldGrouping &&
        sentenceCardConfig.kikuEnabled &&
        sentenceCardConfig.kikuFieldGrouping !== 'disabled';
      let duplicateNoteId: number | null = null;
      if (shouldRunFieldGrouping && hasExpressionText) {
        duplicateNoteId = await this.deps.findDuplicateNote(expressionText, noteId, noteInfo);
      }

      const updatedFields: Record<string, string> = {};
      let updatePerformed = false;
      let miscInfoFilename: string | null = null;
      const sentenceField = sentenceCardConfig.sentenceField;
      const subtitleMiningContext = this.consumeMatchingSubtitleMiningContext(
        fields,
        sentenceField,
        config.fields?.sentence,
      );
      const noteLabel = hasExpressionText ? expressionText : noteId;

      const currentSubtitleText = subtitleMiningContext?.text ?? this.deps.getCurrentSubtitleText();
      if (sentenceField && currentSubtitleText) {
        const processedSentence = this.deps.processSentence(currentSubtitleText, fields);
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
      const sentenceFuriganaField = this.deps.resolveConfiguredFieldName(
        noteInfo,
        'SentenceFurigana',
      );
      const existingSentenceFurigana = sentenceFuriganaField
        ? noteInfo.fields[sentenceFuriganaField]?.value || ''
        : '';
      if (sentenceFuriganaField && existingSentenceFurigana && this.deps.processSentenceFurigana) {
        const processedSentenceFurigana = this.deps.processSentenceFurigana(
          existingSentenceFurigana,
          fields,
        );
        if (processedSentenceFurigana !== existingSentenceFurigana) {
          updatedFields[sentenceFuriganaField] = processedSentenceFurigana;
          updatePerformed = true;
        }
      }

      const mediaCacheQueued =
        (config.media?.generateAudio || config.media?.generateImage) &&
        this.deps.queuePendingYoutubeMediaUpdate
          ? await this.deps.queuePendingYoutubeMediaUpdate({
              noteId,
              noteInfo,
              context: subtitleMiningContext ?? undefined,
              label: noteLabel,
            })
          : false;

      if (!mediaCacheQueued && config.media?.generateAudio) {
        try {
          const audioFilename = this.deps.generateAudioFilename();
          const audioBuffer = await this.deps.generateAudio(subtitleMiningContext ?? undefined);

          if (audioBuffer) {
            await this.deps.client.storeMediaFile(audioFilename, audioBuffer);
            const sentenceAudioField = this.deps.getResolvedSentenceAudioFieldName(noteInfo);
            if (sentenceAudioField) {
              const existingAudio = noteInfo.fields[sentenceAudioField]?.value || '';
              updatedFields[sentenceAudioField] = this.deps.mergeFieldValue(
                existingAudio,
                `[sound:${audioFilename}]`,
                config.behavior?.overwriteAudio !== false,
              );
            }
            miscInfoFilename = audioFilename;
            updatePerformed = true;
          }
        } catch (error) {
          this.deps.logError('Failed to generate audio:', (error as Error).message);
          this.deps.showOsdNotification(`Audio generation failed: ${(error as Error).message}`);
        }
      }

      if (!mediaCacheQueued && config.media?.generateImage) {
        try {
          const animatedLeadInSeconds = await this.deps.getAnimatedImageLeadInSeconds(noteInfo);
          const imageFilename = this.deps.generateImageFilename();
          const imageBuffer = await this.deps.generateImage(
            animatedLeadInSeconds,
            subtitleMiningContext ?? undefined,
          );

          if (imageBuffer) {
            await this.deps.client.storeMediaFile(imageFilename, imageBuffer);
            const imageFieldName = this.deps.resolveConfiguredFieldName(
              noteInfo,
              config.fields?.image,
              DEFAULT_ANKI_CONNECT_CONFIG.fields.image,
            );
            if (!imageFieldName) {
              this.deps.logWarn('Image field not found on note, skipping image update');
            } else {
              const existingImage = noteInfo.fields[imageFieldName]?.value || '';
              updatedFields[imageFieldName] = this.deps.mergeFieldValue(
                existingImage,
                `<img src="${imageFilename}">`,
                config.behavior?.overwriteImage !== false,
              );
              miscInfoFilename = imageFilename;
              updatePerformed = true;
            }
          }
        } catch (error) {
          this.deps.logError('Failed to generate image:', (error as Error).message);
          this.deps.showOsdNotification(`Image generation failed: ${(error as Error).message}`);
        }
      }

      if (!mediaCacheQueued && config.fields?.miscInfo) {
        const miscInfo = this.deps.formatMiscInfoPattern(
          miscInfoFilename || '',
          subtitleMiningContext?.startTime ?? this.deps.getCurrentSubtitleStart(),
        );
        const miscInfoField = this.deps.resolveConfiguredFieldName(
          noteInfo,
          config.fields?.miscInfo,
        );
        if (miscInfo && miscInfoField) {
          updatedFields[miscInfoField] = miscInfo;
          updatePerformed = true;
        }
      }

      if (updatePerformed) {
        await this.deps.client.updateNoteFields(noteId, updatedFields);
        await this.deps.addConfiguredTagsToNote(noteId);
        this.deps.logInfo('Updated card fields for:', noteLabel);
        await this.deps.showNotification(noteId, noteLabel);
      }

      if (shouldRunFieldGrouping && hasExpressionText && duplicateNoteId !== null) {
        let noteInfoForGrouping = noteInfo;
        if (updatePerformed) {
          const refreshedInfoResult = await this.deps.client.notesInfo([noteId]);
          const refreshedInfo = refreshedInfoResult as NoteUpdateWorkflowNoteInfo[];
          if (!refreshedInfo || refreshedInfo.length === 0) {
            this.deps.logWarn('Card not found after update:', noteId);
            return;
          }
          noteInfoForGrouping = refreshedInfo[0]!;
        }

        if (sentenceCardConfig.kikuFieldGrouping === 'auto') {
          await this.deps.handleFieldGroupingAuto(
            duplicateNoteId,
            noteId,
            noteInfoForGrouping,
            expressionText,
          );
          return;
        }
        if (sentenceCardConfig.kikuFieldGrouping === 'manual') {
          await this.deps.handleFieldGroupingManual(
            duplicateNoteId,
            noteId,
            noteInfoForGrouping,
            expressionText,
          );
        }
      }
    } catch (error) {
      if ((error as Error).message.includes('note was not found')) {
        this.deps.logWarn('Card was deleted before update:', noteId);
      } else {
        this.deps.logError('Error processing new card:', (error as Error).message);
      }
    } finally {
      this.deps.endUpdateProgress();
    }
  }
}
