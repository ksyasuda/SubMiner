import { DEFAULT_ANKI_CONNECT_CONFIG } from '../config';

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
      sentence?: string;
      image?: string;
      miscInfo?: string;
    };
    media?: {
      generateAudio?: boolean;
      generateImage?: boolean;
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
  resolveConfiguredFieldName: (
    noteInfo: NoteUpdateWorkflowNoteInfo,
    ...preferredNames: (string | undefined)[]
  ) => string | null;
  getResolvedSentenceAudioFieldName: (noteInfo: NoteUpdateWorkflowNoteInfo) => string | null;
  mergeFieldValue: (existing: string, newValue: string, overwrite: boolean) => string;
  generateAudioFilename: () => string;
  generateAudio: () => Promise<Buffer | null>;
  generateImageFilename: () => string;
  generateImage: () => Promise<Buffer | null>;
  formatMiscInfoPattern: (fallbackFilename: string, startTimeSeconds?: number) => string;
  addConfiguredTagsToNote: (noteId: number) => Promise<void>;
  showNotification: (noteId: number, label: string | number) => Promise<void>;
  showOsdNotification: (message: string) => void;
  beginUpdateProgress: (initialMessage: string) => void;
  endUpdateProgress: () => void;
  logWarn: (message: string, ...args: unknown[]) => void;
  logInfo: (message: string, ...args: unknown[]) => void;
  logError: (message: string, ...args: unknown[]) => void;
}

export class NoteUpdateWorkflow {
  constructor(private readonly deps: NoteUpdateWorkflowDeps) {}

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

      const expressionText = (fields.expression || fields.word || '').trim();
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

      const currentSubtitleText = this.deps.getCurrentSubtitleText();
      if (sentenceField && currentSubtitleText) {
        const processedSentence = this.deps.processSentence(currentSubtitleText, fields);
        updatedFields[sentenceField] = processedSentence;
        updatePerformed = true;
      }

      const config = this.deps.getConfig();

      if (config.media?.generateAudio) {
        try {
          const audioFilename = this.deps.generateAudioFilename();
          const audioBuffer = await this.deps.generateAudio();

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

      if (config.media?.generateImage) {
        try {
          const imageFilename = this.deps.generateImageFilename();
          const imageBuffer = await this.deps.generateImage();

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

      if (config.fields?.miscInfo) {
        const miscInfo = this.deps.formatMiscInfoPattern(
          miscInfoFilename || '',
          this.deps.getCurrentSubtitleStart(),
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
        this.deps.logInfo('Updated card fields for:', hasExpressionText ? expressionText : noteId);
        await this.deps.showNotification(noteId, hasExpressionText ? expressionText : noteId);
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
