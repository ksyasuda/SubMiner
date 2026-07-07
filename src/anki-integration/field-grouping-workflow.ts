import { KikuDuplicateCardInfo, KikuFieldGroupingChoice } from '../types/anki';
import { getPreferredWordValueFromExtractedFields } from '../anki-field-config';

export interface FieldGroupingWorkflowNoteInfo {
  noteId: number;
  fields: Record<string, { value: string }>;
  tags?: string[];
}

export interface FieldGroupingWorkflowDeps {
  client: {
    notesInfo(noteIds: number[]): Promise<unknown>;
    updateNoteFields(noteId: number, fields: Record<string, string>): Promise<void>;
    addTags(noteIds: number[], tags: string[]): Promise<void>;
    deleteNotes(noteIds: number[]): Promise<void>;
  };
  getConfig: () => {
    fields?: {
      word?: string;
      audio?: string;
      image?: string;
    };
  };
  getEffectiveSentenceCardConfig: () => {
    sentenceField: string;
    audioField: string;
    kikuDeleteDuplicateInAuto: boolean;
  };
  getCurrentSubtitleText: () => string | undefined;
  getFieldGroupingCallback:
    | (() => Promise<
        | ((data: {
            original: KikuDuplicateCardInfo;
            duplicate: KikuDuplicateCardInfo;
          }) => Promise<KikuFieldGroupingChoice>)
        | null
      >)
    | (() =>
        | ((data: {
            original: KikuDuplicateCardInfo;
            duplicate: KikuDuplicateCardInfo;
          }) => Promise<KikuFieldGroupingChoice>)
        | null);
  computeFieldGroupingMergedFields: (
    keepNoteId: number,
    deleteNoteId: number,
    keepNoteInfo: FieldGroupingWorkflowNoteInfo,
    deleteNoteInfo: FieldGroupingWorkflowNoteInfo,
    includeGeneratedMedia: boolean,
  ) => Promise<Record<string, string>>;
  extractFields: (fields: Record<string, { value: string }>) => Record<string, string>;
  hasFieldValue: (noteInfo: FieldGroupingWorkflowNoteInfo, preferredFieldName?: string) => boolean;
  addConfiguredTagsToNote: (noteId: number) => Promise<void>;
  removeTrackedNoteId: (noteId: number) => void;
  rememberMergedNoteIds: (deletedNoteId: number, keptNoteId: number) => void;
  showStatusNotification: (message: string) => void;
  showNotification: (noteId: number, label: string | number) => Promise<void>;
  showOsdNotification: (message: string) => void;
  logError: (message: string, ...args: unknown[]) => void;
  logInfo: (message: string, ...args: unknown[]) => void;
  truncateSentence: (sentence: string) => string;
}

export class FieldGroupingWorkflow {
  constructor(private readonly deps: FieldGroupingWorkflowDeps) {}

  async handleAuto(
    originalNoteId: number,
    newNoteId: number,
    newNoteInfo: FieldGroupingWorkflowNoteInfo,
  ): Promise<void> {
    try {
      const sentenceCardConfig = this.deps.getEffectiveSentenceCardConfig();
      await this.performMerge(
        originalNoteId,
        newNoteId,
        this.getExpression(newNoteInfo),
        sentenceCardConfig.kikuDeleteDuplicateInAuto,
      );
    } catch (error) {
      this.deps.logError('Field grouping auto merge failed:', (error as Error).message);
      this.deps.showOsdNotification(`Field grouping failed: ${(error as Error).message}`);
    }
  }

  async handleManual(
    originalNoteId: number,
    _newNoteId: number,
    newNoteInfo: FieldGroupingWorkflowNoteInfo,
  ): Promise<boolean> {
    const callback = await this.resolveFieldGroupingCallback();
    if (!callback) {
      this.deps.showOsdNotification('Field grouping UI unavailable');
      return false;
    }

    try {
      const originalNotesInfoResult = await this.deps.client.notesInfo([originalNoteId]);
      const originalNotesInfo = originalNotesInfoResult as FieldGroupingWorkflowNoteInfo[];
      if (!originalNotesInfo || originalNotesInfo.length === 0) {
        // handleManual owns all user-facing notifications; callers must not re-notify on false.
        this.deps.showOsdNotification('Field grouping failed: original card not found');
        return false;
      }

      const originalNoteInfo = originalNotesInfo[0]!;
      const expression = this.getExpression(newNoteInfo) || this.getExpression(originalNoteInfo);

      const choice = await callback({
        original: this.buildDuplicateCardInfo(originalNoteInfo, expression, true),
        duplicate: this.buildDuplicateCardInfo(newNoteInfo, expression, false),
      });

      if (choice.cancelled) {
        this.deps.showOsdNotification('Field grouping cancelled');
        return false;
      }

      const keepNoteId = choice.keepNoteId;
      const deleteNoteId = choice.deleteNoteId;

      await this.performMerge(keepNoteId, deleteNoteId, expression, choice.deleteDuplicate);
      return true;
    } catch (error) {
      this.deps.logError('Field grouping manual merge failed:', (error as Error).message);
      this.deps.showOsdNotification(`Field grouping failed: ${(error as Error).message}`);
      return false;
    }
  }

  private async performMerge(
    keepNoteId: number,
    deleteNoteId: number,
    expression: string,
    deleteDuplicate = true,
  ): Promise<void> {
    const notesInfoResult = await this.deps.client.notesInfo([keepNoteId, deleteNoteId]);
    const notesInfo = notesInfoResult as FieldGroupingWorkflowNoteInfo[];
    const keepNoteInfo = notesInfo.find((note) => note.noteId === keepNoteId);
    const deleteNoteInfo = notesInfo.find((note) => note.noteId === deleteNoteId);
    if (!keepNoteInfo) {
      this.deps.logInfo('Keep note not found:', keepNoteId);
      return;
    }
    if (!deleteNoteInfo) {
      this.deps.logInfo('Delete note not found:', deleteNoteId);
      return;
    }

    const mergedFields = await this.deps.computeFieldGroupingMergedFields(
      keepNoteId,
      deleteNoteId,
      keepNoteInfo,
      deleteNoteInfo,
      true,
    );

    if (Object.keys(mergedFields).length > 0) {
      await this.deps.client.updateNoteFields(keepNoteId, mergedFields);
      await this.deps.addConfiguredTagsToNote(keepNoteId);
    }

    const tagsToAdd = this.getMergeTagsToAdd(keepNoteInfo, deleteNoteInfo);
    if (tagsToAdd.length > 0) {
      await this.deps.client.addTags([keepNoteId], tagsToAdd);
    }

    if (deleteDuplicate) {
      await this.deps.client.deleteNotes([deleteNoteId]);
      this.deps.removeTrackedNoteId(deleteNoteId);
      this.deps.rememberMergedNoteIds(deleteNoteId, keepNoteId);
    }

    this.deps.logInfo('Merged duplicate card:', expression, 'into note:', keepNoteId);
    this.deps.showStatusNotification(
      deleteDuplicate
        ? `Merged duplicate: ${expression}`
        : `Grouped duplicate (kept both): ${expression}`,
    );
    await this.deps.showNotification(keepNoteId, expression);
  }

  private buildDuplicateCardInfo(
    noteInfo: FieldGroupingWorkflowNoteInfo,
    fallbackExpression: string,
    isOriginal: boolean,
  ): KikuDuplicateCardInfo {
    const sentenceCardConfig = this.deps.getEffectiveSentenceCardConfig();
    const fields = this.deps.extractFields(noteInfo.fields);
    return {
      noteId: noteInfo.noteId,
      expression:
        getPreferredWordValueFromExtractedFields(fields, this.deps.getConfig()) ||
        fallbackExpression,
      sentencePreview: this.deps.truncateSentence(
        fields[(sentenceCardConfig.sentenceField || 'sentence').toLowerCase()] ||
          (isOriginal ? '' : this.deps.getCurrentSubtitleText() || ''),
      ),
      hasAudio:
        this.deps.hasFieldValue(noteInfo, this.deps.getConfig().fields?.audio) ||
        this.deps.hasFieldValue(noteInfo, sentenceCardConfig.audioField),
      hasImage: this.deps.hasFieldValue(noteInfo, this.deps.getConfig().fields?.image),
      isOriginal,
    };
  }

  private getExpression(noteInfo: FieldGroupingWorkflowNoteInfo): string {
    const fields = this.deps.extractFields(noteInfo.fields);
    return getPreferredWordValueFromExtractedFields(fields, this.deps.getConfig());
  }

  private getMergeTagsToAdd(
    keepNoteInfo: FieldGroupingWorkflowNoteInfo,
    deleteNoteInfo: FieldGroupingWorkflowNoteInfo,
  ): string[] {
    const targetTags = new Set((keepNoteInfo.tags ?? []).map((tag) => tag.trim()).filter(Boolean));
    const unwantedSourceTags = new Set(['leech', 'marked', 'potential_leech']);
    const tagsToAdd: string[] = [];

    for (const rawTag of deleteNoteInfo.tags ?? []) {
      const tag = rawTag.trim();
      if (!tag || targetTags.has(tag) || unwantedSourceTags.has(tag)) continue;
      targetTags.add(tag);
      tagsToAdd.push(tag);
    }

    return tagsToAdd;
  }

  private async resolveFieldGroupingCallback(): Promise<
    | ((data: {
        original: KikuDuplicateCardInfo;
        duplicate: KikuDuplicateCardInfo;
      }) => Promise<KikuFieldGroupingChoice>)
    | null
  > {
    const callback = this.deps.getFieldGroupingCallback();
    if (callback instanceof Promise) {
      return callback;
    }
    return callback;
  }
}
