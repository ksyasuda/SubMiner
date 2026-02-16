import { KikuMergePreviewResponse } from "../types";
import { createLogger } from "../logger";

const log = createLogger("anki").child("integration.field-grouping");

interface FieldGroupingNoteInfo {
  noteId: number;
  fields: Record<string, { value: string }>;
}

interface FieldGroupingDeps {
  getEffectiveSentenceCardConfig: () => {
    model?: string;
    sentenceField: string;
    audioField: string;
    lapisEnabled: boolean;
    kikuEnabled: boolean;
    kikuFieldGrouping: "auto" | "manual" | "disabled";
    kikuDeleteDuplicateInAuto: boolean;
  };
  isUpdateInProgress: () => boolean;
  getDeck?: () => string | undefined;
  withUpdateProgress: <T>(initialMessage: string, action: () => Promise<T>) => Promise<T>;
  showOsdNotification: (text: string) => void;
  findNotes: (
    query: string,
    options?: {
      maxRetries?: number;
    },
  ) => Promise<number[]>;
  notesInfo: (noteIds: number[]) => Promise<FieldGroupingNoteInfo[]>;
  extractFields: (fields: Record<string, { value: string }>) => Record<string, string>;
  findDuplicateNote: (
    expression: string,
    excludeNoteId: number,
    noteInfo: FieldGroupingNoteInfo,
  ) => Promise<number | null>;
  hasAllConfiguredFields: (
    noteInfo: FieldGroupingNoteInfo,
    configuredFieldNames: (string | undefined)[],
  ) => boolean;
  processNewCard: (
    noteId: number,
    options?: { skipKikuFieldGrouping?: boolean },
  ) => Promise<void>;
  getSentenceCardImageFieldName: () => string | undefined;
  resolveFieldName: (
    availableFieldNames: string[],
    preferredName: string,
  ) => string | null;
  computeFieldGroupingMergedFields: (
    keepNoteId: number,
    deleteNoteId: number,
    keepNoteInfo: FieldGroupingNoteInfo,
    deleteNoteInfo: FieldGroupingNoteInfo,
    includeGeneratedMedia: boolean,
  ) => Promise<Record<string, string>>;
  getNoteFieldMap: (noteInfo: FieldGroupingNoteInfo) => Record<string, string>;
  handleFieldGroupingAuto: (
    originalNoteId: number,
    newNoteId: number,
    newNoteInfo: FieldGroupingNoteInfo,
    expression: string,
  ) => Promise<void>;
  handleFieldGroupingManual: (
    originalNoteId: number,
    newNoteId: number,
    newNoteInfo: FieldGroupingNoteInfo,
    expression: string,
  ) => Promise<boolean>;
}

export class FieldGroupingService {
  constructor(private readonly deps: FieldGroupingDeps) {}

  async triggerFieldGroupingForLastAddedCard(): Promise<void> {
    const sentenceCardConfig = this.deps.getEffectiveSentenceCardConfig();
    if (!sentenceCardConfig.kikuEnabled) {
      this.deps.showOsdNotification("Kiku mode is not enabled");
      return;
    }
    if (sentenceCardConfig.kikuFieldGrouping === "disabled") {
      this.deps.showOsdNotification("Kiku field grouping is disabled");
      return;
    }

    if (this.deps.isUpdateInProgress()) {
      this.deps.showOsdNotification("Anki update already in progress");
      return;
    }

    try {
      await this.deps.withUpdateProgress("Grouping duplicate cards", async () => {
        const deck = this.deps.getDeck ? this.deps.getDeck() : undefined;
        const query = deck ? `"deck:${deck}" added:1` : "added:1";
        const noteIds = await this.deps.findNotes(query);
        if (!noteIds || noteIds.length === 0) {
          this.deps.showOsdNotification("No recently added cards found");
          return;
        }

        const noteId = Math.max(...noteIds);
        const notesInfoResult = await this.deps.notesInfo([noteId]);
        const notesInfo = notesInfoResult as FieldGroupingNoteInfo[];
        if (!notesInfo || notesInfo.length === 0) {
          this.deps.showOsdNotification("Card not found");
          return;
        }
        const noteInfoBeforeUpdate = notesInfo[0];
        const fields = this.deps.extractFields(noteInfoBeforeUpdate.fields);
        const expressionText = fields.expression || fields.word || "";
        if (!expressionText) {
          this.deps.showOsdNotification("No expression/word field found");
          return;
        }

        const duplicateNoteId = await this.deps.findDuplicateNote(
          expressionText,
          noteId,
          noteInfoBeforeUpdate,
        );
        if (duplicateNoteId === null) {
          this.deps.showOsdNotification("No duplicate card found");
          return;
        }

        if (
          !this.deps.hasAllConfiguredFields(noteInfoBeforeUpdate, [
            this.deps.getSentenceCardImageFieldName(),
          ])
        ) {
          await this.deps.processNewCard(noteId, { skipKikuFieldGrouping: true });
        }

        const refreshedInfoResult = await this.deps.notesInfo([noteId]);
        const refreshedInfo = refreshedInfoResult as FieldGroupingNoteInfo[];
        if (!refreshedInfo || refreshedInfo.length === 0) {
          this.deps.showOsdNotification("Card not found");
          return;
        }

        const noteInfo = refreshedInfo[0];

        if (sentenceCardConfig.kikuFieldGrouping === "auto") {
          await this.deps.handleFieldGroupingAuto(
            duplicateNoteId,
            noteId,
            noteInfo,
            expressionText,
          );
          return;
        }
        const handled = await this.deps.handleFieldGroupingManual(
          duplicateNoteId,
          noteId,
          noteInfo,
          expressionText,
        );
        if (!handled) {
          this.deps.showOsdNotification("Field grouping cancelled");
        }
      });
    } catch (error) {
      log.error(
        "Error triggering field grouping:",
        (error as Error).message,
      );
      this.deps.showOsdNotification(
        `Field grouping failed: ${(error as Error).message}`,
      );
    }
  }

  async buildFieldGroupingPreview(
    keepNoteId: number,
    deleteNoteId: number,
    deleteDuplicate: boolean,
  ): Promise<KikuMergePreviewResponse> {
    try {
      const notesInfoResult = await this.deps.notesInfo([
        keepNoteId,
        deleteNoteId,
      ]);
      const notesInfo = notesInfoResult as FieldGroupingNoteInfo[];
      const keepNoteInfo = notesInfo.find((note) => note.noteId === keepNoteId);
      const deleteNoteInfo = notesInfo.find(
        (note) => note.noteId === deleteNoteId,
      );

      if (!keepNoteInfo || !deleteNoteInfo) {
        return { ok: false, error: "Could not load selected notes" };
      }

      const mergedFields = await this.deps.computeFieldGroupingMergedFields(
        keepNoteId,
        deleteNoteId,
        keepNoteInfo,
        deleteNoteInfo,
        false,
      );
      const keepBefore = this.deps.getNoteFieldMap(keepNoteInfo);
      const keepAfter = { ...keepBefore, ...mergedFields };
      const sourceBefore = this.deps.getNoteFieldMap(deleteNoteInfo);

      const compactFields: Record<string, string> = {};
      for (const fieldName of [
        "Sentence",
        "SentenceFurigana",
        "SentenceAudio",
        "Picture",
        "MiscInfo",
      ]) {
        const resolved = this.deps.resolveFieldName(
          Object.keys(keepAfter),
          fieldName,
        );
        if (!resolved) continue;
        compactFields[fieldName] = keepAfter[resolved] || "";
      }

      return {
        ok: true,
        compact: {
          action: {
            keepNoteId,
            deleteNoteId,
            deleteDuplicate,
          },
          mergedFields: compactFields,
        },
        full: {
          keepNote: {
            id: keepNoteId,
            fieldsBefore: keepBefore,
          },
          sourceNote: {
            id: deleteNoteId,
            fieldsBefore: sourceBefore,
          },
          result: {
            fieldsAfter: keepAfter,
            wouldDeleteNoteId: deleteDuplicate ? deleteNoteId : null,
          },
        },
      };
    } catch (error) {
      return {
        ok: false,
        error: `Failed to build preview: ${(error as Error).message}`,
      };
    }
  }
}
