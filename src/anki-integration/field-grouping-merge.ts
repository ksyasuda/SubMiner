import { AnkiConnectConfig } from '../types/anki';
import { getConfiguredWordFieldName } from '../anki-field-config';

interface FieldGroupingMergeMedia {
  audioField?: string;
  audioValue?: string;
  imageField?: string;
  imageValue?: string;
  miscInfoValue?: string;
}

export interface FieldGroupingMergeNoteInfo {
  noteId: number;
  fields: Record<string, { value: string }>;
}

interface FieldGroupingMergeDeps {
  getConfig: () => AnkiConnectConfig;
  getEffectiveSentenceCardConfig: () => {
    sentenceField: string;
    audioField: string;
    fieldGroupingProvider: 'kiku' | 'senren' | null;
  };
  getCurrentSubtitleText: () => string | undefined;
  resolveFieldName: (availableFieldNames: string[], preferredName: string) => string | null;
  resolveNoteFieldName: (
    noteInfo: FieldGroupingMergeNoteInfo,
    preferredName?: string,
  ) => string | null;
  extractFields: (fields: Record<string, { value: string }>) => Record<string, string>;
  processSentence: (mpvSentence: string, noteFields: Record<string, string>) => string;
  generateMediaForMerge: (noteInfo: FieldGroupingMergeNoteInfo) => Promise<FieldGroupingMergeMedia>;
  warnFieldParseOnce: (fieldName: string, reason: string, detail?: string) => void;
}

export class FieldGroupingMergeCollaborator {
  private readonly strictGroupingFieldDefaults = new Set<string>([
    'picture',
    'sentence',
    'sentenceaudio',
    'sentencefurigana',
    'miscinfo',
  ]);

  constructor(private readonly deps: FieldGroupingMergeDeps) {}

  getGroupableFieldNames(): string[] {
    const config = this.deps.getConfig();
    const fields: string[] = [];
    fields.push('Sentence');
    fields.push('SentenceAudio');
    fields.push('Picture');
    if (config.fields?.image) fields.push(config.fields?.image);
    if (config.fields?.sentence) fields.push(config.fields?.sentence);
    const sentenceCardConfig = this.deps.getEffectiveSentenceCardConfig();
    const sentenceAudioField = sentenceCardConfig.audioField;
    if (!fields.includes(sentenceAudioField)) fields.push(sentenceAudioField);
    if (config.fields?.miscInfo) fields.push(config.fields?.miscInfo);
    fields.push('SentenceFurigana');
    return fields;
  }

  getNoteFieldMap(noteInfo: FieldGroupingMergeNoteInfo): Record<string, string> {
    const fields: Record<string, string> = {};
    for (const [name, field] of Object.entries(noteInfo.fields)) {
      fields[name] = field?.value || '';
    }
    return fields;
  }

  async computeFieldGroupingMergedFields(
    keepNoteId: number,
    deleteNoteId: number,
    keepNoteInfo: FieldGroupingMergeNoteInfo,
    deleteNoteInfo: FieldGroupingMergeNoteInfo,
    includeGeneratedMedia: boolean,
  ): Promise<Record<string, string>> {
    const config = this.deps.getConfig();
    const configuredWordField = getConfiguredWordFieldName(config);
    const groupableFields = this.getGroupableFieldNames();
    const keepFieldNames = Object.keys(keepNoteInfo.fields);
    const sentenceCardConfig = this.deps.getEffectiveSentenceCardConfig();
    const senrenSourceSceneOffset =
      sentenceCardConfig.fieldGroupingProvider === 'senren'
        ? this.countSenrenAudioScenes(
            this.getResolvedFieldValue(keepNoteInfo, sentenceCardConfig.audioField),
          )
        : 0;
    const sourceFields: Record<string, string> = {};
    const resolvedKeepFieldByPreferred = new Map<string, string>();
    for (const preferredFieldName of groupableFields) {
      sourceFields[preferredFieldName] = this.getResolvedFieldValue(
        deleteNoteInfo,
        preferredFieldName,
      );
      const keepResolved = this.deps.resolveFieldName(keepFieldNames, preferredFieldName);
      if (keepResolved) {
        resolvedKeepFieldByPreferred.set(preferredFieldName, keepResolved);
      }
    }

    if (!sourceFields[configuredWordField] && sourceFields['Expression']) {
      sourceFields[configuredWordField] = sourceFields['Expression'];
    }
    if (!sourceFields[configuredWordField] && sourceFields['Word']) {
      sourceFields[configuredWordField] = sourceFields['Word'];
    }
    if (!sourceFields['Expression'] && sourceFields[configuredWordField]) {
      sourceFields['Expression'] = sourceFields[configuredWordField];
    }
    if (!sourceFields['Word'] && sourceFields[configuredWordField]) {
      sourceFields['Word'] = sourceFields[configuredWordField];
    }
    if (
      config.fields?.sentence &&
      !sourceFields[config.fields?.sentence] &&
      this.deps.getCurrentSubtitleText()
    ) {
      const deleteFields = this.deps.extractFields(deleteNoteInfo.fields);
      sourceFields[config.fields?.sentence] = this.deps.processSentence(
        this.deps.getCurrentSubtitleText()!,
        deleteFields,
      );
    }

    if (includeGeneratedMedia) {
      const media = await this.deps.generateMediaForMerge(keepNoteInfo);
      if (media.audioField && media.audioValue && !sourceFields[media.audioField]) {
        sourceFields[media.audioField] = media.audioValue;
      }
      if (media.imageField && media.imageValue && !sourceFields[media.imageField]) {
        sourceFields[media.imageField] = media.imageValue;
      }
      if (
        config.fields?.miscInfo &&
        media.miscInfoValue &&
        !sourceFields[config.fields?.miscInfo]
      ) {
        sourceFields[config.fields?.miscInfo] = media.miscInfoValue;
      }
    }

    const mergedFields: Record<string, string> = {};
    for (const preferredFieldName of groupableFields) {
      const keepFieldName = resolvedKeepFieldByPreferred.get(preferredFieldName);
      if (!keepFieldName) continue;

      const keepFieldNormalized = keepFieldName.toLowerCase();
      if (
        keepFieldNormalized === 'expression' ||
        keepFieldNormalized === configuredWordField.toLowerCase() ||
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

      if (keepFieldNormalized === 'sentencefurigana') {
        const hasBothValues = existingValue.trim().length > 0 && newValue.trim().length > 0;
        const usesSenrenGrouping =
          this.deps.getEffectiveSentenceCardConfig().fieldGroupingProvider === 'senren';
        mergedFields[keepFieldName] =
          hasBothValues || usesSenrenGrouping
            ? this.applyFieldGrouping(
                existingValue,
                newValue,
                keepNoteId,
                deleteNoteId,
                keepFieldName,
                senrenSourceSceneOffset,
              )
            : '';
        continue;
      }

      if (isStrictField) {
        mergedFields[keepFieldName] = this.applyFieldGrouping(
          existingValue,
          newValue,
          keepNoteId,
          deleteNoteId,
          keepFieldName,
          senrenSourceSceneOffset,
        );
      } else if (existingValue.trim() && newValue.trim()) {
        mergedFields[keepFieldName] = this.applyFieldGrouping(
          existingValue,
          newValue,
          keepNoteId,
          deleteNoteId,
          keepFieldName,
          senrenSourceSceneOffset,
        );
      } else {
        if (!newValue.trim()) continue;
        mergedFields[keepFieldName] = newValue;
      }
    }

    return mergedFields;
  }

  private getResolvedFieldValue(
    noteInfo: FieldGroupingMergeNoteInfo,
    preferredFieldName?: string,
  ): string {
    if (!preferredFieldName) return '';
    const resolved = this.deps.resolveNoteFieldName(noteInfo, preferredFieldName);
    if (!resolved) return '';
    return noteInfo.fields[resolved]?.value || '';
  }

  private extractUngroupedValue(value: string): string {
    const ungrouped = this.extractUngroupedRemainder(value);
    if (ungrouped) return ungrouped;
    return value.trim();
  }

  private extractUngroupedRemainder(value: string): string {
    const groupedSpanRegex = /<span\b[^>]*data-group-id="[^"]*"[^>]*>[\s\S]*?<\/span>/gi;
    return value.replace(groupedSpanRegex, '').trim();
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
        this.deps.warnFieldParseOnce(fieldName, 'invalid-group-id', rawId);
      }
    }

    const spanRegex = /<span\b[^>]*data-group-id="(\d+)"[^>]*>([\s\S]*?)<\/span>/gi;
    let match;
    while ((match = spanRegex.exec(value)) !== null) {
      const groupId = Number(match[1]);
      if (!Number.isFinite(groupId) || groupId <= 0) continue;
      const content = this.normalizeStrictGroupedValue(match[2] || '', fieldName);
      if (!content) {
        this.deps.warnFieldParseOnce(fieldName, 'empty-group-content');
        continue;
      }
      entries.push({ groupId, content });
    }
    if (entries.length === 0 && /<span\b/i.test(value)) {
      this.deps.warnFieldParseOnce(fieldName, 'no-usable-span-entries');
    }
    return entries;
  }

  private parseStrictEntries(
    value: string,
    fallbackGroupId: number,
    fieldName: string,
  ): { groupId: number; content: string }[] {
    const entries = this.extractSpanEntries(value, fieldName);
    const ungroupedSource =
      entries.length > 0
        ? this.extractUngroupedRemainder(value)
        : this.extractUngroupedValue(value);
    const ungrouped = this.normalizeStrictGroupedValue(ungroupedSource, fieldName);
    if (ungrouped) {
      entries.push({ groupId: fallbackGroupId, content: ungrouped });
    }

    return entries;
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
          this.deps.warnFieldParseOnce('Picture', 'invalid-group-id', idMatch[1]);
        } else {
          groupId = parsed;
        }
      }
      const normalizedTag = this.ensureImageGroupId(tag, groupId);
      if (!normalizedTag) {
        this.deps.warnFieldParseOnce('Picture', 'empty-image-tag');
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
    if (normalizedField === 'sentenceaudio' && !/\[sound:[^\]]+\]/.test(ungrouped)) {
      this.deps.warnFieldParseOnce(fieldName, 'missing-sound-tag');
    }

    return ungrouped;
  }

  private getStrictSpanGroupingFields(): Set<string> {
    const strictFields = new Set(this.strictGroupingFieldDefaults);
    const sentenceCardConfig = this.deps.getEffectiveSentenceCardConfig();
    strictFields.add((sentenceCardConfig.sentenceField || 'sentence').toLowerCase());
    strictFields.add((sentenceCardConfig.audioField || 'sentenceaudio').toLowerCase());
    const config = this.deps.getConfig();
    if (config.fields?.image) strictFields.add(config.fields.image.toLowerCase());
    if (config.fields?.miscInfo) strictFields.add(config.fields.miscInfo.toLowerCase());
    return strictFields;
  }

  private shouldUseStrictSpanGrouping(fieldName: string): boolean {
    const normalized = fieldName.toLowerCase();
    return this.getStrictSpanGroupingFields().has(normalized);
  }

  private isPictureField(fieldName: string): boolean {
    const normalized = fieldName.toLowerCase();
    const configuredImageField = this.deps.getConfig().fields?.image?.toLowerCase();
    return normalized === 'picture' || normalized === configuredImageField;
  }

  private sortEntriesByGroupIdDescending<T extends { groupId: number }>(entries: T[]): T[] {
    return [...entries].sort((a, b) => b.groupId - a.groupId);
  }

  private isSentenceAudioField(fieldName: string): boolean {
    const normalized = fieldName.toLowerCase();
    const audioField = (
      this.deps.getEffectiveSentenceCardConfig().audioField || 'sentenceaudio'
    ).toLowerCase();
    return normalized === 'sentenceaudio' || normalized === audioField;
  }

  private isSenrenGroupOpenTag(openTag: string): boolean {
    const classMatch =
      openTag.match(/class\s*=\s*"([^"]*)"/i) || openTag.match(/class\s*=\s*'([^']*)'/i);
    if (!classMatch) return false;
    // Senren's templates match class tokens case-sensitively (/^group\d*$/).
    return classMatch[1]!.split(/\s+/).some((token) => /^group\d*$/.test(token));
  }

  private countSenrenAudioScenes(value: string): number {
    const soundEntries = value.match(/\[sound:[^\]]+\]/g)?.length ?? 0;
    if (soundEntries > 0) return soundEntries;
    return this.parseSenrenSceneEntries(value).length;
  }

  private rebaseSenrenNumberedGroup(entry: string, sceneOffset: number): string {
    if (sceneOffset <= 0) return entry;

    return entry.replace(
      /^(\s*<span\b[^>]*?\bclass\s*=\s*)(["'])([^"']*)\2/i,
      (_match: string, prefix: string, quote: string, rawClasses: string) => {
        const classes = rawClasses
          .split(/(\s+)/)
          .map((classToken) => {
            const groupMatch = classToken.match(/^group(\d+)$/);
            if (!groupMatch) return classToken;
            const targetScene = Number(groupMatch[1]);
            if (!Number.isSafeInteger(targetScene) || targetScene <= 0) return classToken;
            return `group${targetScene + sceneOffset}`;
          })
          .join('');
        return `${prefix}${quote}${classes}${quote}`;
      },
    );
  }

  /**
   * Splits a Senren field into ordered scene entries. Top-level
   * `<span class="group">`/`"groupN"` spans are kept verbatim (nested markup like
   * `<span class="highlight">` included); ungrouped runs are wrapped in a group
   * span at their original position, because Senren discards anything outside a
   * group span once scene switching activates.
   */
  private parseSenrenSceneEntries(value: string): string[] {
    const tokenRegex = /<span\b[^>]*>|<\/span>/gi;
    const entries: string[] = [];
    const pushUngrouped = (raw: string): void => {
      const text = raw.replace(/<br\s*\/?>/gi, ' ').trim();
      if (text) entries.push(`<span class="group">${text}</span>`);
    };
    let cursor = 0;
    let depth = 0;
    let entryStart = -1;
    let match;
    while ((match = tokenRegex.exec(value)) !== null) {
      const token = match[0]!;
      if (token[1] !== '/') {
        if (depth === 0 && this.isSenrenGroupOpenTag(token)) {
          pushUngrouped(value.slice(cursor, match.index));
          entryStart = match.index;
          cursor = match.index;
        }
        depth += 1;
      } else {
        depth = Math.max(0, depth - 1);
        if (depth === 0 && entryStart !== -1) {
          const end = match.index + token.length;
          entries.push(value.slice(entryStart, end));
          entryStart = -1;
          cursor = end;
        }
      }
    }
    if (entryStart !== -1) {
      // Unclosed group span: close every span still open (the group and any nested
      // markup) so the following scenes are siblings rather than nested inside it.
      entries.push(`${value.slice(entryStart)}${'</span>'.repeat(depth)}`);
    } else {
      pushUngrouped(value.slice(cursor));
    }
    return entries;
  }

  /**
   * Merges two notes' field values in Senren's scene-switching format. Scenes are
   * appended in order (existing first, never resorted) so indices stay aligned
   * across sentence/picture/miscInfo with the sentenceAudio entries, which alone
   * drive Senren's scene count.
   */
  private applySenrenFieldGrouping(
    existingValue: string,
    newValue: string,
    fieldName: string,
    sourceSceneOffset: number,
  ): string {
    if (this.isPictureField(fieldName)) {
      const tags = [...this.extractImageTags(existingValue), ...this.extractImageTags(newValue)];
      if (tags.length === 0) return existingValue || newValue;
      return tags.join('');
    }

    if (this.isSentenceAudioField(fieldName)) {
      const existing = existingValue.trim();
      const added = newValue.trim();
      if (!existing || !added) return existing || added;
      if (!/\[sound:[^\]]+\]/.test(added)) {
        this.deps.warnFieldParseOnce(fieldName, 'missing-sound-tag');
      }
      return existing + added;
    }

    const sourceEntries = this.parseSenrenSceneEntries(newValue).map((entry) =>
      this.rebaseSenrenNumberedGroup(entry, sourceSceneOffset),
    );
    const merged = [...this.parseSenrenSceneEntries(existingValue), ...sourceEntries];
    if (merged.length === 0) return existingValue || newValue;
    return merged.join('');
  }

  private applyFieldGrouping(
    existingValue: string,
    newValue: string,
    keepGroupId: number,
    sourceGroupId: number,
    fieldName: string,
    senrenSourceSceneOffset: number,
  ): string {
    if (this.deps.getEffectiveSentenceCardConfig().fieldGroupingProvider === 'senren') {
      return this.applySenrenFieldGrouping(
        existingValue,
        newValue,
        fieldName,
        senrenSourceSceneOffset,
      );
    }

    if (this.shouldUseStrictSpanGrouping(fieldName)) {
      if (this.isPictureField(fieldName)) {
        const keepEntries = this.parsePictureEntries(existingValue, keepGroupId);
        const sourceEntries = this.parsePictureEntries(newValue, sourceGroupId);
        if (keepEntries.length === 0 && sourceEntries.length === 0) {
          return existingValue || newValue;
        }
        return this.sortEntriesByGroupIdDescending([...keepEntries, ...sourceEntries])
          .map((entry) => entry.tag)
          .join('');
      }

      const keepEntries = this.parseStrictEntries(existingValue, keepGroupId, fieldName);
      const sourceEntries = this.parseStrictEntries(newValue, sourceGroupId, fieldName);
      if (keepEntries.length === 0 && sourceEntries.length === 0) {
        return existingValue || newValue;
      }
      const merged = this.sortEntriesByGroupIdDescending([...keepEntries, ...sourceEntries]);
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
}
