import { AnkiConnectConfig } from '../types';
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
    if (config.fields?.audio && config.fields?.audio.toLowerCase() !== 'expressionaudio') {
      fields.push(config.fields?.audio);
    }
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

    if (!sourceFields['SentenceFurigana'] && sourceFields['Sentence']) {
      sourceFields['SentenceFurigana'] = sourceFields['Sentence'];
    }
    if (!sourceFields['Sentence'] && sourceFields['SentenceFurigana']) {
      sourceFields['Sentence'] = sourceFields['SentenceFurigana'];
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
    if (!sourceFields['SentenceAudio'] && sourceFields['ExpressionAudio']) {
      sourceFields['SentenceAudio'] = sourceFields['ExpressionAudio'];
    }
    if (!sourceFields['ExpressionAudio'] && sourceFields['SentenceAudio']) {
      sourceFields['ExpressionAudio'] = sourceFields['SentenceAudio'];
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

    const sentenceCardConfig = this.deps.getEffectiveSentenceCardConfig();
    const resolvedSentenceAudioField = this.deps.resolveFieldName(
      keepFieldNames,
      sentenceCardConfig.audioField || 'SentenceAudio',
    );
    const resolvedExpressionAudioField = this.deps.resolveFieldName(
      keepFieldNames,
      config.fields?.audio || 'ExpressionAudio',
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
    const groupedSpanRegex = /<span\s+data-group-id="[^"]*">[\s\S]*?<\/span>/gi;
    const ungrouped = value.replace(groupedSpanRegex, '').trim();
    if (ungrouped) return ungrouped;
    return value.trim();
  }

  private extractLastSoundTag(value: string): string {
    const matches = value.match(/\[sound:[^\]]+\]/g);
    if (!matches || matches.length === 0) return '';
    return matches[matches.length - 1]!;
  }

  private extractLastImageTag(value: string): string {
    const matches = value.match(/<img\b[^>]*>/gi);
    if (!matches || matches.length === 0) return '';
    return matches[matches.length - 1]!;
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

    const spanRegex = /<span\s+data-group-id="(\d+)"[^>]*>([\s\S]*?)<\/span>/gi;
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
      const key = entry.content;
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
    if (normalizedField === 'sentenceaudio' || normalizedField === 'expressionaudio') {
      const lastSoundTag = this.extractLastSoundTag(ungrouped);
      if (!lastSoundTag) {
        this.deps.warnFieldParseOnce(fieldName, 'missing-sound-tag');
      }
      return lastSoundTag || ungrouped;
    }

    if (normalizedField === 'picture') {
      const lastImageTag = this.extractLastImageTag(ungrouped);
      if (!lastImageTag) {
        this.deps.warnFieldParseOnce(fieldName, 'missing-image-tag');
      }
      return lastImageTag || ungrouped;
    }

    return ungrouped;
  }

  private getPictureDedupKey(tag: string): string {
    return tag.replace(/\sdata-group-id="[^"]*"/gi, '').trim();
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
        const seen = new Set(mergedTags.map((tag) => this.getPictureDedupKey(tag)));
        for (const entry of sourceEntries) {
          const normalized = this.ensureImageGroupId(entry.tag, entry.groupId);
          const dedupKey = this.getPictureDedupKey(normalized);
          if (seen.has(dedupKey)) continue;
          seen.add(dedupKey);
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
      const seen = new Set(keepEntries.map((entry) => entry.content));
      for (const entry of sourceEntries) {
        const key = entry.content;
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
}
