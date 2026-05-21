import {
  applyEdits,
  modify,
  parse as parseJsonc,
  parseTree as parseJsoncTree,
  type Edit,
  type Node as JsoncNode,
  type FormattingOptions,
  type ParseError,
} from 'jsonc-parser';
import type { ConfigValidationWarning, RawConfig, ResolvedConfig } from '../../types/config';
import type {
  ConfigSettingsField,
  ConfigSettingsPatchOperation,
  ConfigSettingsSnapshot,
} from '../../types/settings';
import { resolveConfig } from '../resolve';
import { getConfigValueAtPath, SECRET_PATHS } from './registry';

const JSONC_FORMATTING_OPTIONS: FormattingOptions = {
  insertSpaces: true,
  tabSize: 2,
  eol: '\n',
};

export type ConfigSettingsPatchApplyResult =
  | {
      ok: true;
      content: string;
      rawConfig: RawConfig;
      resolvedConfig: ResolvedConfig;
      warnings: ConfigValidationWarning[];
    }
  | {
      ok: false;
      content: string;
      warnings: ConfigValidationWarning[];
      error: string;
    };

interface ApplyConfigSettingsPatchOptions {
  content: string;
  operations: ConfigSettingsPatchOperation[];
  previousWarnings: ConfigValidationWarning[];
}

interface BuildConfigSettingsSnapshotOptions {
  configPath: string;
  rawConfig: RawConfig;
  resolvedConfig: ResolvedConfig;
  warnings: ConfigValidationWarning[];
  fields: ConfigSettingsField[];
}

function pathToSegments(path: string): string[] {
  return path.split('.').filter(Boolean);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function pathStartsWith(path: string, prefix: string): boolean {
  return path === prefix || path.startsWith(`${prefix}.`);
}

function warningBelongsToModifiedPath(
  warning: ConfigValidationWarning,
  operation: ConfigSettingsPatchOperation,
): boolean {
  return (
    pathStartsWith(warning.path, operation.path) || pathStartsWith(operation.path, warning.path)
  );
}

function warningIdentity(warning: ConfigValidationWarning): string {
  return `${warning.path}\n${JSON.stringify(warning.value)}\n${warning.message}`;
}

function parseRawConfig(content: string): RawConfig {
  const errors: ParseError[] = [];
  const parsed = parseJsonc(content || '{}', errors, {
    allowTrailingComma: true,
    disallowComments: false,
  });
  if (errors.length > 0) {
    throw new Error(`Invalid JSONC (${errors[0]?.error ?? 'unknown'})`);
  }
  return isRecord(parsed) ? (parsed as RawConfig) : {};
}

function normalizeContent(content: string): string {
  return content.trim().length > 0 ? content : '{}\n';
}

function applySingleOperation(content: string, operation: ConfigSettingsPatchOperation): string {
  content = removeDuplicatePropertiesAlongPath(content, operation.path);
  const edits = modify(
    content,
    pathToSegments(operation.path),
    operation.op === 'reset' ? undefined : operation.value,
    {
      formattingOptions: JSONC_FORMATTING_OPTIONS,
      getInsertionIndex: (properties) => properties.length,
    },
  );
  return applyEdits(content, edits);
}

function propertyKey(propertyNode: JsoncNode): string | undefined {
  return propertyNode.children?.[0]?.value;
}

function propertyValue(propertyNode: JsoncNode): JsoncNode | undefined {
  return propertyNode.children?.[1];
}

function objectProperties(node: JsoncNode | undefined): JsoncNode[] {
  return node?.type === 'object' ? (node.children ?? []) : [];
}

function isWhitespace(value: string | undefined): boolean {
  return value === ' ' || value === '\t' || value === '\r' || value === '\n';
}

function nextNonWhitespaceOffset(content: string, offset: number): number {
  let index = offset;
  while (index < content.length) {
    if (isWhitespace(content[index])) {
      index += 1;
      continue;
    }
    if (content[index] === '/' && content[index + 1] === '/') {
      index += 2;
      while (index < content.length && content[index] !== '\n') index += 1;
      continue;
    }
    if (content[index] === '/' && content[index + 1] === '*') {
      index += 2;
      while (
        index + 1 < content.length &&
        !(content[index] === '*' && content[index + 1] === '/')
      ) {
        index += 1;
      }
      index = Math.min(content.length, index + 2);
      continue;
    }
    break;
  }
  return index;
}

function previousNonWhitespaceOffset(content: string, offset: number): number {
  let index = offset;
  while (index >= 0) {
    if (isWhitespace(content[index])) {
      index -= 1;
      continue;
    }
    const lineStart = content.lastIndexOf('\n', index) + 1;
    const linePrefix = content.slice(lineStart, index + 1);
    const lineCommentStart = linePrefix.lastIndexOf('//');
    if (lineCommentStart >= 0 && /^[ \t]*$/.test(linePrefix.slice(0, lineCommentStart))) {
      index = lineStart - 1;
      continue;
    }
    if (content[index] === '/' && content[index - 1] === '*') {
      index -= 2;
      while (index > 0 && !(content[index - 1] === '/' && content[index] === '*')) {
        index -= 1;
      }
      index -= 2;
      continue;
    }
    break;
  }
  return index;
}

function lineStartOffset(content: string, offset: number): number {
  return content.lastIndexOf('\n', Math.max(0, offset - 1)) + 1;
}

function removalEditForProperty(content: string, propertyNode: JsoncNode): Edit {
  let offset = propertyNode.offset;
  let end = propertyNode.offset + propertyNode.length;
  const next = nextNonWhitespaceOffset(content, end);

  if (content[next] === ',') {
    end = next + 1;
    const lineStart = lineStartOffset(content, offset);
    if (/^[ \t]*$/.test(content.slice(lineStart, offset))) {
      offset = lineStart;
    }
  } else {
    const previous = previousNonWhitespaceOffset(content, offset - 1);
    if (content[previous] === ',') {
      offset = previous;
    }
  }

  return {
    offset,
    length: Math.max(0, end - offset),
    content: '',
  };
}

function collectDuplicatePropertyRemovalEdits(content: string, path: string): Edit[] {
  const errors: ParseError[] = [];
  let node = parseJsoncTree(content, errors, {
    allowTrailingComma: true,
    disallowComments: false,
  });
  if (!node || errors.length > 0) {
    return [];
  }

  const edits: Edit[] = [];
  for (const segment of pathToSegments(path)) {
    const matches = objectProperties(node).filter((property) => propertyKey(property) === segment);
    if (matches.length === 0) {
      break;
    }

    for (const duplicate of matches.slice(0, -1)) {
      edits.push(removalEditForProperty(content, duplicate));
    }

    node = propertyValue(matches[matches.length - 1]!);
  }

  return edits;
}

function applyRemovalEdits(content: string, edits: Edit[]): string {
  return [...edits]
    .sort((left, right) => right.offset - left.offset)
    .reduce(
      (current, edit) =>
        `${current.slice(0, edit.offset)}${edit.content}${current.slice(edit.offset + edit.length)}`,
      content,
    );
}

function removeDuplicatePropertiesAlongPath(content: string, path: string): string {
  const edits = collectDuplicatePropertyRemovalEdits(content, path);
  return edits.length > 0 ? applyRemovalEdits(content, edits) : content;
}

function collectModifiedWarnings(
  warnings: ConfigValidationWarning[],
  operations: ConfigSettingsPatchOperation[],
  previousWarnings: ConfigValidationWarning[],
): ConfigValidationWarning[] {
  const previous = new Set(previousWarnings.map(warningIdentity));
  return warnings.filter((warning) => {
    if (!operations.some((operation) => warningBelongsToModifiedPath(warning, operation))) {
      return false;
    }
    return !previous.has(warningIdentity(warning));
  });
}

export function applyConfigSettingsPatchToContent(
  options: ApplyConfigSettingsPatchOptions,
): ConfigSettingsPatchApplyResult {
  let content = normalizeContent(options.content);

  try {
    parseRawConfig(content);
  } catch (error) {
    return {
      ok: false,
      content,
      warnings: [],
      error: error instanceof Error ? error.message : 'Invalid JSONC.',
    };
  }

  try {
    for (const operation of options.operations) {
      content = applySingleOperation(content, operation);
    }

    const rawConfig = parseRawConfig(content);
    const { resolved, warnings } = resolveConfig(rawConfig);
    const modifiedWarnings = collectModifiedWarnings(
      warnings,
      options.operations,
      options.previousWarnings,
    );
    if (modifiedWarnings.length > 0) {
      return {
        ok: false,
        content,
        warnings: modifiedWarnings,
        error: 'One or more modified settings failed validation.',
      };
    }

    return {
      ok: true,
      content,
      rawConfig,
      resolvedConfig: resolved,
      warnings,
    };
  } catch (error) {
    return {
      ok: false,
      content,
      warnings: [],
      error: error instanceof Error ? error.message : 'Failed to update config content.',
    };
  }
}

export function buildConfigSettingsSnapshot(
  options: BuildConfigSettingsSnapshotOptions,
): ConfigSettingsSnapshot {
  const values: Record<string, unknown> = {};

  for (const field of options.fields) {
    const rawValue = getConfigValueAtPath(options.rawConfig, field.configPath);
    const resolvedValue = getConfigValueAtPath(options.resolvedConfig, field.configPath);
    if (field.secret) {
      values[field.configPath] = {
        configured:
          (typeof rawValue === 'string' && rawValue.length > 0) ||
          (typeof resolvedValue === 'string' && resolvedValue.length > 0),
      };
      continue;
    }

    values[field.configPath] = structuredClone(rawValue != null ? rawValue : resolvedValue);
  }

  for (const secretPath of SECRET_PATHS) {
    if (Object.hasOwn(values, secretPath)) {
      continue;
    }
    const rawValue = getConfigValueAtPath(options.rawConfig, secretPath);
    const resolvedValue = getConfigValueAtPath(options.resolvedConfig, secretPath);
    if (
      (typeof rawValue === 'string' && rawValue.length > 0) ||
      (typeof resolvedValue === 'string' && resolvedValue.length > 0)
    ) {
      values[secretPath] = { configured: true };
    }
  }

  return {
    configPath: options.configPath,
    fields: options.fields,
    values,
    warnings: options.warnings,
  };
}
