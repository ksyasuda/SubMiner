import type { RawConfig } from '../types/config';
import type { ConfigSettingsPatchOperation } from '../types/settings';
import {
  buildSubtitleCssDeclarationObject,
  getSubtitleCssManagedConfigPaths,
  getSubtitleCssPath,
  type SubtitleCssScope,
} from '../settings/subtitle-style-css';
import { applyConfigSettingsPatchToContent } from './settings/jsonc-edit';

const SUBTITLE_CSS_SCOPES: SubtitleCssScope[] = ['primary', 'secondary', 'sidebar'];
const HEX_COLOR_PATTERN = /^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/;

export type LegacySubtitleStyleCssMigrationResult =
  | {
      migrated: true;
      content: string;
      rawConfig: RawConfig;
    }
  | {
      migrated: false;
      content: string;
      rawConfig: RawConfig;
    };

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function getValueAtPath(root: unknown, path: string): unknown {
  let current = root;
  for (const segment of path.split('.')) {
    if (!isRecord(current)) return undefined;
    current = current[segment];
  }
  return current;
}

function hasPath(root: unknown, path: string): boolean {
  let current = root;
  const segments = path.split('.');
  for (const [index, segment] of segments.entries()) {
    if (!isRecord(current) || !Object.hasOwn(current, segment)) {
      return false;
    }
    if (index === segments.length - 1) {
      return true;
    }
    current = current[segment];
  }
  return false;
}

function isMigratableLegacySubtitleCssValue(path: string, value: unknown): boolean {
  if (path === 'subtitleStyle.hoverTokenColor') {
    return typeof value === 'string' && HEX_COLOR_PATTERN.test(value.trim());
  }
  if (path === 'subtitleStyle.hoverTokenBackgroundColor') {
    return typeof value === 'string';
  }
  return true;
}

export function buildLegacySubtitleStyleCssMigrationOperations(
  rawConfig: RawConfig,
): ConfigSettingsPatchOperation[] {
  const operations: ConfigSettingsPatchOperation[] = [];

  for (const scope of SUBTITLE_CSS_SCOPES) {
    const cssPath = getSubtitleCssPath(scope);
    const values: Record<string, unknown> = {
      [cssPath]: getValueAtPath(rawConfig, cssPath),
    };
    const legacyPaths = getSubtitleCssManagedConfigPaths(scope).filter(
      (legacyPath) =>
        hasPath(rawConfig, legacyPath) &&
        isMigratableLegacySubtitleCssValue(legacyPath, getValueAtPath(rawConfig, legacyPath)),
    );
    if (legacyPaths.length === 0) continue;

    for (const legacyPath of legacyPaths) {
      values[legacyPath] = getValueAtPath(rawConfig, legacyPath);
    }

    operations.push({
      op: 'set',
      path: cssPath,
      value: buildSubtitleCssDeclarationObject(scope, values),
    });
    for (const legacyPath of legacyPaths) {
      operations.push({ op: 'reset', path: legacyPath });
    }
  }

  return operations;
}

export function applyLegacySubtitleStyleCssMigrationToContent(options: {
  content: string;
  rawConfig: RawConfig;
}): LegacySubtitleStyleCssMigrationResult {
  const operations = buildLegacySubtitleStyleCssMigrationOperations(options.rawConfig);
  if (operations.length === 0) {
    return {
      migrated: false,
      content: options.content,
      rawConfig: options.rawConfig,
    };
  }

  const result = applyConfigSettingsPatchToContent({
    content: options.content,
    operations,
    previousWarnings: [],
  });
  if (!result.ok) {
    return {
      migrated: false,
      content: options.content,
      rawConfig: options.rawConfig,
    };
  }

  return {
    migrated: true,
    content: result.content,
    rawConfig: result.rawConfig,
  };
}
