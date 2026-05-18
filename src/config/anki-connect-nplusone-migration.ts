import {
  getNodeValue,
  parseTree as parseJsoncTree,
  type Node as JsoncNode,
  type ParseError,
} from 'jsonc-parser';
import type { RawConfig } from '../types/config';
import type { ConfigSettingsPatchOperation } from '../types/settings';
import { DEFAULT_CONFIG } from './definitions';
import { applyConfigSettingsPatchToContent } from './settings/jsonc-edit';

export type LegacyAnkiConnectNPlusOneMigrationResult =
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

const LEGACY_N_PLUS_ONE_PATH_MAP = {
  highlightEnabled: 'ankiConnect.knownWords.highlightEnabled',
  refreshMinutes: 'ankiConnect.knownWords.refreshMinutes',
  matchMode: 'ankiConnect.knownWords.matchMode',
  decks: 'ankiConnect.knownWords.decks',
  knownWord: 'subtitleStyle.knownWordColor',
  nPlusOne: 'subtitleStyle.nPlusOneColor',
} as const;

function propertyKey(propertyNode: JsoncNode): string | undefined {
  return propertyNode.children?.[0]?.value;
}

function propertyValue(propertyNode: JsoncNode | undefined): JsoncNode | undefined {
  return propertyNode?.children?.[1];
}

function objectProperties(node: JsoncNode | undefined): JsoncNode[] {
  return node?.type === 'object' ? (node.children ?? []) : [];
}

function findLastProperty(node: JsoncNode | undefined, key: string): JsoncNode | undefined {
  const matches = objectProperties(node).filter((property) => propertyKey(property) === key);
  return matches.at(-1);
}

function findProperties(node: JsoncNode | undefined, key: string): JsoncNode[] {
  return objectProperties(node).filter((property) => propertyKey(property) === key);
}

function findValueAtPath(root: JsoncNode | undefined, path: string): JsoncNode | undefined {
  let node = root;
  for (const segment of path.split('.')) {
    node = propertyValue(findLastProperty(node, segment));
    if (!node) return undefined;
  }
  return node;
}

function hasPath(root: JsoncNode | undefined, path: string): boolean {
  return findValueAtPath(root, path) !== undefined;
}

function normalizeLegacyDecks(value: unknown): unknown {
  if (!Array.isArray(value)) {
    return value;
  }

  const defaultFields = [DEFAULT_CONFIG.ankiConnect.fields.word, 'Word', 'Reading', 'Word Reading'];
  const decks = value
    .filter((entry): entry is string => typeof entry === 'string')
    .map((entry) => entry.trim())
    .filter(Boolean);

  const normalized: Record<string, string[]> = {};
  for (const deck of new Set(decks)) {
    normalized[deck] = defaultFields;
  }
  return normalized;
}

function buildLegacyNPlusOneMigrationOperations(root: JsoncNode | undefined): {
  operations: ConfigSettingsPatchOperation[];
  hasLegacy: boolean;
} {
  const operations: ConfigSettingsPatchOperation[] = [];
  const ankiConnect = propertyValue(findLastProperty(root, 'ankiConnect'));
  const nPlusOneProperties = findProperties(ankiConnect, 'nPlusOne');
  const nPlusOneObjects = nPlusOneProperties.map(propertyValue).filter(Boolean) as JsoncNode[];
  if (nPlusOneObjects.length === 0) {
    return { operations, hasLegacy: false };
  }

  const canonicalNPlusOneValues = new Map<string, unknown>();
  const legacyValues = new Map<keyof typeof LEGACY_N_PLUS_ONE_PATH_MAP, unknown>();
  let hasLegacy = false;

  for (const nPlusOne of nPlusOneObjects) {
    for (const property of objectProperties(nPlusOne)) {
      const key = propertyKey(property);
      if (!key) continue;
      const valueNode = propertyValue(property);
      const value = valueNode ? getNodeValue(valueNode) : undefined;
      if (key === 'enabled' || key === 'minSentenceWords') {
        canonicalNPlusOneValues.set(key, value);
        continue;
      }
      if (key in LEGACY_N_PLUS_ONE_PATH_MAP) {
        hasLegacy = true;
        legacyValues.set(key as keyof typeof LEGACY_N_PLUS_ONE_PATH_MAP, value);
      }
    }
  }

  if (nPlusOneObjects.length > 1) {
    for (const [key, value] of canonicalNPlusOneValues) {
      operations.push({
        op: 'set',
        path: `ankiConnect.nPlusOne.${key}`,
        value,
      });
    }
  }

  for (const [key, path] of Object.entries(LEGACY_N_PLUS_ONE_PATH_MAP) as Array<
    [keyof typeof LEGACY_N_PLUS_ONE_PATH_MAP, string]
  >) {
    if (!legacyValues.has(key)) continue;
    if (!hasPath(root, path)) {
      const value =
        key === 'decks' ? normalizeLegacyDecks(legacyValues.get(key)) : legacyValues.get(key);
      operations.push({
        op: 'set',
        path,
        value,
      });
    }
    operations.push({
      op: 'reset',
      path: `ankiConnect.nPlusOne.${key}`,
    });
  }

  return { operations, hasLegacy };
}

export function applyLegacyAnkiConnectNPlusOneMigrationToContent(options: {
  content: string;
  rawConfig: RawConfig;
}): LegacyAnkiConnectNPlusOneMigrationResult {
  const errors: ParseError[] = [];
  const root = parseJsoncTree(options.content || '{}', errors, {
    allowTrailingComma: true,
    disallowComments: false,
  });
  if (!root || errors.length > 0) {
    return {
      migrated: false,
      content: options.content,
      rawConfig: options.rawConfig,
    };
  }

  const { operations, hasLegacy } = buildLegacyNPlusOneMigrationOperations(root);
  if (operations.length === 0 && !hasLegacy) {
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
