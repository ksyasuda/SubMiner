import { ResolvedConfig } from '../types/config';
import {
  CONFIG_OPTION_REGISTRY,
  CONFIG_TEMPLATE_SECTIONS,
  DEFAULT_CONFIG,
  DEFAULT_KEYBINDINGS,
  deepCloneConfig,
} from './definitions';
import {
  buildSubtitleCssDeclarationObject,
  getSubtitleCssManagedConfigPaths,
  getSubtitleCssPath,
  type SubtitleCssScope,
} from '../settings/subtitle-style-css';

const OPTION_REGISTRY_BY_PATH = new Map(CONFIG_OPTION_REGISTRY.map((entry) => [entry.path, entry]));
const TOP_LEVEL_SECTION_DESCRIPTION_BY_KEY = new Map(
  CONFIG_TEMPLATE_SECTIONS.map((section) => [String(section.key), section.description[0] ?? '']),
);
const SUBTITLE_CSS_SCOPES: SubtitleCssScope[] = ['primary', 'secondary', 'sidebar'];

function normalizeCommentText(value: string): string {
  return value.replace(/\s+/g, ' ').replace(/\*\//g, '*\\/').trim();
}

function humanizeKey(key: string): string {
  const spaced = key
    .replace(/^--/, '')
    .replace(/_/g, ' ')
    .replace(/-/g, ' ')
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .toLowerCase();
  return spaced.charAt(0).toUpperCase() + spaced.slice(1);
}

function buildInlineOptionComment(path: string, value: unknown): string {
  const registryEntry = OPTION_REGISTRY_BY_PATH.get(path);
  const baseDescription =
    registryEntry?.description ?? TOP_LEVEL_SECTION_DESCRIPTION_BY_KEY.get(path);
  const description =
    baseDescription && baseDescription.trim().length > 0
      ? normalizeCommentText(baseDescription)
      : `${humanizeKey(path.split('.').at(-1) ?? path)} setting.`;

  if (registryEntry?.enumValues?.length) {
    return `${description} Values: ${registryEntry.enumValues.join(' | ')}`;
  }
  if (typeof value === 'boolean') {
    return `${description} Values: true | false`;
  }
  return description;
}

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

function setValueAtPath(root: unknown, path: string, value: unknown): void {
  const segments = path.split('.').filter(Boolean);
  let current = root;
  for (const [index, segment] of segments.entries()) {
    if (!isRecord(current)) return;
    if (index === segments.length - 1) {
      current[segment] = value;
      return;
    }
    current = current[segment];
  }
}

function deleteValueAtPath(root: unknown, path: string): void {
  const segments = path.split('.').filter(Boolean);
  let current = root;
  for (const [index, segment] of segments.entries()) {
    if (!isRecord(current)) return;
    if (index === segments.length - 1) {
      delete current[segment];
      return;
    }
    current = current[segment];
  }
}

function foldSubtitleCssManagedDefaults(templateConfig: ResolvedConfig): void {
  for (const scope of SUBTITLE_CSS_SCOPES) {
    const cssPath = getSubtitleCssPath(scope);
    const values: Record<string, unknown> = {
      [cssPath]: getValueAtPath(templateConfig, cssPath),
    };
    const managedPaths = getSubtitleCssManagedConfigPaths(scope);
    for (const managedPath of managedPaths) {
      values[managedPath] = getValueAtPath(templateConfig, managedPath);
    }
    setValueAtPath(templateConfig, cssPath, buildSubtitleCssDeclarationObject(scope, values));
    for (const managedPath of managedPaths) {
      deleteValueAtPath(templateConfig, managedPath);
    }
  }
}

function renderValue(value: unknown, indent = 0, path = ''): string {
  const pad = ' '.repeat(indent);
  const nextPad = ' '.repeat(indent + 2);

  if (value === null) return 'null';
  if (typeof value === 'string') return JSON.stringify(value);
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);

  if (Array.isArray(value)) {
    if (value.length === 0) return '[]';
    const items = value.map((item) => `${nextPad}${renderValue(item, indent + 2, `${path}[]`)}`);
    return `\n${items.join(',\n')}\n${pad}`.replace(/^/, '[').concat(']');
  }

  if (typeof value === 'object') {
    const entries = Object.entries(value as Record<string, unknown>).filter(
      ([, child]) => child !== undefined,
    );
    if (entries.length === 0) return '{}';
    const lines = entries.map(([key, child], index) => {
      const isLast = index === entries.length - 1;
      const trailingComma = isLast ? '' : ',';
      const childPath = path ? `${path}.${key}` : key;
      const renderedChild = renderValue(child, indent + 2, childPath);
      const comment = buildInlineOptionComment(childPath, child);
      if (renderedChild.startsWith('\n')) {
        return `${nextPad}${JSON.stringify(key)}: /* ${comment} */ ${renderedChild}${trailingComma}`;
      }
      return `${nextPad}${JSON.stringify(key)}: ${renderedChild}${trailingComma} // ${comment}`;
    });
    return `\n${lines.join('\n')}\n${pad}`.replace(/^/, '{').concat('}');
  }

  return 'null';
}

function renderSection(
  key: keyof ResolvedConfig,
  value: unknown,
  isLast: boolean,
  comments: string[],
): string {
  const lines: string[] = [];
  lines.push('  // ==========================================');
  for (const comment of comments) {
    lines.push(`  // ${comment}`);
  }
  lines.push('  // ==========================================');
  const inlineComment = buildInlineOptionComment(String(key), value);
  const renderedValue = renderValue(value, 2, String(key));
  if (renderedValue.startsWith('\n')) {
    lines.push(
      `  ${JSON.stringify(key)}: /* ${inlineComment} */ ${renderedValue}${isLast ? '' : ','}`,
    );
  } else {
    lines.push(
      `  ${JSON.stringify(key)}: ${renderedValue}${isLast ? '' : ','} // ${inlineComment}`,
    );
  }
  return lines.join('\n');
}

function createTemplateConfig(config: ResolvedConfig): ResolvedConfig {
  const templateConfig = deepCloneConfig(config);
  foldSubtitleCssManagedDefaults(templateConfig);
  if (templateConfig.keybindings.length === 0) {
    templateConfig.keybindings = DEFAULT_KEYBINDINGS.map((binding) => ({
      key: binding.key,
      command: binding.command === null ? null : [...binding.command],
    }));
  }
  return templateConfig;
}

export function generateConfigTemplate(
  config: ResolvedConfig = deepCloneConfig(DEFAULT_CONFIG),
): string {
  const templateConfig = createTemplateConfig(config);
  const lines: string[] = [];
  lines.push('/**');
  lines.push(' * SubMiner Example Configuration File');
  lines.push(' *');
  lines.push(' * This file is auto-generated from src/config/definitions.ts.');
  lines.push(
    ' * Copy to %APPDATA%/SubMiner/config.jsonc on Windows, or $XDG_CONFIG_HOME/SubMiner/config.jsonc (or ~/.config/SubMiner/config.jsonc) on Linux/macOS.',
  );
  lines.push(' */');
  lines.push('{');

  CONFIG_TEMPLATE_SECTIONS.forEach((section, index) => {
    lines.push('');
    const comments = [section.title, ...section.description, ...(section.notes ?? [])];
    lines.push(
      renderSection(
        section.key,
        templateConfig[section.key],
        index === CONFIG_TEMPLATE_SECTIONS.length - 1,
        comments,
      ),
    );
  });

  lines.push('}');
  lines.push('');
  return lines.join('\n');
}
