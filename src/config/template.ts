import { ResolvedConfig } from '../types';
import { CONFIG_TEMPLATE_SECTIONS, DEFAULT_CONFIG, deepCloneConfig } from './definitions';

function renderValue(value: unknown, indent = 0): string {
  const pad = ' '.repeat(indent);
  const nextPad = ' '.repeat(indent + 2);

  if (value === null) return 'null';
  if (typeof value === 'string') return JSON.stringify(value);
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);

  if (Array.isArray(value)) {
    if (value.length === 0) return '[]';
    const items = value.map((item) => `${nextPad}${renderValue(item, indent + 2)}`);
    return `\n${items.join(',\n')}\n${pad}`.replace(/^/, '[').concat(']');
  }

  if (typeof value === 'object') {
    const entries = Object.entries(value as Record<string, unknown>).filter(
      ([, child]) => child !== undefined,
    );
    if (entries.length === 0) return '{}';
    const lines = entries.map(
      ([key, child]) => `${nextPad}${JSON.stringify(key)}: ${renderValue(child, indent + 2)}`,
    );
    return `\n${lines.join(',\n')}\n${pad}`.replace(/^/, '{').concat('}');
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
  lines.push(`  ${JSON.stringify(key)}: ${renderValue(value, 2)}${isLast ? '' : ','}`);
  return lines.join('\n');
}

export function generateConfigTemplate(
  config: ResolvedConfig = deepCloneConfig(DEFAULT_CONFIG),
): string {
  const lines: string[] = [];
  lines.push('/**');
  lines.push(' * SubMiner Example Configuration File');
  lines.push(' *');
  lines.push(' * This file is auto-generated from src/config/definitions.ts.');
  lines.push(
    ' * Copy to $XDG_CONFIG_HOME/SubMiner/config.jsonc (or ~/.config/SubMiner/config.jsonc) and edit as needed.',
  );
  lines.push(' */');
  lines.push('{');

  CONFIG_TEMPLATE_SECTIONS.forEach((section, index) => {
    lines.push('');
    const comments = [section.title, ...section.description, ...(section.notes ?? [])];
    lines.push(
      renderSection(
        section.key,
        config[section.key],
        index === CONFIG_TEMPLATE_SECTIONS.length - 1,
        comments,
      ),
    );
  });

  lines.push('}');
  lines.push('');
  return lines.join('\n');
}
