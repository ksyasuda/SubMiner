import type { CharacterBirthday, CharacterRecord } from './types';

const MONTH_NAMES: ReadonlyArray<[number, string]> = [
  [1, 'January'],
  [2, 'February'],
  [3, 'March'],
  [4, 'April'],
  [5, 'May'],
  [6, 'June'],
  [7, 'July'],
  [8, 'August'],
  [9, 'September'],
  [10, 'October'],
  [11, 'November'],
  [12, 'December'],
];

const SEX_DISPLAY: ReadonlyArray<[string, string]> = [
  ['m', '♂ Male'],
  ['f', '♀ Female'],
  ['male', '♂ Male'],
  ['female', '♀ Female'],
];

function formatBirthday(birthday: CharacterBirthday | null): string {
  if (!birthday) return '';
  const [month, day] = birthday;
  const monthName = MONTH_NAMES.find(([m]) => m === month)?.[1] || 'Unknown';
  return `${monthName} ${day}`;
}

export function formatCharacterStats(character: CharacterRecord): string {
  const parts: string[] = [];
  const normalizedSex = character.sex.trim().toLowerCase();
  const sexDisplay = SEX_DISPLAY.find(([key]) => key === normalizedSex)?.[1];
  if (sexDisplay) parts.push(sexDisplay);
  if (character.age.trim()) parts.push(`${character.age.trim()} years`);
  if (character.bloodType.trim()) parts.push(`Blood Type ${character.bloodType.trim()}`);
  const birthday = formatBirthday(character.birthday);
  if (birthday) parts.push(`Birthday: ${birthday}`);
  return parts.join(' • ');
}

export function parseCharacterDescription(raw: string): {
  fields: Array<{ key: string; value: string }>;
  text: string;
} {
  const cleaned = raw.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, ' ');
  const lines = cleaned.split(/\n/);
  const fields: Array<{ key: string; value: string }> = [];
  const textLines: string[] = [];

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    const match = trimmed.match(/^__([^_]+):__\s*(.+)$/);
    if (match) {
      const value = match[2]!
        .replace(/__([^_]+)__/g, '$1')
        .replace(/\*\*([^*]+)\*\*/g, '$1')
        .replace(/_([^_]+)_/g, '$1')
        .replace(/\*([^*]+)\*/g, '$1')
        .trim();
      fields.push({ key: match[1]!.trim(), value });
    } else {
      textLines.push(trimmed);
    }
  }

  const text = textLines
    .join(' ')
    .replace(/\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)/g, '$1')
    .replace(/https?:\/\/\S+/g, '')
    .replace(/__([^_]+)__/g, '$1')
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/~!/g, '')
    .replace(/!~/g, '')
    .replace(/\s+/g, ' ')
    .trim();

  return { fields, text };
}
