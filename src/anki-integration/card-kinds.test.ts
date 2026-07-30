import assert from 'node:assert/strict';
import test from 'node:test';

import { applyCardKindFlagFields } from './card-kinds';

function resolverFor(availableFieldNames: string[]) {
  return (preferredName: string): string | null =>
    availableFieldNames.find((name) => name.toLowerCase() === preferredName.toLowerCase()) ?? null;
}

const KIKU_FLAG_FIELDS = ['IsWordAndSentenceCard', 'IsClickCard', 'IsSentenceCard', 'IsAudioCard'];

test('flags the requested card kind and clears the others', () => {
  const fields: Record<string, string> = {};

  applyCardKindFlagFields(fields, 'click', resolverFor(KIKU_FLAG_FIELDS));

  assert.deepEqual(fields, {
    IsClickCard: 'x',
    IsWordAndSentenceCard: '',
    IsSentenceCard: '',
    IsAudioCard: '',
  });
});

test('matches flag fields case-insensitively', () => {
  const fields: Record<string, string> = {};

  applyCardKindFlagFields(fields, 'word-and-sentence', resolverFor(['iswordandsentencecard']));

  assert.deepEqual(fields, { iswordandsentencecard: 'x' });
});

test('leaves flags untouched when the note type has no flag for a word card kind', () => {
  const fields: Record<string, string> = {};

  applyCardKindFlagFields(
    fields,
    'click',
    resolverFor(['IsWordAndSentenceCard', 'IsSentenceCard']),
  );

  assert.deepEqual(fields, {});
});

test('clears stale flags for explicit mine actions even without the target flag', () => {
  const fields: Record<string, string> = {};

  applyCardKindFlagFields(
    fields,
    'audio',
    resolverFor(['IsWordAndSentenceCard', 'IsSentenceCard']),
  );

  assert.deepEqual(fields, { IsWordAndSentenceCard: '', IsSentenceCard: '' });
});

test('does not blank the target flag it just set', () => {
  const fields: Record<string, string> = {};

  applyCardKindFlagFields(fields, 'sentence', resolverFor(['IsSentenceCard']));

  assert.deepEqual(fields, { IsSentenceCard: 'x' });
});
