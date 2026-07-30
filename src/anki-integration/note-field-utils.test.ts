import assert from 'node:assert/strict';
import test from 'node:test';

import { resolveWordCardKind, type NoteFieldValueInfo } from './note-field-utils';

function kikuNote(values: Record<string, string> = {}): NoteFieldValueInfo {
  const defaults: Record<string, string> = {
    Expression: '単語',
    Sentence: '',
    IsWordAndSentenceCard: '',
    IsClickCard: '',
    IsSentenceCard: '',
    IsAudioCard: '',
  };
  return {
    fields: Object.fromEntries(
      Object.entries({ ...defaults, ...values }).map(([name, value]) => [name, { value }]),
    ),
  };
}

test('marks word-and-sentence cards by default when Kiku is enabled', () => {
  assert.equal(
    resolveWordCardKind(kikuNote(), { lapisEnabled: false, kikuEnabled: true }),
    'word-and-sentence',
  );
});

test('honors the configured word card kind', () => {
  assert.equal(
    resolveWordCardKind(kikuNote(), {
      lapisEnabled: false,
      kikuEnabled: true,
      wordCardKind: 'click',
    }),
    'click',
  );
});

test('marks nothing when neither Kiku nor Lapis is enabled', () => {
  assert.equal(
    resolveWordCardKind(kikuNote(), {
      lapisEnabled: false,
      kikuEnabled: false,
      wordCardKind: 'click',
    }),
    null,
  );
});

test('marks nothing when the word card kind is "none"', () => {
  assert.equal(
    resolveWordCardKind(kikuNote(), {
      lapisEnabled: true,
      kikuEnabled: false,
      wordCardKind: 'none',
    }),
    null,
  );
});

test('falls back to the default kind for an unrecognized setting', () => {
  assert.equal(
    resolveWordCardKind(kikuNote(), {
      lapisEnabled: false,
      kikuEnabled: true,
      wordCardKind: 'bogus' as never,
    }),
    'word-and-sentence',
  );
});

test('marks nothing when the note type lacks the configured flag field', () => {
  const note: NoteFieldValueInfo = {
    fields: { Expression: { value: '単語' }, Sentence: { value: '' } },
  };

  assert.equal(
    resolveWordCardKind(note, { lapisEnabled: false, kikuEnabled: true, wordCardKind: 'click' }),
    null,
  );
});

test('leaves cards already mined as sentence or audio cards alone', () => {
  for (const flagField of ['IsSentenceCard', 'IsAudioCard']) {
    assert.equal(
      resolveWordCardKind(kikuNote({ [flagField]: 'x' }), {
        lapisEnabled: false,
        kikuEnabled: true,
        wordCardKind: 'click',
      }),
      null,
      flagField,
    );
  }
});

test('re-affirms the configured kind when the note already carries its flag', () => {
  assert.equal(
    resolveWordCardKind(kikuNote({ IsSentenceCard: 'x' }), {
      lapisEnabled: false,
      kikuEnabled: true,
      wordCardKind: 'sentence',
    }),
    'sentence',
  );
});

test('overrides a differently flagged word card', () => {
  assert.equal(
    resolveWordCardKind(kikuNote({ IsWordAndSentenceCard: 'x' }), {
      lapisEnabled: false,
      kikuEnabled: true,
      wordCardKind: 'click',
    }),
    'click',
  );
});
