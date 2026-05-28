import test from 'node:test';
import assert from 'node:assert/strict';
import * as ankiControls from './settings-anki-controls';

test('note field model preference keeps configured sentence-card model before Kiku fallback', () => {
  assert.equal(
    ankiControls.selectPreferredNoteFieldModelName(['Lapis Morph', 'Kiku'], 'Lapis Morph'),
    'Lapis Morph',
  );
});

test('note field model preference keeps configured sentence-card model case-insensitively', () => {
  assert.equal(
    ankiControls.selectPreferredNoteFieldModelName(['Lapis Morph', 'Kiku'], 'lapis morph'),
    'Lapis Morph',
  );
});

test('note field model preference prefers exact Lapis when Kiku is unavailable', () => {
  assert.equal(ankiControls.selectPreferredNoteFieldModelName(['Mining', 'Lapis'], ''), 'Lapis');
});

test('note field model preference prefers exact Kiku over exact Lapis', () => {
  assert.equal(ankiControls.selectPreferredNoteFieldModelName(['Lapis', 'Kiku'], ''), 'Kiku');
});

test('note field model preference does not treat partial Kiku matches as Kiku', () => {
  assert.equal(ankiControls.selectPreferredNoteFieldModelName(['Kikuchi', 'Mining'], ''), '');
});

test('note field model preference does not treat partial Lapis matches as Lapis', () => {
  assert.equal(ankiControls.selectPreferredNoteFieldModelName(['Mining', 'Lapis Morph'], ''), '');
});

test('note field model preference stays blank when no current Kiku or Lapis note type exists', () => {
  assert.equal(
    ankiControls.selectPreferredNoteFieldModelName(['Basic', 'Mining'], 'Lapis Morph'),
    '',
  );
});

test('known word deck rename selection keeps current deck on collision', () => {
  assert.equal(
    ankiControls.chooseKnownWordsDeckRenameValue(
      { Mining: ['Word'], Core: ['Expression'] },
      'Core',
      'Mining',
    ),
    'Core',
  );
});

test('Anki deck autofill uses inferred Yomitan deck only for untouched empty values', () => {
  assert.equal(ankiControls.chooseAnkiDeckAutofillValue('', 'Mining', false), 'Mining');
  assert.equal(ankiControls.chooseAnkiDeckAutofillValue('Current', 'Mining', false), null);
  assert.equal(ankiControls.chooseAnkiDeckAutofillValue('', 'Mining', true), null);
  assert.equal(ankiControls.chooseAnkiDeckAutofillValue('', '   ', false), null);
});
