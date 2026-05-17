import test from 'node:test';
import assert from 'node:assert/strict';
import * as ankiControls from './settings-anki-controls';

test('note field model preference chooses Kiku before configured Lapis default', () => {
  assert.equal(
    ankiControls.selectPreferredNoteFieldModelName(['Lapis Morph', 'Kiku'], 'Lapis Morph'),
    'Kiku',
  );
});

test('note field model preference falls back to Lapis when Kiku is unavailable', () => {
  assert.equal(
    ankiControls.selectPreferredNoteFieldModelName(['Mining', 'Lapis Morph'], 'Lapis Morph'),
    'Lapis Morph',
  );
});

test('note field model preference does not treat partial Kiku matches as Kiku', () => {
  assert.equal(
    ankiControls.selectPreferredNoteFieldModelName(['Kikuchi', 'Lapis Morph'], 'Lapis Morph'),
    'Lapis Morph',
  );
});

test('note field model preference stays blank when no Kiku or Lapis note type exists', () => {
  assert.equal(
    ankiControls.selectPreferredNoteFieldModelName(['Basic', 'Mining'], 'Lapis Morph'),
    '',
  );
});
