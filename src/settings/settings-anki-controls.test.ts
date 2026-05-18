import test from 'node:test';
import assert from 'node:assert/strict';
import * as ankiControls from './settings-anki-controls';

test('note field model preference keeps a matching configured model before Kiku fallback', () => {
  assert.equal(
    ankiControls.selectPreferredNoteFieldModelName(['Lapis Morph', 'Kiku'], 'Lapis Morph'),
    'Lapis Morph',
  );
});

test('note field model preference matches configured model case-insensitively', () => {
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

test('note field model preference accepts partial Lapis matches', () => {
  assert.equal(
    ankiControls.selectPreferredNoteFieldModelName(['Mining', 'Lapis Morph'], ''),
    'Lapis Morph',
  );
});

test('note field model preference stays blank when no Kiku or Lapis note type exists', () => {
  assert.equal(
    ankiControls.selectPreferredNoteFieldModelName(['Basic', 'Mining'], 'Lapis Morph'),
    '',
  );
});
