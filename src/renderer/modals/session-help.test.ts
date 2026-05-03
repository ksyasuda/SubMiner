import assert from 'node:assert/strict';
import test from 'node:test';

import { SPECIAL_COMMANDS } from '../../config/definitions';
import { describeSessionHelpCommand, formatSessionHelpKeybinding } from './session-help.js';

test('session help describes sub-seek commands as subtitle-line navigation', () => {
  assert.equal(describeSessionHelpCommand(['sub-seek', 1]), 'Jump to next subtitle');
  assert.equal(describeSessionHelpCommand(['sub-seek', -1]), 'Jump to previous subtitle');
});

test('session help describes subtitle-delay shift special commands separately from sub-seek', () => {
  assert.equal(
    describeSessionHelpCommand([SPECIAL_COMMANDS.SHIFT_SUB_DELAY_TO_NEXT_SUBTITLE_START]),
    'Shift subtitle delay to next cue',
  );
  assert.equal(
    describeSessionHelpCommand([SPECIAL_COMMANDS.SHIFT_SUB_DELAY_TO_PREVIOUS_SUBTITLE_START]),
    'Shift subtitle delay to previous cue',
  );
});

test('session help formats bracket keybindings as physical keys', () => {
  assert.equal(formatSessionHelpKeybinding('Shift+BracketRight'), 'Shift + ]');
  assert.equal(formatSessionHelpKeybinding('Shift+BracketLeft'), 'Shift + [');
});
