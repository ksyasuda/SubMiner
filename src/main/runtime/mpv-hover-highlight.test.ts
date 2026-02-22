import assert from 'node:assert/strict';
import test from 'node:test';

import { PartOfSpeech, type SubtitleData } from '../../types';
import {
  HOVER_TOKEN_MESSAGE,
  HOVER_SCRIPT_NAME,
  buildHoveredTokenMessageCommand,
  buildHoveredTokenPayload,
  createApplyHoveredTokenOverlayHandler,
} from './mpv-hover-highlight';

const SUBTITLE: SubtitleData = {
  text: '昨日は雨だった。',
  tokens: [
    {
      surface: '昨日',
      reading: 'きのう',
      headword: '昨日',
      startPos: 0,
      endPos: 2,
      partOfSpeech: PartOfSpeech.noun,
      isMerged: false,
      isKnown: false,
      isNPlusOneTarget: false,
    },
    {
      surface: 'は',
      reading: 'は',
      headword: 'は',
      startPos: 2,
      endPos: 3,
      partOfSpeech: PartOfSpeech.particle,
      isMerged: false,
      isKnown: true,
      isNPlusOneTarget: false,
    },
    {
      surface: '雨',
      reading: 'あめ',
      headword: '雨',
      startPos: 3,
      endPos: 4,
      partOfSpeech: PartOfSpeech.noun,
      isMerged: false,
      isKnown: false,
      isNPlusOneTarget: true,
    },
    {
      surface: 'だった。',
      reading: 'だった。',
      headword: 'だ',
      startPos: 4,
      endPos: 8,
      partOfSpeech: PartOfSpeech.other,
      isMerged: false,
      isKnown: false,
      isNPlusOneTarget: false,
    },
  ],
};

test('buildHoveredTokenPayload normalizes metadata and strips empty tokens', () => {
  const payload = buildHoveredTokenPayload({
    subtitle: SUBTITLE,
    hoveredTokenIndex: 2,
    revision: 5,
  });

  assert.equal(payload.revision, 5);
  assert.equal(payload.subtitle, '昨日は雨だった。');
  assert.equal(payload.hoveredTokenIndex, 2);
  assert.equal(payload.tokens.length, 4);
  assert.equal(payload.tokens[0]?.text, '昨日');
  assert.equal(payload.tokens[0]?.index, 0);
  assert.equal(payload.tokens[1]?.index, 1);
  assert.equal(payload.colors.hover, 'E7C06A');
});

test('buildHoveredTokenMessageCommand sends script-message-to subminer payload', () => {
  const payload = buildHoveredTokenPayload({
    subtitle: SUBTITLE,
    hoveredTokenIndex: 0,
    revision: 1,
  });

  const command = buildHoveredTokenMessageCommand(payload);

  assert.equal(command[0], 'script-message-to');
  assert.equal(command[1], HOVER_SCRIPT_NAME);
  assert.equal(command[2], HOVER_TOKEN_MESSAGE);

  const raw = command[3] as string;
  const parsed = JSON.parse(raw);
  assert.equal(parsed.revision, 1);
  assert.equal(parsed.hoveredTokenIndex, 0);
  assert.equal(parsed.subtitle, '昨日は雨だった。');
  assert.equal(parsed.tokens.length, 4);
});

test('createApplyHoveredTokenOverlayHandler sends clear payload when hovered token is missing', () => {
  const commands: Array<(string | number)[]> = [];
  const apply = createApplyHoveredTokenOverlayHandler({
    getMpvClient: () => ({
      connected: true,
      send: ({ command }: { command: (string | number)[] }) => {
        commands.push(command);
        return true;
      },
    }),
    getCurrentSubtitleData: () => SUBTITLE,
    getHoveredTokenIndex: () => null,
    getHoveredSubtitleRevision: () => 3,
  });

  apply();

  const parsed = JSON.parse(commands[0]?.[3] as string);
  assert.equal(parsed.hoveredTokenIndex, null);
  assert.equal(parsed.subtitle, null);
  assert.equal(parsed.tokens.length, 0);
});

test('createApplyHoveredTokenOverlayHandler sends highlight payload when hover is active', () => {
  const commands: Array<(string | number)[]> = [];
  const apply = createApplyHoveredTokenOverlayHandler({
    getMpvClient: () => ({
      connected: true,
      send: ({ command }: { command: (string | number)[] }) => {
        commands.push(command);
        return true;
      },
    }),
    getCurrentSubtitleData: () => SUBTITLE,
    getHoveredTokenIndex: () => 0,
    getHoveredSubtitleRevision: () => 3,
  });

  apply();

  const parsed = JSON.parse(commands[0]?.[3] as string);
  assert.equal(parsed.hoveredTokenIndex, 0);
  assert.equal(parsed.subtitle, '昨日は雨だった。');
  assert.equal(parsed.tokens.length, 4);
  assert.equal(commands[0]?.[0], 'script-message-to');
  assert.equal(commands[0]?.[1], HOVER_SCRIPT_NAME);
});
