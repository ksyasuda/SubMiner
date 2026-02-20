import test from 'node:test';
import assert from 'node:assert/strict';
import { serializeSubtitleMarkup, serializeSubtitleWebsocketMessage } from './subtitle-ws';
import { PartOfSpeech, type SubtitleData } from '../../types';

const frequencyOptions = {
  enabled: true,
  topX: 1000,
  mode: 'banded' as const,
};

test('serializeSubtitleMarkup escapes plain text and preserves line breaks', () => {
  const payload: SubtitleData = {
    text: 'a < b\nx & y',
    tokens: null,
  };

  assert.equal(serializeSubtitleMarkup(payload, frequencyOptions), 'a &lt; b<br>x &amp; y');
});

test('serializeSubtitleMarkup includes known, n+1, jlpt, and frequency classes', () => {
  const payload: SubtitleData = {
    text: 'ignored',
    tokens: [
      {
        surface: '既知',
        reading: '',
        headword: '',
        startPos: 0,
        endPos: 2,
        partOfSpeech: PartOfSpeech.other,
        isMerged: false,
        isKnown: true,
        isNPlusOneTarget: false,
      },
      {
        surface: '新語',
        reading: '',
        headword: '',
        startPos: 2,
        endPos: 4,
        partOfSpeech: PartOfSpeech.other,
        isMerged: false,
        isKnown: false,
        isNPlusOneTarget: true,
      },
      {
        surface: '級',
        reading: '',
        headword: '',
        startPos: 4,
        endPos: 5,
        partOfSpeech: PartOfSpeech.other,
        isMerged: false,
        isKnown: false,
        isNPlusOneTarget: false,
        jlptLevel: 'N3',
      },
      {
        surface: '頻度',
        reading: '',
        headword: '',
        startPos: 5,
        endPos: 7,
        partOfSpeech: PartOfSpeech.other,
        isMerged: false,
        isKnown: false,
        isNPlusOneTarget: false,
        frequencyRank: 10,
      },
    ],
  };

  const markup = serializeSubtitleMarkup(payload, frequencyOptions);
  assert.match(markup, /word word-known/);
  assert.match(markup, /word word-n-plus-one/);
  assert.match(markup, /word word-jlpt-n3/);
  assert.match(markup, /word word-frequency-band-1/);
});

test('serializeSubtitleWebsocketMessage emits sentence payload', () => {
  const payload: SubtitleData = {
    text: '字幕',
    tokens: null,
  };

  const raw = serializeSubtitleWebsocketMessage(payload, frequencyOptions);
  assert.deepEqual(JSON.parse(raw), { sentence: '字幕' });
});
