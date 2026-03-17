import test from 'node:test';
import assert from 'node:assert/strict';
import {
  serializeInitialSubtitleWebsocketMessage,
  serializeSubtitleMarkup,
  serializeSubtitleWebsocketMessage,
} from './subtitle-ws';
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

test('serializeSubtitleMarkup preserves tooltip attrs and name-match precedence', () => {
  const payload: SubtitleData = {
    text: 'ignored',
    tokens: [
      {
        surface: '無事',
        reading: 'ぶじ',
        headword: '無事',
        startPos: 0,
        endPos: 2,
        partOfSpeech: PartOfSpeech.other,
        isMerged: false,
        isKnown: true,
        isNPlusOneTarget: false,
        jlptLevel: 'N2',
        frequencyRank: 745,
      },
      {
        surface: 'アレクシア',
        reading: 'あれくしあ',
        headword: 'アレクシア',
        startPos: 2,
        endPos: 7,
        partOfSpeech: PartOfSpeech.other,
        isMerged: false,
        isKnown: false,
        isNPlusOneTarget: true,
        isNameMatch: true,
        jlptLevel: 'N5',
        frequencyRank: 12,
      },
    ],
  };

  const markup = serializeSubtitleMarkup(payload, frequencyOptions);
  assert.match(
    markup,
    /<span class="word word-known word-jlpt-n2" data-reading="ぶじ" data-headword="無事" data-frequency-rank="745" data-jlpt-level="N2">無事<\/span>/,
  );
  assert.match(
    markup,
    /<span class="word word-name-match" data-reading="あれくしあ" data-headword="アレクシア">アレクシア<\/span>/,
  );
  assert.doesNotMatch(markup, /word-name-match word-known|word-known word-name-match/);
  assert.doesNotMatch(markup, /word-name-match word-n-plus-one|word-n-plus-one word-name-match/);
  assert.doesNotMatch(markup, /data-frequency-rank="12"|data-jlpt-level="N5"|word-jlpt-n5/);
});

test('serializeSubtitleWebsocketMessage emits sentence payload', () => {
  const payload: SubtitleData = {
    text: '字幕',
    tokens: null,
  };

  const raw = serializeSubtitleWebsocketMessage(payload, frequencyOptions);
  assert.deepEqual(JSON.parse(raw), {
    version: 1,
    text: '字幕',
    sentence: '字幕',
    tokens: [],
  });
});

test('serializeSubtitleWebsocketMessage emits structured token api payload', () => {
  const payload: SubtitleData = {
    text: '無事',
    tokens: [
      {
        surface: '無事',
        reading: 'ぶじ',
        headword: '無事',
        startPos: 0,
        endPos: 2,
        partOfSpeech: PartOfSpeech.other,
        isMerged: false,
        isKnown: true,
        isNPlusOneTarget: false,
        jlptLevel: 'N2',
        frequencyRank: 745,
      },
    ],
  };

  const raw = serializeSubtitleWebsocketMessage(payload, frequencyOptions);
  assert.deepEqual(JSON.parse(raw), {
    version: 1,
    text: '無事',
    sentence:
      '<span class="word word-known word-jlpt-n2" data-reading="ぶじ" data-headword="無事" data-frequency-rank="745" data-jlpt-level="N2">無事</span>',
    tokens: [
      {
        surface: '無事',
        reading: 'ぶじ',
        headword: '無事',
        startPos: 0,
        endPos: 2,
        partOfSpeech: PartOfSpeech.other,
        isMerged: false,
        isKnown: true,
        isNPlusOneTarget: false,
        isNameMatch: false,
        jlptLevel: 'N2',
        frequencyRank: 745,
        className: 'word word-known word-jlpt-n2',
        frequencyRankLabel: '745',
        jlptLevelLabel: 'N2',
      },
    ],
  });
});

test('serializeInitialSubtitleWebsocketMessage keeps annotated current subtitle content', () => {
  const payload: SubtitleData = {
    text: 'ignored fallback',
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
    ],
  };

  const raw = serializeInitialSubtitleWebsocketMessage(payload, frequencyOptions);
  assert.deepEqual(JSON.parse(raw ?? ''), {
    version: 1,
    text: 'ignored fallback',
    sentence: '<span class="word word-known">既知</span>',
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
        isNameMatch: false,
        className: 'word word-known',
        frequencyRankLabel: null,
        jlptLevelLabel: null,
      },
    ],
  });
});
