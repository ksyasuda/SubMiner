import assert from 'node:assert/strict';
import test from 'node:test';
import type { SubtitleCue, SubtitleFilterConfig } from '../types/subtitle';
import {
  filterSubtitleCues,
  filterSubtitleText,
  getSubtitleFilterPatternError,
} from './subtitle-text-filter';

const presets = (...ids: Array<keyof NonNullable<SubtitleFilterConfig['presets']>>) => ({
  presets: Object.fromEntries(ids.map((id) => [id, true])),
});

test('parentheses preset removes speaker names, readings and notes anywhere in a line', () => {
  const config = presets('parentheses');
  assert.equal(filterSubtitleText('（山田）　元気ですか？', config), '元気ですか？');
  assert.equal(filterSubtitleText('漢字（かんじ）を読む。', config), '漢字を読む。');
  assert.equal(filterSubtitleText('そうなんだ（笑）', config), 'そうなんだ');
  assert.equal(filterSubtitleText('Hello (laughs) there', config), 'Hello there');
});

test('parentheses preset keeps a line wrapped entirely in parentheses', () => {
  const config = presets('parentheses');
  assert.equal(filterSubtitleText('（どうしよう…）', config), '（どうしよう…）');
  assert.equal(
    filterSubtitleText('（山田）おはよう\n（眠いな…）', config),
    'おはよう\n（眠いな…）',
  );
});

test('a line emptied by a filter is dropped, an untouched symbol-only line is kept', () => {
  assert.equal(filterSubtitleText('［拍手］', presets('brackets')), '');
  assert.equal(filterSubtitleText('♪～\n歌詞だよ', presets('musicSymbols')), '歌詞だよ');
  assert.equal(filterSubtitleText('［拍手］\n……', presets('brackets')), '……');
});

test('speaker prefix preset only strips a short name before a colon at line start', () => {
  const config = presets('speakerPrefix');
  assert.equal(filterSubtitleText('山田：元気？', config), '元気？');
  assert.equal(filterSubtitleText('10:30に集合', config), '10:30に集合');
});

test('custom rules run after presets, in order, with their replacement', () => {
  const config: SubtitleFilterConfig = {
    presets: { parentheses: true },
    customRules: [
      { name: 'skip', pattern: 'x', enabled: false },
      { name: 'join lines', pattern: '\\n', replacement: '' },
      { name: 'invalid', pattern: '(' },
    ],
  };
  assert.equal(filterSubtitleText('（山田）こんにちは\n世界x', config), 'こんにちは世界x');
});

test('each rule sees the tidied output of the rules before it', () => {
  const config: SubtitleFilterConfig = {
    presets: { parentheses: true },
    customRules: [{ name: 'greeting', pattern: '^今日は', replacement: '' }],
  };
  assert.equal(filterSubtitleText('（山田）　今日はいい天気', config), 'いい天気');
});

test('filterSubtitleCues drops emptied cues and keeps parser metadata', () => {
  const cue: SubtitleCue = { startTime: 1, endTime: 2, text: '（山田）はい' };
  Object.defineProperty(cue, 'assFurigana', { value: ['やまだ'], enumerable: false });
  const cues = filterSubtitleCues(
    [cue, { startTime: 3, endTime: 4, text: '［拍手］' }],
    presets('parentheses', 'brackets'),
  );
  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.text, 'はい');
  assert.deepEqual(cues[0]!.assFurigana, ['やまだ']);
});

test('getSubtitleFilterPatternError reports invalid and empty patterns', () => {
  assert.equal(getSubtitleFilterPatternError('[（(].*?[）)]'), null);
  assert.match(getSubtitleFilterPatternError('(') ?? '', /Invalid|Unterminated/);
  assert.equal(getSubtitleFilterPatternError(''), 'Pattern is empty.');
});
