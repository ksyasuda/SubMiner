import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  assOverrideSignature,
  assToPlainText,
  collectAssOverrideCommands,
  extractAssOverrideBlocks,
  hasAssTemporalOverride,
  isAnimatedAssEffectKind,
  isAssTemporalCommand,
  normalizePlainSubtitleText,
  parseAssEffectField,
  removeLiveGlyphFragmentLines,
  removeAssControlDebrisLines,
} from './ass-text';

test('assToPlainText drops vector drawing runs', () => {
  assert.equal(
    assToPlainText(
      '{\\an5\\pos(730,1042)\\p1\\blur1}m 20 0 b 10 0 0 10 0 20 b 0 31 10 40 20 40 {\\p0}',
    ),
    '',
  );
});

test('assToPlainText keeps text around drawing runs on the same event', () => {
  assert.equal(
    assToPlainText('{\\p1}m 0 0 l 10 10{\\p0}本文{\\p1}m 5 5 l 6 6{\\p0}続き'),
    '本文続き',
  );
});

test('assToPlainText leaves \\pos alone when no drawing mode is active', () => {
  assert.equal(assToPlainText('{\\pos(960,1068)\\bord3}位置指定'), '位置指定');
});

test('assToPlainText does not read \\pos as a drawing tag', () => {
  assert.equal(assToPlainText('{\\p1\\pos(1,2)}m 0 0 l 5 5'), '');
});

test('assToPlainText resolves line-break and space escapes', () => {
  assert.equal(assToPlainText('一行目\\N二行目'), '一行目\n二行目');
  assert.equal(assToPlainText('一行目\\n二行目'), '一行目\n二行目');
  assert.equal(assToPlainText('一行目\\N二行目', ' '), '一行目 二行目');
  assert.equal(assToPlainText('間\\h隔'), '間 隔');
});

test('assToPlainText matches mpv on brace and backslash sequences', () => {
  // mpv has no `\{` / `\}` / `\\` escapes: the backslashes are literal text and the
  // braces still open and close an override block.
  assert.equal(assToPlainText('\\{注\\}'), '\\');
  assert.equal(assToPlainText('\\\\N'), '\\\n');
});

test('assToPlainText renders an unclosed override block verbatim', () => {
  // mpv shows the stray brace; guessing where the block ended can eat a whole line.
  assert.equal(assToPlainText('本文{\\pos(1,2)'), '本文{\\pos(1,2)');
});

test('assToPlainText is idempotent', () => {
  const samples = [
    '{\\an5\\p1}m 0 0 l 5 5{\\p0}本文',
    '\\{注\\}',
    '\\\\N',
    '本文{\\pos(1,2)',
    '一行目\\N二行目\\h終わり',
  ];

  for (const sample of samples) {
    const once = assToPlainText(sample);
    assert.equal(assToPlainText(once), once, sample);
  }
});

test('assToPlainText normalizes CRLF before converting', () => {
  assert.equal(assToPlainText('一行目\r\n二行目'), '一行目\n二行目');
});

test('removeAssControlDebrisLines drops malformed spacer resets without eating dialogue', () => {
  assert.equal(
    removeAssControlDebrisLines('Visible line\n\\\n{\\fr0\n\\{\\frz287.5'),
    'Visible line',
  );
  assert.equal(removeAssControlDebrisLines('本文{\\pos(1,2)'), '本文{\\pos(1,2)');
});

test('normalizePlainSubtitleText settles whitespace without decoding ASS', () => {
  // A brace reaching this layer is literal text mpv chose to show, not markup.
  assert.equal(normalizePlainSubtitleText('本文{\\pos(1,2)'), '本文{\\pos(1,2)');
  assert.equal(normalizePlainSubtitleText('一行目\\N二行目'), '一行目\n二行目');
  assert.equal(
    normalizePlainSubtitleText('一行目\\N二行目', { collapseLineBreaks: true }),
    '一行目 二行目',
  );
  assert.equal(normalizePlainSubtitleText('  余白  ', { trim: false }), '  余白  ');
});

test('normalizePlainSubtitleText is idempotent', () => {
  for (const sample of ['一行目\\N二行目', '間\\h隔', '本文{\\pos(1,2)', '  余白  ']) {
    const once = normalizePlainSubtitleText(sample);
    assert.equal(normalizePlainSubtitleText(once), once, sample);
  }
});

test('extractAssOverrideBlocks returns block contents', () => {
  assert.deepEqual(extractAssOverrideBlocks('{\\an8}上{\\fad(200,200)}下'), [
    '\\an8',
    '\\fad(200,200)',
  ]);
  assert.deepEqual(extractAssOverrideBlocks('括弧なし'), []);
});

test('collectAssOverrideCommands captures names and arguments from blocks only', () => {
  const commands = collectAssOverrideCommands('{\\pos(1,2)\\1c&HFFFFFF&\\kf30}歌詞');

  assert.deepEqual(commands, [
    { name: 'pos', args: '1,2', animated: false },
    { name: '1c', args: '&HFFFFFF&', animated: false },
    { name: 'kf', args: '30', animated: false },
  ]);

  // A `\pos(...)` sitting in visible text is not typesetting markup.
  assert.deepEqual(collectAssOverrideCommands('\\pos(730,1042) と書いてある'), []);
});

test('collectAssOverrideCommands marks tags animated by a wrapping \\t', () => {
  const commands = collectAssOverrideCommands('{\\clip(0,0,10,10)\\t(0,500,\\frz30)}文字');

  assert.deepEqual(
    commands.map((command) => [command.name, command.animated]),
    [
      ['clip', false],
      ['t', false],
      ['frz', true],
    ],
  );
  assert.equal(hasAssTemporalOverride(commands), true);
});

test('collectAssOverrideCommands stops descending into deeply nested \\t tags', () => {
  // Nested far past the recursion cap. Uncapped, this recurses once per level, and a
  // pathological line (real files reach one or two levels) overflows the stack.
  const nesting = 32;
  const block = `{${'\\t(0,500,'.repeat(nesting)}\\frz30${')'.repeat(nesting)}}文字`;

  const commands = collectAssOverrideCommands(block);

  // The outer `\t` plus one per allowed recursion level, and nothing from below the cap.
  assert.equal(commands.length, 9);
  assert.deepEqual(new Set(commands.map((command) => command.name)), new Set(['t']));
  assert.equal(hasAssTemporalOverride(commands), true);
});

test('hasAssTemporalOverride ignores static placement and shape tags', () => {
  assert.equal(
    hasAssTemporalOverride(collectAssOverrideCommands('{\\pos(1,2)\\clip(m 1 1)\\blur2}文字')),
    false,
  );
  assert.equal(hasAssTemporalOverride(collectAssOverrideCommands('{\\move(1,2,3,4)}文字')), true);
});

test('isAssTemporalCommand covers only intrinsically animated tags', () => {
  for (const command of ['t', 'move', 'k', 'kf', 'ko', 'K']) {
    assert.equal(isAssTemporalCommand(command), true, command);
  }
  for (const command of ['clip', 'iclip', 'frz', 'fscx', 'blur', 'be', 'pos', 'fad']) {
    assert.equal(isAssTemporalCommand(command), false, command);
  }
});

test('assOverrideSignature distinguishes events by their override values', () => {
  const first = assOverrideSignature(collectAssOverrideCommands('{\\clip(m 1 1)}歌詞'));
  const second = assOverrideSignature(collectAssOverrideCommands('{\\clip(m 2 2)}歌詞'));
  const repeat = assOverrideSignature(collectAssOverrideCommands('{\\clip(m 1 1)}別の行'));

  assert.notEqual(first, second);
  assert.equal(first, repeat);
});

test('parseAssEffectField classifies the event-level Effect column', () => {
  assert.equal(parseAssEffectField(''), 'none');
  assert.equal(parseAssEffectField('   '), 'none');
  assert.equal(parseAssEffectField('Banner;20;1;0'), 'banner');
  assert.equal(parseAssEffectField('Scroll up;0;0;30;10'), 'scroll');
  assert.equal(parseAssEffectField('Scroll down;0;0;30;10'), 'scroll');
  assert.equal(parseAssEffectField('Karaoke'), 'karaoke');
  assert.equal(parseAssEffectField('fx-template'), 'other');
});

test('parseAssEffectField matches stock effect names exactly', () => {
  // Custom effect names that merely start with a stock name are not stock effects.
  assert.equal(parseAssEffectField('scrolling-credit'), 'other');
  assert.equal(parseAssEffectField('bannerfx;1'), 'other');
  assert.equal(parseAssEffectField('karaoke-template'), 'other');
  assert.equal(parseAssEffectField('Scroll'), 'other');
});

test('isAnimatedAssEffectKind covers the stock animated effects only', () => {
  assert.equal(isAnimatedAssEffectKind('karaoke'), true);
  assert.equal(isAnimatedAssEffectKind('banner'), true);
  assert.equal(isAnimatedAssEffectKind('scroll'), true);
  // Typesetting groups put static template names in this column too.
  assert.equal(isAnimatedAssEffectKind('other'), false);
  assert.equal(isAnimatedAssEffectKind('none'), false);
});

test('removeLiveGlyphFragmentLines drops a per-glyph typesetting wall and its syllable', () => {
  const wall = [...'wansdumretoikhI'].join('\n');
  assert.equal(removeLiveGlyphFragmentLines(`${wall}\ntai`), '');
});

test('removeLiveGlyphFragmentLines keeps concurrent dialogue beside a glyph wall', () => {
  const wall = [...'wansdumretoikhI'].join('\n');
  assert.equal(removeLiveGlyphFragmentLines(`${wall}\nそれよりも　ノート…`), 'それよりも　ノート…');
});

test('removeLiveGlyphFragmentLines leaves ordinary short lines alone', () => {
  const text = 'え\nはい。\nそうだな';
  assert.equal(removeLiveGlyphFragmentLines(text), text);
});

test('normalizePlainSubtitleText folds cue-boundary blank lines for text consumers', () => {
  // The display layer splits on the blank line before normalizing; everyone else --
  // tokenizer, cache key, dedup gate, mined sentence -- wants the plain line form.
  assert.equal(
    normalizePlainSubtitleText('\u4e00\u884c\u76ee\n\n\u4e8c\u884c\u76ee'),
    '\u4e00\u884c\u76ee\n\u4e8c\u884c\u76ee',
  );
  assert.equal(
    normalizePlainSubtitleText('\u4e00\u884c\u76ee\n\n\u4e8c\u884c\u76ee', {
      collapseLineBreaks: true,
    }),
    '\u4e00\u884c\u76ee \u4e8c\u884c\u76ee',
  );
});
