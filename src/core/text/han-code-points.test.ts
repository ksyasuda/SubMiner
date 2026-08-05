import assert from 'node:assert/strict';
import test from 'node:test';
import { HAN_CODE_POINT_RANGES, HAN_REGEXP_CLASS_BODY, isHanCodePoint } from './han-code-points';

test('every range boundary is inside the table', () => {
  for (const [start, end] of HAN_CODE_POINT_RANGES) {
    for (const codePoint of [start, end]) {
      assert.ok(isHanCodePoint(codePoint), `expected U+${codePoint.toString(16)} to be Han`);
    }
  }

  // Extension J (Unicode 17) and the Compatibility blocks are the ones a
  // BMP-only table used to miss.
  assert.ok(isHanCodePoint(0x323b0));
  assert.ok(isHanCodePoint(0x33479));
  assert.ok(isHanCodePoint(0xf900));
  assert.ok(isHanCodePoint(0x2f800));
});

test('no unified ideograph the runtime knows about falls outside the table', () => {
  // One direction only: a runtime with older Unicode data simply checks fewer
  // code points, where asserting the reverse would fail on Extension J.
  const unifiedIdeograph = /\p{Unified_Ideograph}/u;

  for (let codePoint = 0x3000; codePoint <= 0x40000; codePoint += 1) {
    if (unifiedIdeograph.test(String.fromCodePoint(codePoint))) {
      assert.ok(
        isHanCodePoint(codePoint),
        `expected unified ideograph U+${codePoint.toString(16)} to be in the table`,
      );
    }
  }
});

test('code points just outside the table are rejected', () => {
  for (const codePoint of [0x33ff, 0x4dc0, 0xa000, 0x1f000, 0x3347a]) {
    assert.equal(
      isHanCodePoint(codePoint),
      false,
      `expected U+${codePoint.toString(16)} not to be Han`,
    );
  }
});

test('the regexp class body matches the same code points as the predicate', () => {
  const classRegExp = new RegExp(`^[${HAN_REGEXP_CLASS_BODY}]$`, 'u');

  for (const codePoint of [0x3400, 0x4e00, 0x9fff, 0xf900, 0x20000, 0x323b0, 0x33479]) {
    assert.match(String.fromCodePoint(codePoint), classRegExp);
  }
  for (const codePoint of [0x3040, 0x30ff, 0x33fa, 0x3347a]) {
    assert.doesNotMatch(String.fromCodePoint(codePoint), classRegExp);
  }
});
