import test from 'node:test';
import assert from 'node:assert/strict';
import { resolveConfig } from '../resolve';

test('subtitleFilter resolves presets and fills custom rule defaults', () => {
  const { resolved, warnings } = resolveConfig({
    subtitleFilter: {
      presets: { parentheses: true },
      customRules: [
        { pattern: '\\n' },
        { name: 'tag', pattern: '<i>', replacement: '', enabled: false },
      ],
    },
  });
  assert.deepEqual(warnings, []);
  assert.equal(resolved.subtitleFilter.presets.parentheses, true);
  assert.equal(resolved.subtitleFilter.presets.brackets, false);
  assert.deepEqual(resolved.subtitleFilter.customRules, [
    { name: '', pattern: '\\n', replacement: '', enabled: true },
    { name: 'tag', pattern: '<i>', replacement: '', enabled: false },
  ]);
});

test('subtitleFilter keeps empty and invalid patterns, warns on invalid, drops malformed rules', () => {
  const { resolved, warnings } = resolveConfig({
    subtitleFilter: {
      presets: { brackets: 'yes' as unknown as boolean },
      customRules: [
        { name: 'broken', pattern: '(' },
        { name: 'placeholder', pattern: '' },
        { name: 'no pattern' },
        'oops' as never,
      ],
    },
  });
  assert.deepEqual(
    resolved.subtitleFilter.customRules.map((rule) => rule.name),
    ['broken', 'placeholder'],
  );
  assert.equal(resolved.subtitleFilter.presets.brackets, false);
  assert.deepEqual(
    warnings.map((warning) => warning.path),
    [
      'subtitleFilter.presets.brackets',
      'subtitleFilter.customRules[0].pattern',
      'subtitleFilter.customRules[2]',
      'subtitleFilter.customRules[3]',
    ],
  );
});
