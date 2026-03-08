import test from 'node:test';
import assert from 'node:assert/strict';

import { extractAiText, normalizeOpenAiBaseUrl } from './client';

test('normalizeOpenAiBaseUrl appends v1 when missing', () => {
  assert.equal(normalizeOpenAiBaseUrl('https://openrouter.ai/api'), 'https://openrouter.ai/api/v1');
  assert.equal(normalizeOpenAiBaseUrl('https://example.test/v1'), 'https://example.test/v1');
});

test('extractAiText joins OpenAI structured text parts', () => {
  assert.equal(
    extractAiText([
      { type: 'text', text: 'hello ' },
      { type: 'text', text: 'world' },
      { type: 'image', text: 'ignored' },
    ]),
    'hello world',
  );
  assert.equal(extractAiText(' plain text '), 'plain text');
});
