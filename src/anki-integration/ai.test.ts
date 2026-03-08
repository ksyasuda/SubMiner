import test from 'node:test';
import assert from 'node:assert/strict';

import { resolveSentenceBackText } from './ai';

test('resolveSentenceBackText returns secondary subtitle when ai is disabled', async () => {
  const result = await resolveSentenceBackText(
    {
      sentence: '日本語',
      secondarySubText: 'existing translation',
      aiEnabled: false,
      aiConfig: {},
    },
    {
      logWarning: () => undefined,
    },
  );

  assert.equal(result, 'existing translation');
});

test('resolveSentenceBackText uses shared ai config when enabled', async () => {
  const result = await resolveSentenceBackText(
    {
      sentence: '日本語',
      secondarySubText: '',
      aiEnabled: true,
      aiConfig: {
        enabled: true,
        apiKey: 'abc',
        model: 'openai/gpt-4o-mini',
      },
    },
    {
      logWarning: () => undefined,
      translateSentence: async (request) => {
        assert.equal(request.apiKey, 'abc');
        assert.equal(request.model, 'openai/gpt-4o-mini');
        return 'translated';
      },
    },
  );

  assert.equal(result, 'translated');
});

test('resolveSentenceBackText falls back to sentence when ai translation fails with no secondary subtitle', async () => {
  const result = await resolveSentenceBackText(
    {
      sentence: '日本語',
      aiEnabled: true,
      aiConfig: {
        enabled: true,
        apiKey: 'abc',
      },
    },
    {
      logWarning: () => undefined,
      translateSentence: async () => null,
    },
  );

  assert.equal(result, '日本語');
});
