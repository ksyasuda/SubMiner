import assert from 'node:assert/strict';
import test from 'node:test';
import { setMpvCurrentSecondarySubText } from './autoplay-subtitle-priming-runtime';

test('setMpvCurrentSecondarySubText uses client setter when available', () => {
  const calls: string[] = [];
  const client = {
    currentSecondarySubText: '',
    setCurrentSecondarySubText: (text: string) => {
      calls.push(text);
    },
  };

  setMpvCurrentSecondarySubText(client, 'secondary');

  assert.deepEqual(calls, ['secondary']);
  assert.equal(client.currentSecondarySubText, '');
});
