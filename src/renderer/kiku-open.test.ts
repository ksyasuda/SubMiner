import assert from 'node:assert/strict';
import test from 'node:test';
import { prepareForKikuFieldGroupingOpen } from './kiku-open';

test('prepareForKikuFieldGroupingOpen closes lookup popup before pausing playback', () => {
  const calls: string[] = [];

  prepareForKikuFieldGroupingOpen({
    closeLookupWindow: () => {
      calls.push('close');
      return true;
    },
    pausePlayback: () => {
      calls.push('pause');
    },
  });

  assert.deepEqual(calls, ['close', 'pause']);
});

test('prepareForKikuFieldGroupingOpen still pauses playback when no popup is open', () => {
  const calls: string[] = [];

  prepareForKikuFieldGroupingOpen({
    closeLookupWindow: () => {
      calls.push('close');
      return false;
    },
    pausePlayback: () => {
      calls.push('pause');
    },
  });

  assert.deepEqual(calls, ['close', 'pause']);
});
