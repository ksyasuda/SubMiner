import assert from 'node:assert/strict';
import test from 'node:test';

import { buildAnimatedImageVideoFilter } from './media-generator';

test('buildAnimatedImageVideoFilter prepends a cloned first frame when lead-in is provided', () => {
  assert.equal(
    buildAnimatedImageVideoFilter({
      fps: 10,
      maxWidth: 640,
      leadingStillDuration: 1.25,
    }),
    'tpad=start_duration=1.25:start_mode=clone,fps=10,scale=w=640:h=-2',
  );
});
