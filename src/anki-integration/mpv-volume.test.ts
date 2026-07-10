import assert from 'node:assert/strict';
import test from 'node:test';

import { resolveMpvVolumeScale } from './mpv-volume';

test('resolveMpvVolumeScale converts numeric mpv volume with mpv software volume curve', async () => {
  const requested: string[] = [];

  const scale = await resolveMpvVolumeScale(
    {
      requestProperty: async (name) => {
        requested.push(name);
        return 75;
      },
    },
    true,
  );

  assert.equal(scale, 0.421875);
  assert.deepEqual(requested, ['volume']);
});

test('resolveMpvVolumeScale skips mpv when mirroring is disabled', async () => {
  let requested = false;

  const scale = await resolveMpvVolumeScale(
    {
      requestProperty: async () => {
        requested = true;
        return 50;
      },
    },
    false,
  );

  assert.equal(scale, undefined);
  assert.equal(requested, false);
});

test('resolveMpvVolumeScale falls back to unity for missing, failed, or invalid values', async () => {
  assert.equal(await resolveMpvVolumeScale({}, true), 1);
  assert.equal(
    await resolveMpvVolumeScale(
      {
        requestProperty: async () => {
          throw new Error('disconnected');
        },
      },
      true,
    ),
    1,
  );
  assert.equal(await resolveMpvVolumeScale({ requestProperty: async () => '50' }, true), 1);
  assert.equal(await resolveMpvVolumeScale({ requestProperty: async () => Number.NaN }, true), 1);
  assert.equal(await resolveMpvVolumeScale({ requestProperty: async () => -1 }, true), 1);
});
