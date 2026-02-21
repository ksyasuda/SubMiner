import test from 'node:test';
import assert from 'node:assert/strict';
import { createResolveContext } from './context';
import { applyIntegrationConfig } from './integrations';

test('jellyfin directPlayContainers are normalized', () => {
  const { context } = createResolveContext({
    jellyfin: {
      directPlayContainers: [' MKV ', 'mp4', '', ' WebM  ', 42 as unknown as string],
    },
  });

  applyIntegrationConfig(context);

  assert.deepEqual(context.resolved.jellyfin.directPlayContainers, ['mkv', 'mp4', 'webm']);
});
