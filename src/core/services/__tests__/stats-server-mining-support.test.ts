import assert from 'node:assert/strict';
import test from 'node:test';
import { createStatsMiningContext } from '../stats-server/mining-support.js';

test('mining timing observer errors do not replace successful phase results', async () => {
  const { timeMiningPhase } = createStatsMiningContext({
    onMiningTiming: () => {
      throw new Error('observer failed');
    },
  });

  const result = await timeMiningPhase('word', 'test', async () => 42);

  assert.equal(result, 42);
});

test('mining timing observer errors do not replace phase errors', async () => {
  const phaseError = new Error('phase failed');
  const { timeMiningPhase } = createStatsMiningContext({
    onMiningTiming: () => {
      throw new Error('observer failed');
    },
  });

  await assert.rejects(
    timeMiningPhase('word', 'test', async () => {
      throw phaseError;
    }),
    (error: unknown) => error === phaseError,
  );
});
