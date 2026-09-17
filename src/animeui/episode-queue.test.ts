import test from 'node:test';
import assert from 'node:assert/strict';
import { describeQueuePosition, queueKey, queuePositions } from './episode-queue';
import type { AnimeBrowserQueueEntry } from '../types/anime-browser';

function entry(overrides: Partial<AnimeBrowserQueueEntry> = {}): AnimeBrowserQueueEntry {
  return {
    sourceId: 'source',
    animeUrl: '/anime',
    animeTitle: 'Anime',
    episodeUrl: '/episode-1',
    episodeName: 'Episode 1',
    episodeNumber: 1,
    ...overrides,
  };
}

test('positions count across the whole queue, not within one anime', () => {
  const positions = queuePositions([
    entry(),
    entry({ animeUrl: '/other', episodeUrl: '/other-1' }),
    entry({ episodeUrl: '/episode-2' }),
  ]);

  assert.equal(positions.get(queueKey('source', '/episode-1')), 1);
  assert.equal(positions.get(queueKey('source', '/other-1')), 2);
  assert.equal(positions.get(queueKey('source', '/episode-2')), 3);
});

test('the same episode url on two sources keeps two positions', () => {
  const positions = queuePositions([entry(), entry({ sourceId: 'other' })]);

  assert.equal(positions.get(queueKey('source', '/episode-1')), 1);
  assert.equal(positions.get(queueKey('other', '/episode-1')), 2);
});

test('the episode about to play is named rather than numbered', () => {
  assert.equal(describeQueuePosition(1), 'next up');
  assert.equal(describeQueuePosition(4), '#4 in queue');
});
