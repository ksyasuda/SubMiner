import assert from 'node:assert/strict';
import test from 'node:test';
import {
  isLibraryGroupCollapsed,
  toggleLibraryGroupCollapse,
} from './LibraryTab';

interface FakeGroup {
  key: string;
  items: { videoId: number }[];
}

const multiVideoGroup: FakeGroup = {
  key: 'series-a',
  items: [{ videoId: 1 }, { videoId: 2 }, { videoId: 3 }],
};

const singletonGroup: FakeGroup = {
  key: 'video-b',
  items: [{ videoId: 99 }],
};

test('isLibraryGroupCollapsed defaults to collapsed for multi-video groups', () => {
  assert.equal(isLibraryGroupCollapsed(multiVideoGroup, new Map()), true);
});

test('isLibraryGroupCollapsed defaults to expanded for singleton groups', () => {
  assert.equal(isLibraryGroupCollapsed(singletonGroup, new Map()), false);
});

test('isLibraryGroupCollapsed honors an explicit user override', () => {
  const overrides = new Map<string, boolean>([
    ['series-a', false],
    ['video-b', true],
  ]);
  assert.equal(isLibraryGroupCollapsed(multiVideoGroup, overrides), false);
  assert.equal(isLibraryGroupCollapsed(singletonGroup, overrides), true);
});

test('toggleLibraryGroupCollapse flips a multi-video group from collapsed to expanded', () => {
  const next = toggleLibraryGroupCollapse(new Map(), multiVideoGroup);
  assert.equal(next.get('series-a'), false);
  assert.equal(isLibraryGroupCollapsed(multiVideoGroup, next), false);
});

test('toggleLibraryGroupCollapse flips a singleton group from expanded to collapsed', () => {
  const next = toggleLibraryGroupCollapse(new Map(), singletonGroup);
  assert.equal(next.get('video-b'), true);
  assert.equal(isLibraryGroupCollapsed(singletonGroup, next), true);
});

test('toggleLibraryGroupCollapse toggles back when called twice', () => {
  const once = toggleLibraryGroupCollapse(new Map(), multiVideoGroup);
  const twice = toggleLibraryGroupCollapse(once, multiVideoGroup);
  assert.equal(isLibraryGroupCollapsed(multiVideoGroup, twice), true);
});
