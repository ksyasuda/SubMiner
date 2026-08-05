import assert from 'node:assert/strict';
import test from 'node:test';

import type { ChangelogEntry, ChangelogSnapshot } from '../../types/changelog';
import {
  createChangelogEntryNode,
  describeChangelogSource,
  resolveEntryBadge,
  shouldEntryStartExpanded,
  tokenizeInlineMarkdown,
} from './changelog-render';

type FakeNode = {
  tagName: string;
  className: string;
  title: string;
  open: boolean;
  tabIndex: number;
  dataset: Record<string, string>;
  children: FakeNode[];
  textContent: string;
};

function createFakeNode(tagName: string): FakeNode {
  const node: FakeNode = {
    tagName: tagName.toLowerCase(),
    className: '',
    title: '',
    open: false,
    tabIndex: 0,
    dataset: {},
    children: [],
    textContent: '',
  };
  return Object.assign(node, {
    appendChild: (child: FakeNode) => {
      node.children.push(child);
      return child;
    },
  });
}

function withFakeDocument<T>(run: () => T): T {
  const previous = Object.getOwnPropertyDescriptor(globalThis, 'document');
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    writable: true,
    value: {
      createElement: (tagName: string) => createFakeNode(tagName),
      createTextNode: (value: string) => ({
        tagName: '#text',
        textContent: value,
        children: [],
      }),
    },
  });
  try {
    return run();
  } finally {
    if (previous) {
      Object.defineProperty(globalThis, 'document', previous);
    } else {
      delete (globalThis as { document?: unknown }).document;
    }
  }
}

function flatten(node: FakeNode): FakeNode[] {
  return [node, ...(node.children ?? []).flatMap((child) => flatten(child as FakeNode))];
}

const ENTRY: ChangelogEntry = {
  version: '0.19.2',
  date: '2026-08-04',
  groupKey: '0.19',
  sections: [
    {
      heading: 'Fixed',
      items: [
        {
          text: '**Overlay:**',
          children: [
            { text: 'Fixed `something`.', children: [] },
            { text: 'Fixed another thing.', children: [] },
          ],
        },
        { text: 'Standalone fix.', children: [] },
      ],
      internal: false,
    },
    {
      heading: 'Internal',
      items: [{ text: 'Patched deps.', children: [] }],
      internal: true,
    },
  ],
};

test('inline markdown tokenizer splits code, bold, and links', () => {
  assert.deepEqual(tokenizeInlineMarkdown('Patched `undici` and **brace**.'), [
    { kind: 'text', value: 'Patched ' },
    { kind: 'code', value: 'undici' },
    { kind: 'text', value: ' and ' },
    { kind: 'strong', value: 'brace' },
    { kind: 'text', value: '.' },
  ]);
  assert.deepEqual(tokenizeInlineMarkdown('See [docs](https://example.com).'), [
    { kind: 'text', value: 'See ' },
    { kind: 'link', value: 'docs', href: 'https://example.com' },
    { kind: 'text', value: '.' },
  ]);
});

test('inline markdown tokenizer leaves plain text untouched', () => {
  assert.deepEqual(tokenizeInlineMarkdown('No markup here'), [
    { kind: 'text', value: 'No markup here' },
  ]);
});

test('entry badges mark the installed version and newer releases', () => {
  assert.equal(resolveEntryBadge('0.19.2', '0.19.2'), 'installed');
  assert.equal(resolveEntryBadge('0.20.0', '0.19.2'), 'newer');
  assert.equal(resolveEntryBadge('0.19.1', '0.19.2'), null);
});

test('only entries in the newest major.minor line start expanded', () => {
  const snapshot = { expandedGroupKey: '0.19' } as Pick<ChangelogSnapshot, 'expandedGroupKey'>;

  assert.equal(shouldEntryStartExpanded(ENTRY, snapshot), true);
  assert.equal(shouldEntryStartExpanded({ ...ENTRY, groupKey: '0.18' }, snapshot), false);
  assert.equal(shouldEntryStartExpanded(ENTRY, { expandedGroupKey: null }), false);
});

test('entry node renders a collapsible section with badge and internal block', () => {
  const node = withFakeDocument(() =>
    createChangelogEntryNode(ENTRY, { expanded: true, badge: 'installed', index: 0 }),
  ) as unknown as FakeNode;

  assert.equal(node.tagName, 'details');
  assert.equal(node.open, true);
  assert.equal(node.dataset.changelogVersion, '0.19.2');

  const nodes = flatten(node);
  const summary = nodes.find((child) => child.className === 'changelog-entry-summary');
  assert.ok(summary);
  assert.equal(summary?.dataset.changelogIndex, '0');
  assert.equal(
    nodes.find((child) => child.className === 'changelog-entry-version')?.textContent,
    'v0.19.2',
  );
  assert.equal(
    nodes.find((child) => child.className?.includes('changelog-entry-badge-installed'))
      ?.textContent,
    'Installed',
  );
  assert.equal(
    nodes.filter((child) => child.className === 'changelog-internal').length,
    1,
    'internal sections stay behind their own fold',
  );
});

test('entry node renders sub-bullets as a nested list under their lead bullet', () => {
  const node = withFakeDocument(() =>
    createChangelogEntryNode(ENTRY, { expanded: true, badge: null, index: 0 }),
  ) as unknown as FakeNode;

  const lists = flatten(node).filter((child) => child.className?.startsWith('changelog-items'));
  const topLevel = lists.find((list) => list.className === 'changelog-items');
  const nested = lists.filter((list) => list.className?.includes('changelog-items-nested'));

  assert.ok(topLevel);
  assert.equal(topLevel?.children.length, 2, 'lead bullet and standalone fix stay siblings');
  assert.equal(nested.length, 1, 'children render in exactly one nested list');
  assert.equal(nested[0]?.children.length, 2);

  // The nested list hangs off its parent <li>, not off the section.
  const leadItem = topLevel?.children[0];
  assert.equal(
    leadItem?.children.some((child) => child.className?.includes('changelog-items-nested')),
    true,
  );
});

test('entry node renders collapsed when it is outside the current line', () => {
  const node = withFakeDocument(() =>
    createChangelogEntryNode(ENTRY, { expanded: false, badge: null, index: 3 }),
  ) as unknown as FakeNode;

  assert.equal(node.open, false);
  assert.equal(
    flatten(node).some((child) => child.className?.startsWith('changelog-entry-badge')),
    false,
  );
});

test('source description names the release the changelog came from', () => {
  const base: ChangelogSnapshot = {
    entries: [],
    installedVersion: '0.19.2',
    latestVersion: null,
    expandedGroupKey: null,
    source: 'remote',
  };

  assert.equal(
    describeChangelogSource({ ...base, releaseTag: 'v0.20.0' }),
    'Latest release v0.20.0',
  );
  assert.equal(describeChangelogSource(base), 'Latest published changelog');
  assert.equal(describeChangelogSource({ ...base, source: 'bundled' }), 'Bundled changelog');
});
