import assert from 'node:assert/strict';
import type { PathLike } from 'node:fs';
import test from 'node:test';
import {
  isCompiledMacOSHelperCurrent,
  MacOSWindowTracker,
  parseMacOSHelperOutput,
} from './macos-tracker';

const GEOMETRY = { x: 10, y: 20, width: 1280, height: 720 };

const PARSE_CASES = [
  {
    name: 'geometry and focused state',
    input: '120,240,1280,720,1',
    expected: { geometry: { x: 120, y: 240, width: 1280, height: 720 }, focused: true },
  },
  {
    name: 'geometry with unfocused state',
    input: '120,240,1280,720,0',
    expected: { geometry: { x: 120, y: 240, width: 1280, height: 720 }, focused: false },
  },
  {
    name: 'minimized state',
    input: 'minimized',
    expected: { geometry: null, focused: false, minimized: true },
  },
  {
    name: 'active focused state without geometry',
    input: 'active',
    expected: { geometry: null, focused: true, active: true },
  },
  {
    name: 'inactive state without geometry',
    input: 'inactive',
    expected: { geometry: null, focused: false, inactive: true },
  },
];

for (const c of PARSE_CASES) {
  test(`parseMacOSHelperOutput parses ${c.name}`, () => {
    assert.deepEqual(parseMacOSHelperOutput(c.input), c.expected);
  });
}

test('isCompiledMacOSHelperCurrent rejects binaries older than the Swift source', () => {
  const binaryPath = '/tmp/get-mpv-window-macos';
  const sourcePath = '/tmp/get-mpv-window-macos.swift';
  const statSync = (binaryMtimeMs: number, sourceMtimeMs: number) => (targetPath: PathLike) =>
    ({
      mtimeMs: String(targetPath) === binaryPath ? binaryMtimeMs : sourceMtimeMs,
    }) as never;
  const helperFs = {
    existsSync: () => true,
    statSync: statSync(1_000, 2_000),
  };

  assert.equal(isCompiledMacOSHelperCurrent(binaryPath, sourcePath, helperFs), false);

  helperFs.statSync = statSync(2_000, 1_000);

  assert.equal(isCompiledMacOSHelperCurrent(binaryPath, sourcePath, helperFs), true);
});

test('MacOSWindowTracker slows polling while focused target is stable', async () => {
  const scheduledDelays: number[] = [];
  let callIndex = 0;
  const tracker = new MacOSWindowTracker('/tmp/mpv.sock', {
    resolveHelper: () => ({
      helperPath: 'helper',
      helperType: 'binary',
    }),
    runHelper: async () => {
      callIndex += 1;
      return { stdout: '10,20,1280,720,1', stderr: '' };
    },
    fastPollIntervalMs: 250,
    stablePollIntervalMs: 1_000,
    setPollTimeout: ((_callback: () => void, delayMs: number) => {
      scheduledDelays.push(delayMs);
      return {} as ReturnType<typeof setTimeout>;
    }) as never,
    clearPollTimeout: (() => {}) as never,
  } as never);

  tracker.start();
  await new Promise((resolve) => setTimeout(resolve, 0));
  tracker.stop();

  assert.equal(callIndex, 1);
  assert.deepEqual(scheduledDelays, [1_000]);
});

test('MacOSWindowTracker keeps fast polling while target is not focused', async () => {
  const scheduledDelays: number[] = [];
  const tracker = new MacOSWindowTracker('/tmp/mpv.sock', {
    resolveHelper: () => ({
      helperPath: 'helper',
      helperType: 'binary',
    }),
    runHelper: async () => ({ stdout: '10,20,1280,720,0', stderr: '' }),
    fastPollIntervalMs: 250,
    stablePollIntervalMs: 1_000,
    setPollTimeout: ((_callback: () => void, delayMs: number) => {
      scheduledDelays.push(delayMs);
      return {} as ReturnType<typeof setTimeout>;
    }) as never,
    clearPollTimeout: (() => {}) as never,
  } as never);

  tracker.start();
  await new Promise((resolve) => setTimeout(resolve, 0));
  tracker.stop();

  assert.deepEqual(scheduledDelays, [250]);
});

type HelperReply = string | Error;

type PollExpectation = {
  tracking: boolean;
  focused: boolean;
  geometry: typeof GEOMETRY | null;
  minimized?: boolean;
  focusChanges?: boolean[];
};

type PollStep = {
  reply: HelperReply;
  // Clock advance before this poll; drives the tracking-loss grace window.
  advanceMs?: number;
  expect: PollExpectation;
};

const FOCUSED = `${GEOMETRY.x},${GEOMETRY.y},${GEOMETRY.width},${GEOMETRY.height},1`;
const UNFOCUSED = `${GEOMETRY.x},${GEOMETRY.y},${GEOMETRY.width},${GEOMETRY.height},0`;
const tracked = (focused: boolean, extra: Partial<PollExpectation> = {}): PollExpectation => ({
  tracking: true,
  focused,
  geometry: GEOMETRY,
  ...extra,
});
const lost: PollExpectation = { tracking: false, focused: false, geometry: null };

function createTracker(
  replies: HelperReply[],
  options: ConstructorParameters<typeof MacOSWindowTracker>[1] = {},
) {
  let callIndex = 0;
  let now = 1_000;
  const focusChanges: boolean[] = [];
  const tracker = new MacOSWindowTracker('/tmp/mpv.sock', {
    resolveHelper: () => ({ helperPath: 'helper.swift', helperType: 'swift' }),
    runHelper: async () => {
      const reply = replies[callIndex++] ?? replies.at(-1)!;
      if (reply instanceof Error) throw Object.assign(reply, { stderr: 'timeout' });
      return { stdout: reply, stderr: '' };
    },
    now: () => now,
    ...options,
  });
  tracker.onWindowFocusChange = (focused) => {
    focusChanges.push(focused);
  };
  return {
    tracker,
    focusChanges,
    async poll(advanceMs = 0): Promise<void> {
      now += advanceMs;
      await tracker.refreshNow();
    },
  };
}

const POLL_SEQUENCES: Array<{
  name: string;
  options: ConstructorParameters<typeof MacOSWindowTracker>[1];
  steps: PollStep[];
}> = [
  {
    name: 'keeps the last geometry through a single helper miss',
    options: { trackingLossGraceMs: 0 },
    steps: [
      { reply: FOCUSED, expect: tracked(true) },
      { reply: 'not-found', expect: tracked(true) },
      { reply: FOCUSED, expect: tracked(true) },
    ],
  },
  {
    name: 'preserves target focus on helper not-found while retaining geometry',
    options: { trackingLossGraceMs: 1_500 },
    steps: [
      { reply: FOCUSED, expect: tracked(true, { focusChanges: [true] }) },
      { reply: 'not-found', expect: tracked(true, { focusChanges: [true] }) },
    ],
  },
  {
    name: 'keeps focused fullscreen target through active helper misses after grace',
    options: { trackingLossGraceMs: 500 },
    steps: [
      { reply: FOCUSED, expect: tracked(true) },
      { reply: 'active', advanceMs: 1_000, expect: tracked(true) },
      { reply: 'active', advanceMs: 1_000, expect: tracked(true) },
    ],
  },
  {
    name: 'drops previously focused target after repeated not-found misses exceed grace',
    options: { trackingLossGraceMs: 500 },
    steps: [
      { reply: FOCUSED, expect: tracked(true) },
      { reply: 'not-found', advanceMs: 1_000, expect: tracked(true, { focusChanges: [true] }) },
      { reply: 'not-found', advanceMs: 1_000, expect: { ...lost, focusChanges: [true, false] } },
    ],
  },
  {
    name: 'drops previously focused target after repeated helper execution failures exceed grace',
    options: { trackingLossGraceMs: 500 },
    steps: [
      { reply: FOCUSED, expect: tracked(true) },
      { reply: new Error('helper timed out'), advanceMs: 1_000, expect: tracked(true) },
      {
        reply: new Error('helper timed out'),
        advanceMs: 1_000,
        expect: { ...lost, focusChanges: [true, false] },
      },
    ],
  },
  {
    name: 'marks target unfocused on explicit inactive helper signal',
    options: { trackingLossGraceMs: 1_500 },
    steps: [
      { reply: FOCUSED, expect: tracked(true) },
      { reply: 'inactive', expect: tracked(false, { focusChanges: [true, false] }) },
    ],
  },
  {
    name: 'refreshes to focused when the helper reports the target active',
    options: {},
    steps: [
      { reply: UNFOCUSED, expect: tracked(false) },
      { reply: 'active', expect: tracked(true) },
    ],
  },
  {
    name: 'drops tracking after consecutive helper misses',
    options: { trackingLossGraceMs: 0 },
    steps: [
      { reply: UNFOCUSED, expect: tracked(false) },
      { reply: 'not-found', expect: tracked(false) },
      { reply: 'not-found', expect: lost },
    ],
  },
  {
    name: 'keeps tracking through repeated helper misses inside grace window',
    options: { trackingLossGraceMs: 1_500 },
    steps: [
      { reply: FOCUSED, expect: tracked(true) },
      { reply: 'not-found', advanceMs: 250, expect: tracked(true) },
      { reply: 'not-found', advanceMs: 250, expect: tracked(true) },
      { reply: 'not-found', advanceMs: 250, expect: tracked(true) },
    ],
  },
  {
    name: 'drops tracking after grace window expires',
    options: { trackingLossGraceMs: 500 },
    steps: [
      { reply: UNFOCUSED, expect: tracked(false) },
      { reply: 'not-found', advanceMs: 250, expect: tracked(false) },
      { reply: 'not-found', advanceMs: 250, expect: tracked(false) },
      { reply: 'not-found', advanceMs: 250, expect: tracked(false) },
      { reply: 'not-found', advanceMs: 250, expect: lost },
    ],
  },
  {
    name: 'reports minimized target, then drops tracking after the minimized grace',
    options: { minimizedTrackingLossGraceMs: 200 },
    steps: [
      { reply: FOCUSED, expect: tracked(true, { minimized: false }) },
      { reply: 'minimized', advanceMs: 250, expect: tracked(false, { minimized: true }) },
      { reply: 'minimized', advanceMs: 250, expect: { ...lost, minimized: true } },
    ],
  },
];

for (const sequence of POLL_SEQUENCES) {
  test(`MacOSWindowTracker ${sequence.name}`, async () => {
    const { tracker, focusChanges, poll } = createTracker(
      sequence.steps.map((step) => step.reply),
      sequence.options,
    );

    for (const [index, step] of sequence.steps.entries()) {
      await poll(step.advanceMs);
      const label = `poll ${index + 1}`;
      assert.equal(tracker.isTracking(), step.expect.tracking, `${label}: isTracking`);
      assert.equal(tracker.isTargetWindowFocused(), step.expect.focused, `${label}: focused`);
      assert.deepEqual(tracker.getGeometry(), step.expect.geometry, `${label}: geometry`);
      if (step.expect.minimized !== undefined) {
        assert.equal(
          tracker.isTargetWindowMinimized(),
          step.expect.minimized,
          `${label}: minimized`,
        );
      }
      if (step.expect.focusChanges) {
        assert.deepEqual(focusChanges, step.expect.focusChanges, `${label}: focus changes`);
      }
    }
  });
}
