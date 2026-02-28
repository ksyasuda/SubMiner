import test from 'node:test';
import assert from 'node:assert/strict';
import type { MpvSubtitleRenderMetrics } from '../../types';
import {
  dispatchMpvProtocolMessage,
  MPV_REQUEST_ID_TRACK_LIST_SECONDARY,
  MpvProtocolHandleMessageDeps,
  splitMpvMessagesFromBuffer,
  parseVisibilityProperty,
  asBoolean,
} from './mpv-protocol';

function createDeps(overrides: Partial<MpvProtocolHandleMessageDeps> = {}): {
  deps: MpvProtocolHandleMessageDeps;
  state: {
    subText: string;
    secondarySubText: string;
    events: Array<unknown>;
    commands: unknown[];
    mediaPath: string;
    restored: number;
  };
} {
  const state = {
    subText: '',
    secondarySubText: '',
    events: [] as Array<unknown>,
    commands: [] as unknown[],
    mediaPath: '',
    restored: 0,
  };
  const metrics: MpvSubtitleRenderMetrics = {
    subPos: 100,
    subFontSize: 36,
    subScale: 1,
    subMarginY: 0,
    subMarginX: 0,
    subFont: '',
    subSpacing: 0,
    subBold: false,
    subItalic: false,
    subBorderSize: 0,
    subShadowOffset: 0,
    subAssOverride: 'yes',
    subScaleByWindow: true,
    subUseMargins: true,
    osdHeight: 0,
    osdDimensions: null,
  };

  return {
    state,
    deps: {
      getResolvedConfig: () => ({
        secondarySub: { secondarySubLanguages: ['ja'] },
      }),
      getSubtitleMetrics: () => metrics,
      isVisibleOverlayVisible: () => false,
      emitSubtitleChange: (payload) => state.events.push(payload),
      emitSubtitleAssChange: (payload) => state.events.push(payload),
      emitSubtitleTiming: (payload) => state.events.push(payload),
      emitSecondarySubtitleChange: (payload) => state.events.push(payload),
      getCurrentSubText: () => state.subText,
      setCurrentSubText: (text) => {
        state.subText = text;
      },
      setCurrentSubStart: () => {},
      getCurrentSubStart: () => 0,
      setCurrentSubEnd: () => {},
      getCurrentSubEnd: () => 0,
      emitMediaPathChange: (payload) => {
        state.mediaPath = payload.path;
      },
      emitMediaTitleChange: (payload) => state.events.push(payload),
      emitSubtitleMetricsChange: (payload) => state.events.push(payload),
      setCurrentSecondarySubText: (text) => {
        state.secondarySubText = text;
      },
      resolvePendingRequest: () => false,
      setSecondarySubVisibility: () => {},
      syncCurrentAudioStreamIndex: () => {},
      setCurrentAudioTrackId: () => {},
      setCurrentTimePos: () => {},
      getCurrentTimePos: () => 0,
      getPendingPauseAtSubEnd: () => false,
      setPendingPauseAtSubEnd: () => {},
      getPauseAtTime: () => null,
      setPauseAtTime: () => {},
      emitTimePosChange: () => {},
      emitPauseChange: () => {},
      autoLoadSecondarySubTrack: () => {},
      setCurrentVideoPath: () => {},
      emitSecondarySubtitleVisibility: (payload) => state.events.push(payload),
      setCurrentAudioStreamIndex: () => {},
      sendCommand: (payload) => {
        state.commands.push(payload);
        return true;
      },
      restorePreviousSecondarySubVisibility: () => {
        state.restored += 1;
      },
      setPreviousSecondarySubVisibility: () => {
        // intentionally not tracked in this unit test
      },
      ...overrides,
    },
  };
}

test('dispatchMpvProtocolMessage emits subtitle text on property change', async () => {
  const { deps, state } = createDeps();

  await dispatchMpvProtocolMessage(
    { event: 'property-change', name: 'sub-text', data: '字幕' },
    deps,
  );

  assert.equal(state.subText, '字幕');
  assert.deepEqual(state.events, [{ text: '字幕', isOverlayVisible: false }]);
});

test('dispatchMpvProtocolMessage enforces sub-visibility hidden when overlay suppression is enabled', async () => {
  const { deps, state } = createDeps({
    isVisibleOverlayVisible: () => true,
  });

  await dispatchMpvProtocolMessage(
    { event: 'property-change', name: 'sub-visibility', data: 'yes' },
    deps,
  );

  assert.deepEqual(state.commands, [
    {
      command: ['set_property', 'sub-visibility', false],
    },
    {
      command: ['set_property', 'sub-visibility', 'no'],
    },
    {
      command: ['set', 'sub-visibility', 'no'],
    },
  ]);
});

test('dispatchMpvProtocolMessage skips sub-visibility suppression when overlay is hidden', async () => {
  const { deps, state } = createDeps({
    isVisibleOverlayVisible: () => false,
  });

  await dispatchMpvProtocolMessage(
    { event: 'property-change', name: 'sub-visibility', data: 'yes' },
    deps,
  );

  assert.equal(state.commands.length, 0);
});

test('dispatchMpvProtocolMessage sets secondary subtitle track based on track list response', async () => {
  const { deps, state } = createDeps();

  await dispatchMpvProtocolMessage(
    {
      request_id: MPV_REQUEST_ID_TRACK_LIST_SECONDARY,
      data: [
        { type: 'audio', id: 1, lang: 'eng' },
        { type: 'sub', id: 2, lang: 'ja' },
      ],
    },
    deps,
  );

  assert.deepEqual(state.commands, [{ command: ['set_property', 'secondary-sid', 2] }]);
});

test('dispatchMpvProtocolMessage restores secondary visibility on shutdown', async () => {
  const { deps, state } = createDeps();

  await dispatchMpvProtocolMessage({ event: 'shutdown' }, deps);

  assert.equal(state.restored, 1);
});

test('dispatchMpvProtocolMessage pauses on sub-end when pendingPauseAtSubEnd is set', async () => {
  let pendingPauseAtSubEnd = true;
  let pauseAtTime: number | null = null;
  const { deps, state } = createDeps({
    getPendingPauseAtSubEnd: () => pendingPauseAtSubEnd,
    setPendingPauseAtSubEnd: (next) => {
      pendingPauseAtSubEnd = next;
    },
    getCurrentSubText: () => '字幕',
    setCurrentSubEnd: () => {},
    getCurrentSubEnd: () => 0,
    setPauseAtTime: (next) => {
      pauseAtTime = next;
    },
  });

  await dispatchMpvProtocolMessage({ event: 'property-change', name: 'sub-end', data: 42 }, deps);

  assert.equal(pendingPauseAtSubEnd, false);
  assert.equal(pauseAtTime, 42);
  assert.deepEqual(state.events, [{ text: '字幕', start: 0, end: 0 }]);
  assert.deepEqual(state.commands[state.commands.length - 1], {
    command: ['set_property', 'pause', false],
  });
});

test('splitMpvMessagesFromBuffer parses complete lines and preserves partial buffer', () => {
  const parsed = splitMpvMessagesFromBuffer(
    '{"event":"shutdown"}\n{"event":"property-change","name":"media-title","data":"x"}\n{"partial"',
  );

  assert.equal(parsed.messages.length, 2);
  assert.equal(parsed.nextBuffer, '{"partial"');
  assert.equal(parsed.messages[0]!.event, 'shutdown');
  assert.equal(parsed.messages[1]!.name, 'media-title');
});

test('splitMpvMessagesFromBuffer reports invalid JSON lines', () => {
  const errors: Array<{ line: string; error?: string }> = [];

  splitMpvMessagesFromBuffer('{"event":"x"}\n{invalid}\n', undefined, (line, error) => {
    errors.push({ line, error: String(error) });
  });

  assert.equal(errors.length, 1);
  assert.equal(errors[0]!.line, '{invalid}');
});

test('visibility and boolean parsers handle text values', () => {
  assert.equal(parseVisibilityProperty('true'), true);
  assert.equal(parseVisibilityProperty('0'), false);
  assert.equal(parseVisibilityProperty('unknown'), null);
  assert.equal(asBoolean('yes', false), true);
  assert.equal(asBoolean('0', true), false);
});
