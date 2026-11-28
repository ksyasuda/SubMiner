import test from 'node:test';
import assert from 'node:assert/strict';

import { createSubsyncModal } from './subsync.js';
import type { SubsyncManualPayload, SubsyncManualRunRequest } from '../../types';

type Listener = () => void;

function createClassList() {
  const classes = new Set<string>();
  return {
    add: (...tokens: string[]) => {
      for (const token of tokens) classes.add(token);
    },
    remove: (...tokens: string[]) => {
      for (const token of tokens) classes.delete(token);
    },
    toggle: (token: string, force?: boolean) => {
      if (force === undefined) {
        if (classes.has(token)) classes.delete(token);
        else classes.add(token);
        return classes.has(token);
      }
      if (force) classes.add(token);
      else classes.delete(token);
      return force;
    },
    contains: (token: string) => classes.has(token),
  };
}

function createEventTarget() {
  const listeners = new Map<string, Listener[]>();
  return {
    addEventListener: (event: string, listener: Listener) => {
      const existing = listeners.get(event) ?? [];
      existing.push(listener);
      listeners.set(event, existing);
    },
    dispatch: (event: string) => {
      for (const listener of listeners.get(event) ?? []) {
        listener();
      }
    },
  };
}

function createDeferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((nextResolve) => {
    resolve = nextResolve;
  });
  return { promise, resolve };
}

function createSelectStub() {
  const options: Array<{ value: string; textContent: string }> = [];
  const events = createEventTarget();
  let innerHTML = '';
  let value = '';

  return {
    options,
    disabled: false,
    addEventListener: events.addEventListener,
    dispatch: events.dispatch,
    get innerHTML(): string {
      return innerHTML;
    },
    set innerHTML(next: string) {
      innerHTML = next;
      if (next === '') {
        options.length = 0;
        value = '';
      }
    },
    get value(): string {
      return value;
    },
    set value(next: string) {
      value = next;
    },
    appendChild(option: { value: string; textContent: string }) {
      options.push(option);
      if (!value) value = option.value;
      return option;
    },
  };
}

function createTestHarness(
  runSubsyncManual: (request: SubsyncManualRunRequest) => Promise<{ ok: boolean; message: string }>,
) {
  const overlayClassList = createClassList();
  const modalClassList = createClassList();
  const statusClassList = createClassList();
  const referenceLabelClassList = createClassList();
  const targetLabelClassList = createClassList();
  const runButtonEvents = createEventTarget();
  const closeButtonEvents = createEventTarget();
  const engineAlassEvents = createEventTarget();
  const engineFfsubsyncEvents = createEventTarget();

  const runButton = {
    disabled: false,
    addEventListener: runButtonEvents.addEventListener,
    dispatch: runButtonEvents.dispatch,
  };

  const closeButton = {
    addEventListener: closeButtonEvents.addEventListener,
    dispatch: closeButtonEvents.dispatch,
  };

  const subsyncEngineAlass = {
    checked: false,
    disabled: false,
    addEventListener: engineAlassEvents.addEventListener,
    dispatch: engineAlassEvents.dispatch,
  };

  const subsyncEngineFfsubsync = {
    checked: false,
    disabled: false,
    addEventListener: engineFfsubsyncEvents.addEventListener,
    dispatch: engineFfsubsyncEvents.dispatch,
  };

  const referenceSelect = createSelectStub();
  const targetSelect = createSelectStub();

  let notifyClosedCalls = 0;
  let notifyOpenedCalls = 0;

  const previousWindow = (globalThis as { window?: unknown }).window;
  const previousDocument = (globalThis as { document?: unknown }).document;

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        runSubsyncManual,
        notifyOverlayModalOpened: () => {
          notifyOpenedCalls += 1;
        },
        notifyOverlayModalClosed: () => {
          notifyClosedCalls += 1;
        },
      },
    },
  });

  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => ({ value: '', textContent: '' }),
    },
  });

  const ctx = {
    dom: {
      overlay: { classList: overlayClassList },
      subsyncModal: {
        classList: modalClassList,
        setAttribute: () => {},
      },
      subsyncCloseButton: closeButton,
      subsyncEngineAlass,
      subsyncEngineFfsubsync,
      subsyncReferenceLabel: { classList: referenceLabelClassList },
      subsyncReferenceSelect: referenceSelect,
      subsyncTargetLabel: { classList: targetLabelClassList },
      subsyncTargetSelect: targetSelect,
      subsyncRunButton: runButton,
      subsyncStatus: {
        textContent: '',
        classList: statusClassList,
      },
    },
    state: {
      subsyncModalOpen: false,
      subsyncSubtitleTracks: [],
      subsyncSubmitting: false,
      isOverSubtitle: false,
    },
  };

  const modal = createSubsyncModal(ctx as never, {
    modalStateReader: {
      isAnyModalOpen: () => false,
    },
    syncSettingsModalSubtitleSuppression: () => {},
  });

  return {
    ctx,
    modal,
    runButton,
    referenceSelect,
    targetSelect,
    referenceLabelClassList,
    statusClassList,
    getNotifyClosedCalls: () => notifyClosedCalls,
    getNotifyOpenedCalls: () => notifyOpenedCalls,
    restoreGlobals: () => {
      Object.defineProperty(globalThis, 'window', {
        configurable: true,
        value: previousWindow,
      });
      Object.defineProperty(globalThis, 'document', {
        configurable: true,
        value: previousDocument,
      });
    },
  };
}

async function flushMicrotasks(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
}

const BASE_PAYLOAD: SubsyncManualPayload = {
  subtitleTracks: [
    { id: 1, label: 'External #1 - jpn (active)' },
    { id: 2, label: 'External #2 - eng' },
  ],
  defaultReferenceTrackId: 2,
  defaultTargetTrackId: 1,
  videoReferenceAvailable: true,
  ffsubsyncAvailable: true,
};

function payloadWith(overrides: Partial<SubsyncManualPayload>): SubsyncManualPayload {
  return { ...BASE_PAYLOAD, ...overrides };
}

test('manual subsync failure closes during run, then reopens modal with error', async () => {
  const deferred = createDeferred<{ ok: boolean; message: string }>();
  const harness = createTestHarness(async () => deferred.promise);

  try {
    harness.modal.wireDomEvents();
    harness.modal.openSubsyncModal(BASE_PAYLOAD);

    harness.runButton.dispatch('click');
    await Promise.resolve();

    assert.equal(harness.ctx.state.subsyncModalOpen, false);
    assert.equal(harness.getNotifyClosedCalls(), 1);
    assert.equal(harness.getNotifyOpenedCalls(), 0);

    deferred.resolve({
      ok: false,
      message: 'alass synchronization failed: code=1 stderr: invalid subtitle format',
    });
    await flushMicrotasks();

    assert.equal(harness.ctx.state.subsyncModalOpen, true);
    assert.equal(
      harness.ctx.dom.subsyncStatus.textContent,
      'alass synchronization failed: code=1 stderr: invalid subtitle format',
    );
    assert.equal(harness.statusClassList.contains('error'), true);
    assert.equal(harness.ctx.dom.subsyncRunButton.disabled, false);
    assert.equal(harness.ctx.dom.subsyncEngineAlass.checked, true);
    assert.equal(harness.referenceSelect.value, '2');
    assert.equal(harness.targetSelect.value, '1');
    assert.equal(harness.getNotifyClosedCalls(), 1);
    assert.equal(harness.getNotifyOpenedCalls(), 1);
  } finally {
    harness.restoreGlobals();
  }
});

test('subsync modal disables ffsubsync when payload marks it unavailable', () => {
  const harness = createTestHarness(async () => ({ ok: true, message: 'ok' }));

  try {
    harness.modal.openSubsyncModal(
      payloadWith({ ffsubsyncAvailable: false, videoReferenceAvailable: false }),
    );

    assert.equal(harness.ctx.dom.subsyncEngineAlass.checked, true);
    assert.equal(harness.ctx.dom.subsyncEngineFfsubsync.checked, false);
    assert.equal(harness.ctx.dom.subsyncEngineFfsubsync.disabled, true);
    assert.equal(
      harness.ctx.dom.subsyncStatus.textContent,
      'Choose the alass reference and out-of-sync subtitle, then run.',
    );
  } finally {
    harness.restoreGlobals();
  }
});

test('subsync modal ignores enter submission when no sync engine is available', async () => {
  let runCalls = 0;
  const harness = createTestHarness(async () => {
    runCalls += 1;
    return { ok: true, message: 'ok' };
  });

  try {
    harness.modal.openSubsyncModal(
      payloadWith({
        subtitleTracks: [{ id: 1, label: 'External #1 - jpn (active)' }],
        defaultReferenceTrackId: null,
        videoReferenceAvailable: false,
        ffsubsyncAvailable: false,
      }),
    );

    harness.modal.handleSubsyncKeydown({
      key: 'Enter',
      preventDefault: () => {},
    } as KeyboardEvent);
    await flushMicrotasks();

    assert.equal(runCalls, 0);
    assert.equal(harness.ctx.state.subsyncModalOpen, true);
  } finally {
    harness.restoreGlobals();
  }
});

test('subsync modal defaults reference to the secondary track and target to the primary', async () => {
  let request: SubsyncManualRunRequest | null = null;
  const harness = createTestHarness(async (nextRequest) => {
    request = nextRequest;
    return { ok: true, message: 'ok' };
  });

  try {
    harness.modal.wireDomEvents();
    harness.modal.openSubsyncModal(BASE_PAYLOAD);

    assert.equal(harness.referenceSelect.value, '2');
    assert.equal(harness.targetSelect.value, '1');

    harness.runButton.dispatch('click');
    await flushMicrotasks();

    assert.deepEqual(request, {
      engine: 'alass',
      targetTrackId: 1,
      referenceMode: 'track',
      referenceTrackId: 2,
    });
  } finally {
    harness.restoreGlobals();
  }
});

test('subsync modal offers the video file as an alass reference and excludes the target track', () => {
  const harness = createTestHarness(async () => ({ ok: true, message: 'ok' }));

  try {
    harness.modal.wireDomEvents();
    harness.modal.openSubsyncModal(BASE_PAYLOAD);

    assert.deepEqual(
      harness.referenceSelect.options.map((option) => option.value),
      ['2', 'video'],
    );

    harness.targetSelect.value = '2';
    harness.targetSelect.dispatch('change');

    assert.deepEqual(
      harness.referenceSelect.options.map((option) => option.value),
      ['1', 'video'],
    );
    assert.equal(harness.referenceSelect.value, '1');
  } finally {
    harness.restoreGlobals();
  }
});

test('subsync modal sends the video reference mode when the video file is selected', async () => {
  let request: SubsyncManualRunRequest | null = null;
  const harness = createTestHarness(async (nextRequest) => {
    request = nextRequest;
    return { ok: true, message: 'ok' };
  });

  try {
    harness.modal.wireDomEvents();
    harness.modal.openSubsyncModal(BASE_PAYLOAD);

    harness.referenceSelect.value = 'video';
    harness.runButton.dispatch('click');
    await flushMicrotasks();

    assert.deepEqual(request, {
      engine: 'alass',
      targetTrackId: 1,
      referenceMode: 'video',
      referenceTrackId: null,
    });
  } finally {
    harness.restoreGlobals();
  }
});

test('subsync modal hides the reference picker for ffsubsync but keeps the target picker', () => {
  const harness = createTestHarness(async () => ({ ok: true, message: 'ok' }));

  try {
    harness.modal.wireDomEvents();
    harness.modal.openSubsyncModal(BASE_PAYLOAD);

    assert.equal(harness.referenceLabelClassList.contains('hidden'), false);

    harness.ctx.dom.subsyncEngineAlass.checked = false;
    harness.ctx.dom.subsyncEngineFfsubsync.checked = true;
    harness.ctx.dom.subsyncEngineFfsubsync.dispatch('change');

    assert.equal(harness.referenceLabelClassList.contains('hidden'), true);
    assert.equal(harness.targetSelect.value, '1');
  } finally {
    harness.restoreGlobals();
  }
});

test('guided subsync modal shows guidance banner, ranked candidates, browse file, and built-in options', async () => {
  let request: SubsyncManualRunRequest | null = null;
  const harness = createTestHarness(async (nextRequest) => {
    request = nextRequest;
    return { ok: true, message: 'ok' };
  });

  const guidanceClassList = createClassList();
  guidanceClassList.add('hidden');
  const optionsClassList = createClassList();
  const allowSplitsClassList = createClassList();
  const offsetLabelClassList = createClassList();
  offsetLabelClassList.add('hidden');
  const browseClassList = createClassList();
  browseClassList.add('hidden');

  const engineBuiltinEvents = createEventTarget();
  const engineOffsetEvents = createEventTarget();
  const browseEvents = createEventTarget();

  const subsyncEngineBuiltin = {
    checked: true,
    disabled: false,
    addEventListener: engineBuiltinEvents.addEventListener,
    dispatch: engineBuiltinEvents.dispatch,
  };
  const subsyncEngineOffset = {
    checked: false,
    disabled: false,
    addEventListener: engineOffsetEvents.addEventListener,
    dispatch: engineOffsetEvents.dispatch,
  };
  const subsyncBrowseReferenceButton = {
    classList: browseClassList,
    addEventListener: browseEvents.addEventListener,
    dispatch: browseEvents.dispatch,
  };
  const subsyncGuidance = {
    textContent: '',
    classList: guidanceClassList,
  };
  const subsyncStripSignsCheckbox = { checked: true };
  const subsyncAllowSplitsCheckbox = { checked: true };
  const subsyncOffsetInput = { value: '0' };

  Object.assign(harness.ctx.dom, {
    subsyncGuidance,
    subsyncEngineBuiltin,
    subsyncEngineOffset,
    subsyncBrowseReferenceButton,
    subsyncOptionsGroup: { classList: optionsClassList },
    subsyncStripSignsCheckbox,
    subsyncAllowSplitsLabel: { classList: allowSplitsClassList },
    subsyncAllowSplitsCheckbox,
    subsyncOffsetLabel: { classList: offsetLabelClassList },
    subsyncOffsetInput,
  });

  (
    globalThis as unknown as { window: { electronAPI: Record<string, unknown> } }
  ).window.electronAPI.pickSubsyncReferenceFile = async () => '/tmp/browsed-ref.srt';

  try {
    harness.modal.wireDomEvents();
    harness.modal.openSubsyncModal(
      payloadWith({
        mediaPath: '/tmp/ep01.mkv',
        guidanceMessage: 'Loaded English track appears +2.4s off audio.',
        defaultEngine: 'builtin',
        builtinAvailable: true,
        alassAvailable: false,
        canPickReferenceFile: true,
        referenceCandidates: [
          {
            key: 'track:2',
            kind: 'track',
            label: 'External #2 - eng (secondary)',
            detail: '+2.4s off audio',
            trackId: 2,
          },
          {
            key: 'video',
            kind: 'video',
            label: 'Video audio (built-in speech detection)',
            recommended: true,
          },
        ],
        defaultReferenceKey: 'video',
      }),
    );

    assert.equal(guidanceClassList.contains('hidden'), false);
    assert.equal(subsyncGuidance.textContent, 'Loaded English track appears +2.4s off audio.');
    assert.equal(subsyncEngineBuiltin.checked, true);
    assert.equal(harness.ctx.dom.subsyncEngineAlass.disabled, true);
    assert.equal(browseClassList.contains('hidden'), false);
    assert.equal(harness.referenceSelect.value, 'video');

    // Browse a custom file from disk and submit built-in sync with it.
    subsyncBrowseReferenceButton.dispatch('click');
    await flushMicrotasks();
    assert.equal(harness.referenceSelect.value, 'file:/tmp/browsed-ref.srt');

    subsyncAllowSplitsCheckbox.checked = false;
    harness.runButton.dispatch('click');
    await flushMicrotasks();

    assert.deepEqual(request, {
      engine: 'builtin',
      mediaPath: '/tmp/ep01.mkv',
      targetTrackId: 1,
      referenceMode: 'file',
      referenceTrackId: null,
      referenceFilePath: '/tmp/browsed-ref.srt',
      stripSigns: true,
      allowSplits: false,
    });
  } finally {
    harness.restoreGlobals();
  }
});
