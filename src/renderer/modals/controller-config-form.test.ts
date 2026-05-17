import assert from 'node:assert/strict';
import test from 'node:test';

import { createControllerConfigForm } from './controller-config-form.js';

function createClassList(initialTokens: string[] = []) {
  const tokens = new Set(initialTokens);
  return {
    add: (...entries: string[]) => {
      for (const entry of entries) tokens.add(entry);
    },
    remove: (...entries: string[]) => {
      for (const entry of entries) tokens.delete(entry);
    },
    toggle: (entry: string, force?: boolean) => {
      if (force === undefined) {
        if (tokens.has(entry)) tokens.delete(entry);
        else tokens.add(entry);
        return tokens.has(entry);
      }
      if (force) tokens.add(entry);
      else tokens.delete(entry);
      return force;
    },
    contains: (entry: string) => tokens.has(entry),
  };
}

function createFakeElement() {
  const attributes = new Map<string, string>();
  const el = {
    className: '',
    textContent: '',
    _innerHTML: '',
    value: '',
    disabled: false,
    selected: false,
    type: '',
    children: [] as any[],
    listeners: new Map<string, Array<(e?: any) => void>>(),
    classList: createClassList(),
    appendChild(child: any) {
      this.children.push(child);
      return child;
    },
    addEventListener(type: string, listener: (e?: any) => void) {
      const existing = this.listeners.get(type) ?? [];
      existing.push(listener);
      this.listeners.set(type, existing);
    },
    dispatch(type: string) {
      const fakeEvent = { stopPropagation: () => {}, preventDefault: () => {} };
      for (const listener of this.listeners.get(type) ?? []) {
        listener(fakeEvent);
      }
    },
    setAttribute(name: string, value: string) {
      attributes.set(name, value);
    },
    getAttribute(name: string) {
      return attributes.get(name) ?? null;
    },
  };
  Object.defineProperty(el, 'innerHTML', {
    get() {
      return el._innerHTML;
    },
    set(v: string) {
      el._innerHTML = v;
      if (v === '') el.children.length = 0;
    },
  });
  return el;
}

test('controller config form renders rows and dispatches learn clear reset callbacks', () => {
  const previousDocumentDescriptor = Object.getOwnPropertyDescriptor(globalThis, 'document');
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createFakeElement(),
    },
  });

  try {
    const calls: string[] = [];
    const container = createFakeElement();
    const form = createControllerConfigForm({
      container: container as never,
      getBindings: () =>
        ({
          toggleLookup: { kind: 'button', buttonIndex: 0 },
          closeLookup: { kind: 'button', buttonIndex: 1 },
          toggleKeyboardOnlyMode: { kind: 'button', buttonIndex: 3 },
          mineCard: { kind: 'button', buttonIndex: 2 },
          quitMpv: { kind: 'button', buttonIndex: 6 },
          previousAudio: { kind: 'none' },
          nextAudio: { kind: 'button', buttonIndex: 5 },
          playCurrentAudio: { kind: 'button', buttonIndex: 4 },
          toggleMpvPause: { kind: 'button', buttonIndex: 9 },
          leftStickHorizontal: { kind: 'axis', axisIndex: 0, dpadFallback: 'horizontal' },
          leftStickVertical: { kind: 'axis', axisIndex: 1, dpadFallback: 'vertical' },
          rightStickHorizontal: { kind: 'axis', axisIndex: 3, dpadFallback: 'none' },
          rightStickVertical: { kind: 'axis', axisIndex: 4, dpadFallback: 'none' },
        }) as never,
      getLearningActionId: () => 'toggleLookup',
      getDpadLearningActionId: () => null,
      onLearn: (actionId, bindingType) => calls.push(`learn:${actionId}:${bindingType}`),
      onClear: (actionId) => calls.push(`clear:${actionId}`),
      onReset: (actionId) => calls.push(`reset:${actionId}`),
      onDpadLearn: (actionId) => calls.push(`dpadLearn:${actionId}`),
      onDpadClear: (actionId) => calls.push(`dpadClear:${actionId}`),
      onDpadReset: (actionId) => calls.push(`dpadReset:${actionId}`),
    });

    form.render();

    // In the new compact list layout, children are:
    // [0] group header, [1] first binding row (auto-expanded because learning), [2] edit panel, [3] next row, ...
    const firstRow = container.children[1];
    assert.equal(firstRow.classList.contains('expanded'), true);

    // After expanding, the edit panel is inserted after the row:
    // [0] group header, [1] row, [2] edit panel, [3] next row, ...
    const editPanel = container.children[2];
    // editPanel > inner > actions > learnButton
    const inner = editPanel.children[0];
    const actions = inner.children[1];
    const learnButton = actions.children[0];
    learnButton.dispatch('click');
    actions.children[1].dispatch('click');
    actions.children[2].dispatch('click');

    assert.deepEqual(calls, [
      'learn:toggleLookup:discrete',
      'clear:toggleLookup',
      'reset:toggleLookup',
    ]);
  } finally {
    if (previousDocumentDescriptor) {
      Object.defineProperty(globalThis, 'document', previousDocumentDescriptor);
    } else {
      Reflect.deleteProperty(globalThis, 'document');
    }
  }
});

test('controller config form starts learn from badge or edit and resets from row button', () => {
  const previousDocumentDescriptor = Object.getOwnPropertyDescriptor(globalThis, 'document');
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createFakeElement(),
    },
  });

  try {
    const calls: string[] = [];
    const container = createFakeElement();
    const form = createControllerConfigForm({
      container: container as never,
      getBindings: () =>
        ({
          toggleLookup: { kind: 'button', buttonIndex: 0 },
          closeLookup: { kind: 'button', buttonIndex: 1 },
          toggleKeyboardOnlyMode: { kind: 'button', buttonIndex: 3 },
          mineCard: { kind: 'button', buttonIndex: 2 },
          quitMpv: { kind: 'button', buttonIndex: 6 },
          previousAudio: { kind: 'none' },
          nextAudio: { kind: 'button', buttonIndex: 5 },
          playCurrentAudio: { kind: 'button', buttonIndex: 4 },
          toggleMpvPause: { kind: 'button', buttonIndex: 9 },
          leftStickHorizontal: { kind: 'axis', axisIndex: 0, dpadFallback: 'horizontal' },
          leftStickVertical: { kind: 'axis', axisIndex: 1, dpadFallback: 'vertical' },
          rightStickHorizontal: { kind: 'axis', axisIndex: 3, dpadFallback: 'none' },
          rightStickVertical: { kind: 'axis', axisIndex: 4, dpadFallback: 'none' },
        }) as never,
      getLearningActionId: () => null,
      getDpadLearningActionId: () => null,
      onLearn: (actionId, bindingType) => calls.push(`learn:${actionId}:${bindingType}`),
      onClear: (actionId) => calls.push(`clear:${actionId}`),
      onReset: (actionId) => calls.push(`reset:${actionId}`),
      onDpadLearn: (actionId) => calls.push(`dpadLearn:${actionId}`),
      onDpadClear: (actionId) => calls.push(`dpadClear:${actionId}`),
      onDpadReset: (actionId) => calls.push(`dpadReset:${actionId}`),
    });

    form.render();

    const firstRow = container.children[1];
    const right = firstRow.children[1];
    const badge = right.children[0];
    const resetButton = right.children[1];
    const editButton = right.children[2];

    badge.dispatch('click');
    resetButton.dispatch('click');
    editButton.dispatch('click');

    assert.deepEqual(calls, [
      'learn:toggleLookup:discrete',
      'reset:toggleLookup',
      'learn:toggleLookup:discrete',
    ]);
  } finally {
    if (previousDocumentDescriptor) {
      Object.defineProperty(globalThis, 'document', previousDocumentDescriptor);
    } else {
      Reflect.deleteProperty(globalThis, 'document');
    }
  }
});
