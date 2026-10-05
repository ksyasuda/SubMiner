import assert from 'node:assert/strict';
import test from 'node:test';

import type { ResolvedControllerConfig } from '../../types';
import { createControllerConfigForm } from './controller-config-form.js';

type ActionId = keyof ResolvedControllerConfig['bindings'];

const BINDINGS: ResolvedControllerConfig['bindings'] = {
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
};

/** Minimal DOM element: enough for the form's createElement/appendChild/click usage. */
class FakeElement {
  className = '';
  textContent = '';
  title = '';
  type = '';
  children: FakeElement[] = [];
  private readonly listeners = new Map<string, Array<(event: Event) => void>>();
  private readonly attributes = new Map<string, string>();

  readonly classList = {
    add: (...tokens: string[]) => {
      this.className = [...new Set([...this.classes(), ...tokens])].join(' ');
    },
    contains: (token: string) => this.classes().includes(token),
  };

  set innerHTML(value: string) {
    if (value === '') this.children = [];
  }

  appendChild(child: FakeElement): FakeElement {
    this.children.push(child);
    return child;
  }

  addEventListener(type: string, listener: (event: Event) => void): void {
    this.listeners.set(type, [...(this.listeners.get(type) ?? []), listener]);
  }

  click(): void {
    const event = { stopPropagation: () => {}, preventDefault: () => {} } as unknown as Event;
    for (const listener of this.listeners.get('click') ?? []) listener(event);
  }

  setAttribute(name: string, value: string): void {
    this.attributes.set(name, value);
  }

  getAttribute(name: string): string | null {
    return this.attributes.get(name) ?? null;
  }

  private classes(): string[] {
    return this.className.split(' ').filter(Boolean);
  }
}

function findAll(root: FakeElement, className: string): FakeElement[] {
  return root.children.flatMap((child) => [
    ...(child.classList.contains(className) ? [child] : []),
    ...findAll(child, className),
  ]);
}

function findOne(root: FakeElement, className: string, text?: string): FakeElement {
  const match = findAll(root, className).find(
    (el) => text === undefined || el.textContent === text,
  );
  assert.ok(match, `missing .${className}${text === undefined ? '' : ` "${text}"`}`);
  return match;
}

/** Renders the form into a fake container with `document` stubbed for the duration. */
function renderForm(learningActionId: ActionId | null) {
  const previousDocument = Object.getOwnPropertyDescriptor(globalThis, 'document');
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: { createElement: () => new FakeElement() },
  });

  try {
    const calls: string[] = [];
    const container = new FakeElement();
    createControllerConfigForm({
      container: container as unknown as HTMLElement,
      getBindings: () => BINDINGS,
      getLearningActionId: () => learningActionId,
      getDpadLearningActionId: () => null,
      onLearn: (actionId, bindingType) => calls.push(`learn:${actionId}:${bindingType}`),
      onClear: (actionId) => calls.push(`clear:${actionId}`),
      onReset: (actionId) => calls.push(`reset:${actionId}`),
      onDpadLearn: (actionId) => calls.push(`dpadLearn:${actionId}`),
      onDpadClear: (actionId) => calls.push(`dpadClear:${actionId}`),
      onDpadReset: (actionId) => calls.push(`dpadReset:${actionId}`),
    }).render();

    const row = (label: string) => {
      const match = findAll(container, 'controller-config-row').find(
        (el) => findOne(el, 'controller-config-label').textContent === label,
      );
      assert.ok(match, `missing row "${label}"`);
      return match;
    };
    return { calls, container, row };
  } finally {
    if (previousDocument) {
      Object.defineProperty(globalThis, 'document', previousDocument);
    } else {
      Reflect.deleteProperty(globalThis, 'document');
    }
  }
}

test('controller config form renders grouped rows with friendly binding labels', () => {
  const { container } = renderForm(null);

  assert.deepEqual(
    findAll(container, 'controller-config-group').map((el) => el.textContent),
    ['Lookup', 'Playback', 'Popup Audio', 'Navigation'],
  );
  assert.deepEqual(
    findAll(container, 'controller-config-row').map((row) => [
      findOne(row, 'controller-config-label').textContent,
      findOne(row, 'controller-config-badge').textContent,
    ]),
    [
      ['Toggle Lookup', 'A / Cross'],
      ['Close Lookup', 'B / Circle'],
      ['Mine Card', 'X / Square'],
      ['Toggle Keyboard-Only Mode', 'Y / Triangle'],
      ['Toggle MPV Pause', 'R3 / RS'],
      ['Quit MPV', 'Back / Select'],
      ['Previous Audio', 'None'],
      ['Next Audio', 'RB / R1'],
      ['Play Current Audio', 'LB / L1'],
      ['Token Move (Stick)', 'Left Stick X'],
      ['Token Move (D-pad)', 'D-pad ↔'],
      ['Popup Scroll (Stick)', 'Left Stick Y'],
      ['Popup Scroll (D-pad)', 'D-pad ↕'],
      ['Alt Horizontal (Stick)', 'Right Stick X'],
      ['Alt Horizontal (D-pad)', 'None'],
      ['Popup Jump (Stick)', 'Right Stick Y'],
      ['Popup Jump (D-pad)', 'None'],
    ],
  );
  assert.deepEqual(
    findAll(container, 'disabled').map((badge) => badge.textContent),
    ['None', 'None', 'None'],
  );
  assert.deepEqual(findAll(container, 'expanded'), []);
  assert.deepEqual(findAll(container, 'controller-config-edit-panel'), []);
});

test('controller config form expands the learning row with a listening edit panel', () => {
  const { container, row } = renderForm('toggleLookup');

  assert.deepEqual(
    findAll(container, 'expanded').map((el) => findOne(el, 'controller-config-label').textContent),
    ['Toggle Lookup'],
  );
  // The edit panel is inserted directly after the expanded row.
  const rowIndex = container.children.indexOf(row('Toggle Lookup'));
  const panel = container.children[rowIndex + 1];
  assert.ok(panel);
  assert.equal(panel.classList.contains('controller-config-edit-panel'), true);

  const hint = findOne(panel, 'controller-config-edit-hint');
  assert.equal(hint.textContent, 'Press a button, trigger, or move a stick…');
  assert.equal(hint.classList.contains('learning'), true);
  assert.equal(findOne(panel, 'btn-learn').textContent, 'Listening…');
});

type Locate = (row: FakeElement, panel: FakeElement) => FakeElement;

for (const c of [
  {
    control: 'row badge',
    locate: (row) => findOne(row, 'controller-config-badge'),
    expected: 'learn:toggleLookup:discrete',
  },
  {
    control: 'row reset icon',
    locate: (row) => findOne(row, 'controller-config-reset-icon'),
    expected: 'reset:toggleLookup',
  },
  {
    control: 'row edit icon',
    locate: (row) => findOne(row, 'controller-config-edit-icon'),
    expected: 'learn:toggleLookup:discrete',
  },
  {
    control: 'panel Learn',
    locate: (_row, panel) => findOne(panel, 'btn-learn'),
    expected: 'learn:toggleLookup:discrete',
  },
  {
    control: 'panel Clear',
    locate: (_row, panel) => findOne(panel, 'btn-secondary', 'Clear'),
    expected: 'clear:toggleLookup',
  },
  {
    control: 'panel Reset',
    locate: (_row, panel) => findOne(panel, 'btn-secondary', 'Reset'),
    expected: 'reset:toggleLookup',
  },
] satisfies Array<{ control: string; locate: Locate; expected: string }>) {
  test(`controller config form ${c.control} dispatches ${c.expected}`, () => {
    const { calls, container, row } = renderForm('toggleLookup');

    c.locate(row('Toggle Lookup'), findOne(container, 'controller-config-edit-panel')).click();

    assert.deepEqual(calls, [c.expected]);
  });
}
