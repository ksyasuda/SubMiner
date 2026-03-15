import type {
  ControllerDpadFallback,
  ResolvedControllerAxisBinding,
  ResolvedControllerConfig,
  ResolvedControllerDiscreteBinding,
} from '../../types';

type ControllerBindingActionId = keyof ResolvedControllerConfig['bindings'];

type ControllerBindingDefinition = {
  id: ControllerBindingActionId;
  label: string;
  group: string;
  bindingType: 'discrete' | 'axis';
  defaultBinding: ResolvedControllerConfig['bindings'][ControllerBindingActionId];
};

export const CONTROLLER_BINDING_DEFINITIONS: ControllerBindingDefinition[] = [
  {
    id: 'toggleLookup',
    label: 'Toggle Lookup',
    group: 'Lookup',
    bindingType: 'discrete',
    defaultBinding: { kind: 'button', buttonIndex: 0 },
  },
  {
    id: 'closeLookup',
    label: 'Close Lookup',
    group: 'Lookup',
    bindingType: 'discrete',
    defaultBinding: { kind: 'button', buttonIndex: 1 },
  },
  {
    id: 'mineCard',
    label: 'Mine Card',
    group: 'Lookup',
    bindingType: 'discrete',
    defaultBinding: { kind: 'button', buttonIndex: 2 },
  },
  {
    id: 'toggleKeyboardOnlyMode',
    label: 'Toggle Keyboard-Only Mode',
    group: 'Playback',
    bindingType: 'discrete',
    defaultBinding: { kind: 'button', buttonIndex: 3 },
  },
  {
    id: 'toggleMpvPause',
    label: 'Toggle MPV Pause',
    group: 'Playback',
    bindingType: 'discrete',
    defaultBinding: { kind: 'button', buttonIndex: 9 },
  },
  {
    id: 'quitMpv',
    label: 'Quit MPV',
    group: 'Playback',
    bindingType: 'discrete',
    defaultBinding: { kind: 'button', buttonIndex: 6 },
  },
  {
    id: 'previousAudio',
    label: 'Previous Audio',
    group: 'Popup Audio',
    bindingType: 'discrete',
    defaultBinding: { kind: 'none' },
  },
  {
    id: 'nextAudio',
    label: 'Next Audio',
    group: 'Popup Audio',
    bindingType: 'discrete',
    defaultBinding: { kind: 'button', buttonIndex: 5 },
  },
  {
    id: 'playCurrentAudio',
    label: 'Play Current Audio',
    group: 'Popup Audio',
    bindingType: 'discrete',
    defaultBinding: { kind: 'button', buttonIndex: 4 },
  },
  {
    id: 'leftStickHorizontal',
    label: 'Token Move',
    group: 'Navigation',
    bindingType: 'axis',
    defaultBinding: { kind: 'axis', axisIndex: 0, dpadFallback: 'horizontal' },
  },
  {
    id: 'leftStickVertical',
    label: 'Popup Scroll',
    group: 'Navigation',
    bindingType: 'axis',
    defaultBinding: { kind: 'axis', axisIndex: 1, dpadFallback: 'vertical' },
  },
  {
    id: 'rightStickHorizontal',
    label: 'Alt Horizontal',
    group: 'Navigation',
    bindingType: 'axis',
    defaultBinding: { kind: 'axis', axisIndex: 3, dpadFallback: 'none' },
  },
  {
    id: 'rightStickVertical',
    label: 'Popup Jump',
    group: 'Navigation',
    bindingType: 'axis',
    defaultBinding: { kind: 'axis', axisIndex: 4, dpadFallback: 'none' },
  },
];

export function getControllerBindingDefinition(actionId: ControllerBindingActionId) {
  return CONTROLLER_BINDING_DEFINITIONS.find((definition) => definition.id === actionId) ?? null;
}

export function getDefaultControllerBinding(actionId: ControllerBindingActionId) {
  const definition = getControllerBindingDefinition(actionId);
  if (!definition) {
    return { kind: 'none' } as const;
  }
  return JSON.parse(JSON.stringify(definition.defaultBinding)) as ResolvedControllerConfig['bindings'][ControllerBindingActionId];
}

export function getDefaultDpadFallback(actionId: ControllerBindingActionId): ControllerDpadFallback {
  const definition = getControllerBindingDefinition(actionId);
  if (!definition || definition.defaultBinding.kind !== 'axis') return 'none';
  const binding = definition.defaultBinding;
  return 'dpadFallback' in binding && binding.dpadFallback ? binding.dpadFallback : 'none';
}

const STANDARD_BUTTON_NAMES: Record<number, string> = {
  0: 'A / Cross',
  1: 'B / Circle',
  2: 'X / Square',
  3: 'Y / Triangle',
  4: 'LB / L1',
  5: 'RB / R1',
  6: 'Back / Select',
  7: 'Start / Options',
  8: 'L3 / LS',
  9: 'R3 / RS',
  10: 'Left Stick Click',
  11: 'Right Stick Click',
  12: 'D-pad Up',
  13: 'D-pad Down',
  14: 'D-pad Left',
  15: 'D-pad Right',
  16: 'Guide / Home',
};

const STANDARD_AXIS_NAMES: Record<number, string> = {
  0: 'Left Stick X',
  1: 'Left Stick Y',
  2: 'Left Trigger',
  3: 'Right Stick X',
  4: 'Right Stick Y',
  5: 'Right Trigger',
};

const DPAD_FALLBACK_LABELS: Record<ControllerDpadFallback, string> = {
  none: 'None',
  horizontal: 'D-pad \u2194',
  vertical: 'D-pad \u2195',
};

function getFriendlyButtonName(buttonIndex: number): string {
  return STANDARD_BUTTON_NAMES[buttonIndex] ?? `Button ${buttonIndex}`;
}

function getFriendlyAxisName(axisIndex: number): string {
  return STANDARD_AXIS_NAMES[axisIndex] ?? `Axis ${axisIndex}`;
}

export function formatControllerBindingSummary(
  binding: ResolvedControllerDiscreteBinding | ResolvedControllerAxisBinding,
): string {
  if (binding.kind === 'none') {
    return 'Disabled';
  }
  if ('direction' in binding) {
    return `Axis ${binding.axisIndex} ${binding.direction === 'positive' ? '+' : '-'}`;
  }
  if ('buttonIndex' in binding) {
    return `Button ${binding.buttonIndex}`;
  }
  if (binding.dpadFallback === 'none') {
    return `Axis ${binding.axisIndex}`;
  }
  return `Axis ${binding.axisIndex} + D-pad ${binding.dpadFallback}`;
}

function formatFriendlyStickLabel(binding: ResolvedControllerAxisBinding): string {
  if (binding.kind === 'none') return 'None';
  return getFriendlyAxisName(binding.axisIndex);
}

function formatFriendlyBindingLabel(
  binding: ResolvedControllerDiscreteBinding | ResolvedControllerAxisBinding,
): string {
  if (binding.kind === 'none') return 'None';
  if ('direction' in binding) {
    const name = getFriendlyAxisName(binding.axisIndex);
    return `${name} ${binding.direction === 'positive' ? '+' : '\u2212'}`;
  }
  if ('buttonIndex' in binding) return getFriendlyButtonName(binding.buttonIndex);
  return getFriendlyAxisName(binding.axisIndex);
}

/** Unique key for expanded rows. Stick rows use the action id, dpad rows append ':dpad'. */
type ExpandedRowKey = string;

export function createControllerConfigForm(options: {
  container: HTMLElement;
  getBindings: () => ResolvedControllerConfig['bindings'];
  getLearningActionId: () => ControllerBindingActionId | null;
  getDpadLearningActionId: () => ControllerBindingActionId | null;
  onLearn: (actionId: ControllerBindingActionId, bindingType: 'discrete' | 'axis') => void;
  onClear: (actionId: ControllerBindingActionId) => void;
  onReset: (actionId: ControllerBindingActionId) => void;
  onDpadLearn: (actionId: ControllerBindingActionId) => void;
  onDpadClear: (actionId: ControllerBindingActionId) => void;
  onDpadReset: (actionId: ControllerBindingActionId) => void;
}) {
  let expandedRowKey: ExpandedRowKey | null = null;

  function render(): void {
    options.container.innerHTML = '';
    let lastGroup = '';
    const learningActionId = options.getLearningActionId();
    const dpadLearningActionId = options.getDpadLearningActionId();

    // Auto-expand when learning starts
    if (learningActionId) {
      expandedRowKey = learningActionId;
    } else if (dpadLearningActionId) {
      expandedRowKey = `${dpadLearningActionId}:dpad`;
    }

    for (const definition of CONTROLLER_BINDING_DEFINITIONS) {
      if (definition.group !== lastGroup) {
        const header = document.createElement('div');
        header.className = 'controller-config-group';
        header.textContent = definition.group;
        options.container.appendChild(header);
        lastGroup = definition.group;
      }

      const binding = options.getBindings()[definition.id];

      if (definition.bindingType === 'axis') {
        renderAxisStickRow(definition, binding as ResolvedControllerAxisBinding, learningActionId);
        renderAxisDpadRow(definition, binding as ResolvedControllerAxisBinding, dpadLearningActionId);
      } else {
        renderDiscreteRow(definition, binding, learningActionId);
      }
    }
  }

  function renderDiscreteRow(
    definition: ControllerBindingDefinition,
    binding: ResolvedControllerConfig['bindings'][ControllerBindingActionId],
    learningActionId: ControllerBindingActionId | null,
  ): void {
    const rowKey = definition.id as string;
    const isExpanded = expandedRowKey === rowKey;
    const isLearning = learningActionId === definition.id;

    const row = createRow(definition.label, formatFriendlyBindingLabel(binding), binding.kind === 'none', isExpanded);
    row.addEventListener('click', () => {
      expandedRowKey = expandedRowKey === rowKey ? null : rowKey;
      render();
    });
    options.container.appendChild(row);

    if (isExpanded) {
      const hint = isLearning
        ? 'Press a button, trigger, or move a stick\u2026'
        : `Currently: ${formatControllerBindingSummary(binding)}`;
      const panel = createEditPanel(hint, isLearning, {
        onLearn: (e) => { e.stopPropagation(); options.onLearn(definition.id, definition.bindingType); },
        onClear: (e) => { e.stopPropagation(); options.onClear(definition.id); },
        onReset: (e) => { e.stopPropagation(); options.onReset(definition.id); },
      });
      options.container.appendChild(panel);
    }
  }

  function renderAxisStickRow(
    definition: ControllerBindingDefinition,
    binding: ResolvedControllerAxisBinding,
    learningActionId: ControllerBindingActionId | null,
  ): void {
    const rowKey = definition.id as string;
    const isExpanded = expandedRowKey === rowKey;
    const isLearning = learningActionId === definition.id;

    const row = createRow(`${definition.label} (Stick)`, formatFriendlyStickLabel(binding), binding.kind === 'none', isExpanded);
    row.addEventListener('click', () => {
      expandedRowKey = expandedRowKey === rowKey ? null : rowKey;
      render();
    });
    options.container.appendChild(row);

    if (isExpanded) {
      const summary = binding.kind === 'none' ? 'Disabled' : `Axis ${binding.axisIndex}`;
      const hint = isLearning ? 'Move a stick or trigger\u2026' : `Currently: ${summary}`;
      const panel = createEditPanel(hint, isLearning, {
        onLearn: (e) => { e.stopPropagation(); options.onLearn(definition.id, 'axis'); },
        onClear: (e) => { e.stopPropagation(); options.onClear(definition.id); },
        onReset: (e) => { e.stopPropagation(); options.onReset(definition.id); },
      });
      options.container.appendChild(panel);
    }
  }

  function renderAxisDpadRow(
    definition: ControllerBindingDefinition,
    binding: ResolvedControllerAxisBinding,
    dpadLearningActionId: ControllerBindingActionId | null,
  ): void {
    const rowKey = `${definition.id as string}:dpad`;
    const isExpanded = expandedRowKey === rowKey;
    const isLearning = dpadLearningActionId === definition.id;

    const dpadFallback: ControllerDpadFallback = binding.kind === 'none' ? 'none' : binding.dpadFallback;
    const badgeText = DPAD_FALLBACK_LABELS[dpadFallback];
    const row = createRow(`${definition.label} (D-pad)`, badgeText, dpadFallback === 'none', isExpanded);
    row.addEventListener('click', () => {
      expandedRowKey = expandedRowKey === rowKey ? null : rowKey;
      render();
    });
    options.container.appendChild(row);

    if (isExpanded) {
      const hint = isLearning
        ? 'Press a D-pad direction\u2026'
        : `Currently: ${DPAD_FALLBACK_LABELS[dpadFallback]}`;
      const panel = createEditPanel(hint, isLearning, {
        onLearn: (e) => { e.stopPropagation(); options.onDpadLearn(definition.id); },
        onClear: (e) => { e.stopPropagation(); options.onDpadClear(definition.id); },
        onReset: (e) => { e.stopPropagation(); options.onDpadReset(definition.id); },
      });
      options.container.appendChild(panel);
    }
  }

  function createRow(labelText: string, badgeText: string, isDisabled: boolean, isExpanded: boolean): HTMLDivElement {
    const row = document.createElement('div');
    row.className = 'controller-config-row';
    if (isExpanded) row.classList.add('expanded');

    const label = document.createElement('div');
    label.className = 'controller-config-label';
    label.textContent = labelText;

    const right = document.createElement('div');
    right.className = 'controller-config-right';

    const badge = document.createElement('span');
    badge.className = 'controller-config-badge';
    if (isDisabled) badge.classList.add('disabled');
    badge.textContent = badgeText;

    const editIcon = document.createElement('span');
    editIcon.className = 'controller-config-edit-icon';
    editIcon.textContent = '\u270E';

    right.appendChild(badge);
    right.appendChild(editIcon);
    row.appendChild(label);
    row.appendChild(right);

    return row;
  }

  function createEditPanel(
    hintText: string,
    isLearning: boolean,
    callbacks: {
      onLearn: (e: Event) => void;
      onClear: (e: Event) => void;
      onReset: (e: Event) => void;
    },
  ): HTMLDivElement {
    const panel = document.createElement('div');
    panel.className = 'controller-config-edit-panel';

    const inner = document.createElement('div');
    inner.className = 'controller-config-edit-inner';

    const hint = document.createElement('div');
    hint.className = 'controller-config-edit-hint';
    if (isLearning) hint.classList.add('learning');
    hint.textContent = hintText;

    const actions = document.createElement('div');
    actions.className = 'controller-config-edit-actions';

    const learnButton = document.createElement('button');
    learnButton.type = 'button';
    learnButton.className = isLearning ? 'btn-learn active' : 'btn-learn';
    learnButton.textContent = isLearning ? 'Listening\u2026' : 'Learn';
    learnButton.addEventListener('click', callbacks.onLearn);

    const clearButton = document.createElement('button');
    clearButton.type = 'button';
    clearButton.className = 'btn-secondary';
    clearButton.textContent = 'Clear';
    clearButton.addEventListener('click', callbacks.onClear);

    const resetButton = document.createElement('button');
    resetButton.type = 'button';
    resetButton.className = 'btn-secondary';
    resetButton.textContent = 'Reset';
    resetButton.addEventListener('click', callbacks.onReset);

    actions.appendChild(learnButton);
    actions.appendChild(clearButton);
    actions.appendChild(resetButton);

    inner.appendChild(hint);
    inner.appendChild(actions);
    panel.appendChild(inner);

    return panel;
  }

  return { render };
}
