import type {
  ControllerDpadFallback,
  ResolvedControllerAxisBinding,
  ResolvedControllerConfig,
  ResolvedControllerDiscreteBinding,
} from '../../types';
import { i18n } from '../../i18n/index.js';

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
    label: i18n.t('controller.toggleLookup'),
    group: i18n.t('controller.lookupGroup'),
    bindingType: 'discrete',
    defaultBinding: { kind: 'button', buttonIndex: 0 },
  },
  {
    id: 'closeLookup',
    label: i18n.t('controller.closeLookup'),
    group: i18n.t('controller.lookupGroup'),
    bindingType: 'discrete',
    defaultBinding: { kind: 'button', buttonIndex: 1 },
  },
  {
    id: 'mineCard',
    label: i18n.t('controller.mineCard'),
    group: i18n.t('controller.lookupGroup'),
    bindingType: 'discrete',
    defaultBinding: { kind: 'button', buttonIndex: 2 },
  },
  {
    id: 'toggleKeyboardOnlyMode',
    label: i18n.t('controller.keyboardOnlyMode'),
    group: i18n.t('controller.playbackGroup'),
    bindingType: 'discrete',
    defaultBinding: { kind: 'button', buttonIndex: 3 },
  },
  {
    id: 'toggleMpvPause',
    label: i18n.t('controller.toggleMpvPause'),
    group: i18n.t('controller.playbackGroup'),
    bindingType: 'discrete',
    defaultBinding: { kind: 'button', buttonIndex: 9 },
  },
  {
    id: 'quitMpv',
    label: i18n.t('controller.quitMpv'),
    group: i18n.t('controller.playbackGroup'),
    bindingType: 'discrete',
    defaultBinding: { kind: 'button', buttonIndex: 6 },
  },
  {
    id: 'previousAudio',
    label: i18n.t('controller.previousAudio'),
    group: i18n.t('controller.popupAudioGroup'),
    bindingType: 'discrete',
    defaultBinding: { kind: 'none' },
  },
  {
    id: 'nextAudio',
    label: i18n.t('controller.nextAudio'),
    group: i18n.t('controller.popupAudioGroup'),
    bindingType: 'discrete',
    defaultBinding: { kind: 'button', buttonIndex: 5 },
  },
  {
    id: 'playCurrentAudio',
    label: i18n.t('controller.playCurrentAudio'),
    group: i18n.t('controller.popupAudioGroup'),
    bindingType: 'discrete',
    defaultBinding: { kind: 'button', buttonIndex: 4 },
  },
  {
    id: 'leftStickHorizontal',
    label: i18n.t('controller.leftStickHorizontal'),
    group: i18n.t('controller.navigationGroup'),
    bindingType: 'axis',
    defaultBinding: { kind: 'axis', axisIndex: 0, dpadFallback: 'horizontal' },
  },
  {
    id: 'leftStickVertical',
    label: i18n.t('controller.leftStickVertical'),
    group: i18n.t('controller.navigationGroup'),
    bindingType: 'axis',
    defaultBinding: { kind: 'axis', axisIndex: 1, dpadFallback: 'vertical' },
  },
  {
    id: 'rightStickHorizontal',
    label: i18n.t('controller.rightStickHorizontal'),
    group: i18n.t('controller.navigationGroup'),
    bindingType: 'axis',
    defaultBinding: { kind: 'axis', axisIndex: 3, dpadFallback: 'none' },
  },
  {
    id: 'rightStickVertical',
    label: i18n.t('controller.rightStickVertical'),
    group: i18n.t('controller.navigationGroup'),
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
  return JSON.parse(
    JSON.stringify(definition.defaultBinding),
  ) as ResolvedControllerConfig['bindings'][ControllerBindingActionId];
}

export function getDefaultDpadFallback(
  actionId: ControllerBindingActionId,
): ControllerDpadFallback {
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

const STANDARD_BUTTON_NAME_KEYS: Record<number, string> = {
  0: 'controller.standardButton.cross',
  1: 'controller.standardButton.circle',
  2: 'controller.standardButton.square',
  3: 'controller.standardButton.triangle',
  4: 'controller.standardButton.lbL1',
  5: 'controller.standardButton.rbR1',
  6: 'controller.standardButton.backSelect',
  7: 'controller.standardButton.startOptions',
  8: 'controller.standardButton.l3',
  9: 'controller.standardButton.r3',
  10: 'controller.standardButton.leftStickClick',
  11: 'controller.standardButton.rightStickClick',
  12: 'controller.standardButton.dpadUp',
  13: 'controller.standardButton.dpadDown',
  14: 'controller.standardButton.dpadLeft',
  15: 'controller.standardButton.dpadRight',
  16: 'controller.standardButton.guideHome',
};

const STANDARD_AXIS_NAMES: Record<number, string> = {
  0: 'Left Stick X',
  1: 'Left Stick Y',
  2: 'Left Trigger',
  3: 'Right Stick X',
  4: 'Right Stick Y',
  5: 'Right Trigger',
};

const STANDARD_AXIS_NAME_KEYS: Record<number, string> = {
  0: 'controller.standardAxis.leftStickX',
  1: 'controller.standardAxis.leftStickY',
  2: 'controller.standardAxis.leftTrigger',
  3: 'controller.standardAxis.rightStickX',
  4: 'controller.standardAxis.rightStickY',
  5: 'controller.standardAxis.rightTrigger',
};

const DPAD_FALLBACK_LABELS: Record<ControllerDpadFallback, string> = {
  none: i18n.t('controller.none'),
  horizontal: i18n.t('controller.dpadHorizontal'),
  vertical: i18n.t('controller.dpadVertical'),
};

function getFriendlyButtonName(buttonIndex: number): string {
  const key = STANDARD_BUTTON_NAME_KEYS[buttonIndex];
  if (key) {
    return i18n.t(key, undefined, STANDARD_BUTTON_NAMES[buttonIndex]);
  }
  return i18n.t('controller.buttonIndex', { index: buttonIndex });
}

function getFriendlyAxisName(axisIndex: number): string {
  const key = STANDARD_AXIS_NAME_KEYS[axisIndex];
  if (key) {
    return i18n.t(key, undefined, STANDARD_AXIS_NAMES[axisIndex]);
  }
  return i18n.t('controller.axisIndex', { index: axisIndex });
}

export function formatControllerBindingSummary(
  binding: ResolvedControllerDiscreteBinding | ResolvedControllerAxisBinding,
): string {
  if (binding.kind === 'none') {
    return i18n.t('controller.disabled');
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
  if (binding.kind === 'none') return i18n.t('controller.none');
  return getFriendlyAxisName(binding.axisIndex);
}

function formatFriendlyBindingLabel(
  binding: ResolvedControllerDiscreteBinding | ResolvedControllerAxisBinding,
): string {
  if (binding.kind === 'none') return i18n.t('controller.none');
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
        renderAxisDpadRow(
          definition,
          binding as ResolvedControllerAxisBinding,
          dpadLearningActionId,
        );
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

    const row = createRow(
      definition.label,
      formatFriendlyBindingLabel(binding),
      binding.kind === 'none',
      isExpanded,
      i18n.t('controller.learnLabel', { label: definition.label }),
      (e) => {
        e.stopPropagation();
        expandedRowKey = rowKey;
        options.onLearn(definition.id, definition.bindingType);
      },
      i18n.t('controller.resetLabel', { label: definition.label }),
      (e) => {
        e.stopPropagation();
        options.onReset(definition.id);
      },
    );
    row.addEventListener('click', () => {
      expandedRowKey = expandedRowKey === rowKey ? null : rowKey;
      render();
    });
    options.container.appendChild(row);

    if (isExpanded) {
      const hint = isLearning
        ? i18n.t('controller.learnPrompt')
        : i18n.t('controller.currently', { binding: formatControllerBindingSummary(binding) });
      const panel = createEditPanel(hint, isLearning, {
        onLearn: (e) => {
          e.stopPropagation();
          options.onLearn(definition.id, definition.bindingType);
        },
        onClear: (e) => {
          e.stopPropagation();
          options.onClear(definition.id);
        },
        onReset: (e) => {
          e.stopPropagation();
          options.onReset(definition.id);
        },
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

    const row = createRow(
      i18n.t('controller.stickRow', { label: definition.label }),
      formatFriendlyStickLabel(binding),
      binding.kind === 'none',
      isExpanded,
      i18n.t('controller.learnStick', { label: definition.label }),
      (e) => {
        e.stopPropagation();
        expandedRowKey = rowKey;
        options.onLearn(definition.id, 'axis');
      },
      i18n.t('controller.resetStick', { label: definition.label }),
      (e) => {
        e.stopPropagation();
        options.onReset(definition.id);
      },
    );
    row.addEventListener('click', () => {
      expandedRowKey = expandedRowKey === rowKey ? null : rowKey;
      render();
    });
    options.container.appendChild(row);

    if (isExpanded) {
      const summary = binding.kind === 'none' ? i18n.t('controller.disabled') : `Axis ${binding.axisIndex}`;
      const hint = isLearning
        ? i18n.t('controller.learnStickPrompt')
        : i18n.t('controller.currently', { binding: summary });
      const panel = createEditPanel(hint, isLearning, {
        onLearn: (e) => {
          e.stopPropagation();
          options.onLearn(definition.id, 'axis');
        },
        onClear: (e) => {
          e.stopPropagation();
          options.onClear(definition.id);
        },
        onReset: (e) => {
          e.stopPropagation();
          options.onReset(definition.id);
        },
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

    const dpadFallback: ControllerDpadFallback =
      binding.kind === 'none' ? 'none' : binding.dpadFallback;
    const badgeText = DPAD_FALLBACK_LABELS[dpadFallback];
    const row = createRow(
      i18n.t('controller.dpadRow', { label: definition.label }),
      badgeText,
      dpadFallback === 'none',
      isExpanded,
      i18n.t('controller.learnDpad', { label: definition.label }),
      (e) => {
        e.stopPropagation();
        expandedRowKey = rowKey;
        options.onDpadLearn(definition.id);
      },
      i18n.t('controller.resetDpad', { label: definition.label }),
      (e) => {
        e.stopPropagation();
        options.onDpadReset(definition.id);
      },
    );
    row.addEventListener('click', () => {
      expandedRowKey = expandedRowKey === rowKey ? null : rowKey;
      render();
    });
    options.container.appendChild(row);

    if (isExpanded) {
      const hint = isLearning
        ? i18n.t('controller.learnDpadPrompt')
        : i18n.t('controller.currently', { binding: DPAD_FALLBACK_LABELS[dpadFallback] });
      const panel = createEditPanel(hint, isLearning, {
        onLearn: (e) => {
          e.stopPropagation();
          options.onDpadLearn(definition.id);
        },
        onClear: (e) => {
          e.stopPropagation();
          options.onDpadClear(definition.id);
        },
        onReset: (e) => {
          e.stopPropagation();
          options.onDpadReset(definition.id);
        },
      });
      options.container.appendChild(panel);
    }
  }

  function createRow(
    labelText: string,
    badgeText: string,
    isDisabled: boolean,
    isExpanded: boolean,
    editLabel: string,
    onEdit: (e: Event) => void,
    resetLabel: string,
    onReset: (e: Event) => void,
  ): HTMLDivElement {
    const row = document.createElement('div');
    row.className = 'controller-config-row';
    if (isExpanded) row.classList.add('expanded');

    const label = document.createElement('div');
    label.className = 'controller-config-label';
    label.textContent = labelText;

    const right = document.createElement('div');
    right.className = 'controller-config-right';

    const badge = document.createElement('button');
    badge.type = 'button';
    badge.className = 'controller-config-badge';
    if (isDisabled) badge.classList.add('disabled');
    badge.setAttribute('aria-label', editLabel);
    badge.title = editLabel;
    badge.textContent = badgeText;
    badge.addEventListener('click', onEdit);

    const resetIcon = document.createElement('button');
    resetIcon.type = 'button';
    resetIcon.className = 'controller-config-reset-icon';
    resetIcon.setAttribute('aria-label', resetLabel);
    resetIcon.title = resetLabel;
    resetIcon.textContent = '\u21ba';
    resetIcon.addEventListener('click', onReset);

    const editIcon = document.createElement('button');
    editIcon.type = 'button';
    editIcon.className = 'controller-config-edit-icon';
    editIcon.setAttribute('aria-label', editLabel);
    editIcon.title = editLabel;
    editIcon.textContent = '\u270E';
    editIcon.addEventListener('click', onEdit);

    right.appendChild(badge);
    right.appendChild(resetIcon);
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
    learnButton.textContent = isLearning ? i18n.t('controller.learnListening') : i18n.t('controller.learnButton');
    learnButton.addEventListener('click', callbacks.onLearn);

    const clearButton = document.createElement('button');
    clearButton.type = 'button';
    clearButton.className = 'btn-secondary';
    clearButton.textContent = i18n.t('controller.clear');
    clearButton.addEventListener('click', callbacks.onClear);

    const resetButton = document.createElement('button');
    resetButton.type = 'button';
    resetButton.className = 'btn-secondary';
    resetButton.textContent = i18n.t('controller.reset');
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
