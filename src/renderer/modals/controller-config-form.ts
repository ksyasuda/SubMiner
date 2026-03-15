import type {
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

export function createControllerConfigForm(options: {
  container: HTMLElement;
  getBindings: () => ResolvedControllerConfig['bindings'];
  getLearningActionId: () => ControllerBindingActionId | null;
  onLearn: (actionId: ControllerBindingActionId, bindingType: 'discrete' | 'axis') => void;
  onClear: (actionId: ControllerBindingActionId) => void;
  onReset: (actionId: ControllerBindingActionId) => void;
}) {
  function render(): void {
    options.container.innerHTML = '';
    let lastGroup = '';

    for (const definition of CONTROLLER_BINDING_DEFINITIONS) {
      if (definition.group !== lastGroup) {
        const header = document.createElement('div');
        header.className = 'controller-config-group';
        header.textContent = definition.group;
        options.container.appendChild(header);
        lastGroup = definition.group;
      }

      const row = document.createElement('div');
      row.className = 'controller-config-row';
      row.setAttribute('data-testid', `controller-row-${definition.id}`);
      row.classList.toggle('learning', options.getLearningActionId() === definition.id);

      const label = document.createElement('div');
      label.className = 'controller-config-label';
      label.textContent = definition.label;

      const value = document.createElement('div');
      value.className = 'controller-config-value';
      value.textContent = formatControllerBindingSummary(options.getBindings()[definition.id]);

      const actions = document.createElement('div');
      actions.className = 'controller-config-actions';

      const learnButton = document.createElement('button');
      learnButton.type = 'button';
      learnButton.className = 'kiku-confirm-button';
      learnButton.setAttribute('data-testid', 'learn-button');
      learnButton.textContent =
        options.getLearningActionId() === definition.id ? 'Learning...' : 'Learn';
      learnButton.addEventListener('click', () => {
        options.onLearn(definition.id, definition.bindingType);
      });

      const clearButton = document.createElement('button');
      clearButton.type = 'button';
      clearButton.className = 'kiku-cancel-button';
      clearButton.textContent = 'Clear';
      clearButton.addEventListener('click', () => {
        options.onClear(definition.id);
      });

      const resetButton = document.createElement('button');
      resetButton.type = 'button';
      resetButton.className = 'kiku-cancel-button';
      resetButton.textContent = 'Reset';
      resetButton.addEventListener('click', () => {
        options.onReset(definition.id);
      });

      actions.appendChild(learnButton);
      actions.appendChild(clearButton);
      actions.appendChild(resetButton);

      row.appendChild(label);
      row.appendChild(value);
      row.appendChild(actions);
      options.container.appendChild(row);
    }
  }

  return { render };
}
