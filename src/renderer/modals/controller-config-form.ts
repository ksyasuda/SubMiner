import {
  findControllerScreenBindingIssues,
  getControllerFaceButtonName,
  formatControllerReviewBinding,
  type ControllerReviewInputContext,
  type ControllerScreenActionDefinition,
  type ControllerScreenBindings,
  type ControllerScreenDefinition,
} from '../../shared/controller-review';
import type { ControllerScreenId } from '../../shared/controller-screens';
import type { ResolvedControllerReviewBinding } from '../../types';

/** A dialog layout shown as a collapsible group in Alt+C. */
export type ControllerScreenGroup = {
  screen: ControllerScreenDefinition;
  bindings: ControllerScreenBindings;
};

/** The dialog action being learned, and the hold button pressed so far, if any. */
export type ControllerScreenLearning = {
  screenId: ControllerScreenId;
  actionId: string;
  holdLabel: string | null;
};
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
    id: 'previousEntry',
    label: 'Previous Entry',
    group: 'Lookup',
    bindingType: 'discrete',
    defaultBinding: { kind: 'none' },
  },
  {
    id: 'nextEntry',
    label: 'Next Entry',
    group: 'Lookup',
    bindingType: 'discrete',
    defaultBinding: { kind: 'none' },
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
    id: 'holdSecondarySub',
    label: 'Show Secondary Subtitle (Hold)',
    group: 'Playback',
    bindingType: 'discrete',
    defaultBinding: { kind: 'none' },
  },
  {
    id: 'quitMpv',
    label: 'Quit MPV',
    group: 'Playback',
    bindingType: 'discrete',
    defaultBinding: { kind: 'button', buttonIndex: 6 },
  },
  {
    id: 'previousSubtitle',
    label: 'Previous Subtitle',
    group: 'Subtitles',
    bindingType: 'discrete',
    defaultBinding: { kind: 'none' },
  },
  {
    id: 'nextSubtitle',
    label: 'Next Subtitle',
    group: 'Subtitles',
    bindingType: 'discrete',
    defaultBinding: { kind: 'none' },
  },
  {
    id: 'replaySubtitle',
    label: 'Replay Subtitle',
    group: 'Subtitles',
    bindingType: 'discrete',
    defaultBinding: { kind: 'none' },
  },
  {
    id: 'playNextSubtitle',
    label: 'Play Next Subtitle',
    group: 'Subtitles',
    bindingType: 'discrete',
    defaultBinding: { kind: 'none' },
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

// Names for pads that do not report the W3C "standard" mapping. Raw layouts
// vary by driver, so these are a best guess.
const RAW_LAYOUT_BUTTON_NAMES: Record<number, string> = {
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

const RAW_LAYOUT_AXIS_NAMES: Record<number, string> = {
  0: 'Left Stick X',
  1: 'Left Stick Y',
  2: 'Left Trigger',
  3: 'Right Stick X',
  4: 'Right Stick Y',
  5: 'Right Trigger',
};

// Names for pads that report `mapping === 'standard'` (most Xbox, PlayStation,
// Switch and 8BitDo pads). Index meanings are fixed by the W3C Gamepad spec:
// https://w3c.github.io/gamepad/#remapping
const W3C_STANDARD_BUTTON_NAMES: Record<number, string> = {
  0: 'A / Cross',
  1: 'B / Circle',
  2: 'X / Square',
  3: 'Y / Triangle',
  4: 'LB / L1',
  5: 'RB / R1',
  6: 'LT / L2',
  7: 'RT / R2',
  8: 'Back / Select',
  9: 'Start / Options',
  10: 'L3 / LS',
  11: 'R3 / RS',
  12: 'D-pad Up',
  13: 'D-pad Down',
  14: 'D-pad Left',
  15: 'D-pad Right',
  16: 'Guide / Home',
};

const W3C_STANDARD_AXIS_NAMES: Record<number, string> = {
  0: 'Left Stick X',
  1: 'Left Stick Y',
  2: 'Right Stick X',
  3: 'Right Stick Y',
};

type ControllerMapping = string | null | undefined;
type FaceButtonLabels = ResolvedControllerConfig['faceButtonLabels'] | undefined;

const DPAD_FALLBACK_LABELS: Record<ControllerDpadFallback, string> = {
  none: 'None',
  horizontal: 'D-pad \u2194',
  vertical: 'D-pad \u2195',
};

export function getFriendlyButtonName(
  buttonIndex: number,
  mapping?: ControllerMapping,
  faceButtonLabels?: FaceButtonLabels,
): string {
  // The default Xbox style keeps the shared "A / Cross" names; the others print their own.
  // Face positions 0-3 are the same on raw and standard layouts.
  const styledName =
    faceButtonLabels && faceButtonLabels !== 'xbox'
      ? getControllerFaceButtonName(buttonIndex, faceButtonLabels)
      : null;
  if (styledName) return styledName;
  const names = mapping === 'standard' ? W3C_STANDARD_BUTTON_NAMES : RAW_LAYOUT_BUTTON_NAMES;
  return names[buttonIndex] ?? `Button ${buttonIndex}`;
}

export function getFriendlyAxisName(axisIndex: number, mapping?: ControllerMapping): string {
  const names = mapping === 'standard' ? W3C_STANDARD_AXIS_NAMES : RAW_LAYOUT_AXIS_NAMES;
  return names[axisIndex] ?? `Axis ${axisIndex}`;
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

function formatFriendlyStickLabel(
  binding: ResolvedControllerAxisBinding,
  mapping: ControllerMapping,
): string {
  if (binding.kind === 'none') return 'None';
  return getFriendlyAxisName(binding.axisIndex, mapping);
}

function formatFriendlyBindingLabel(
  binding: ResolvedControllerDiscreteBinding | ResolvedControllerAxisBinding,
  mapping: ControllerMapping,
  faceButtonLabels: FaceButtonLabels,
): string {
  if (binding.kind === 'none') return 'None';
  if ('direction' in binding) {
    const name = getFriendlyAxisName(binding.axisIndex, mapping);
    return `${name} ${binding.direction === 'positive' ? '+' : '\u2212'}`;
  }
  if ('buttonIndex' in binding) {
    return getFriendlyButtonName(binding.buttonIndex, mapping, faceButtonLabels);
  }
  return getFriendlyAxisName(binding.axisIndex, mapping);
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
  /** Gamepad API `mapping` of the controller being configured; picks the button names. */
  getControllerMapping?: () => ControllerMapping;
  /** `controller.faceButtonLabels`; names the face buttons as printed on the pad. */
  getFaceButtonLabels?: () => FaceButtonLabels;
  /** Dialog layouts (review, field grouping) of the controller being configured. */
  getScreenGroups?: () => ControllerScreenGroup[];
  getScreenInputContext?: () => ControllerReviewInputContext;
  getScreenLearning?: () => ControllerScreenLearning | null;
  onScreenLearn?: (screenId: ControllerScreenId, actionId: string) => void;
  onScreenClear?: (screenId: ControllerScreenId, actionId: string) => void;
  onScreenReset?: (screenId: ControllerScreenId, actionId: string) => void;
  onScreenResetAll?: (screenId: ControllerScreenId) => void;
}) {
  let expandedRowKey: ExpandedRowKey | null = null;
  /** Dialog groups start collapsed; this holds the opened ones. */
  const expandedScreens = new Set<ControllerScreenId>();

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

    const context = options.getScreenInputContext?.();
    if (context) {
      for (const group of options.getScreenGroups?.() ?? []) {
        renderScreenGroup(group, context);
      }
    }
  }

  function renderScreenGroup(
    { screen, bindings }: ControllerScreenGroup,
    context: ControllerReviewInputContext,
  ): void {
    const learning = options.getScreenLearning?.() ?? null;
    const learningHere = learning?.screenId === screen.id ? learning : null;
    if (learningHere) {
      expandedScreens.add(screen.id);
      expandedRowKey = `${screen.id}:${learningHere.actionId}`;
    }
    const expanded = expandedScreens.has(screen.id);

    const header = document.createElement('div');
    header.className = 'controller-config-group controller-config-group-toggle';
    header.setAttribute('role', 'button');
    header.setAttribute('aria-expanded', String(expanded));
    header.textContent = `${expanded ? '\u25be' : '\u25b8'} ${screen.title}`;
    header.addEventListener('click', () => {
      if (expanded) expandedScreens.delete(screen.id);
      else expandedScreens.add(screen.id);
      render();
    });
    if (expanded) {
      const resetAll = document.createElement('button');
      resetAll.type = 'button';
      resetAll.className = 'btn-secondary controller-config-group-action';
      resetAll.textContent = 'Reset all';
      resetAll.addEventListener('click', (e) => {
        e.stopPropagation();
        options.onScreenResetAll?.(screen.id);
      });
      header.appendChild(resetAll);
    }
    options.container.appendChild(header);
    if (!expanded) return;

    const issues = findControllerScreenBindingIssues(screen, bindings, context);
    for (const step of screen.steps) {
      const subgroup = document.createElement('div');
      subgroup.className = 'controller-config-subgroup';
      subgroup.textContent = step.label;
      options.container.appendChild(subgroup);
      for (const action of screen.actions) {
        if (action.step !== step.id) continue;
        renderScreenRow(
          screen,
          action,
          bindings[action.id] ?? { kind: 'none' },
          context,
          issues[action.id],
          learningHere,
        );
      }
    }
  }

  function renderScreenRow(
    screen: ControllerScreenDefinition,
    action: ControllerScreenActionDefinition,
    binding: ResolvedControllerReviewBinding,
    context: ControllerReviewInputContext,
    issue: string | undefined,
    learning: ControllerScreenLearning | null,
  ): void {
    const rowKey = `${screen.id}:${action.id}`;
    const isExpanded = expandedRowKey === rowKey;
    const isLearning = learning?.actionId === action.id;
    const holdLabel = isLearning ? learning.holdLabel : null;
    const current = formatControllerReviewBinding(binding, context, { dpadPrefix: true });

    const learn = (e: Event) => {
      e.stopPropagation();
      expandedRowKey = rowKey;
      options.onScreenLearn?.(screen.id, action.id);
    };
    const reset = (e: Event) => {
      e.stopPropagation();
      options.onScreenReset?.(screen.id, action.id);
    };
    const row = createRow(
      action.label,
      holdLabel ? `${holdLabel} + \u2026` : current,
      binding.kind === 'none',
      isExpanded,
      `Learn ${action.label}`,
      learn,
      `Reset ${action.label}`,
      reset,
      issue,
    );
    row.addEventListener('click', () => {
      expandedRowKey = expandedRowKey === rowKey ? null : rowKey;
      render();
    });
    options.container.appendChild(row);

    if (isExpanded) {
      const hint = !isLearning
        ? `Currently: ${current}`
        : holdLabel
          ? `Holding ${holdLabel}: press another button for a combo, or let go to use ${holdLabel} alone\u2026`
          : 'Press a button, or hold one and press another for a combo\u2026';
      const panel = createEditPanel(hint, isLearning, {
        onLearn: learn,
        onClear: (e) => {
          e.stopPropagation();
          options.onScreenClear?.(screen.id, action.id);
        },
        onReset: reset,
      });
      options.container.appendChild(panel);
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
      formatFriendlyBindingLabel(
        binding,
        options.getControllerMapping?.(),
        options.getFaceButtonLabels?.(),
      ),
      binding.kind === 'none',
      isExpanded,
      `Learn ${definition.label}`,
      (e) => {
        e.stopPropagation();
        expandedRowKey = rowKey;
        options.onLearn(definition.id, definition.bindingType);
      },
      `Reset ${definition.label}`,
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
        ? 'Press a button, trigger, or move a stick\u2026'
        : `Currently: ${formatControllerBindingSummary(binding)}`;
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
      `${definition.label} (Stick)`,
      formatFriendlyStickLabel(binding, options.getControllerMapping?.()),
      binding.kind === 'none',
      isExpanded,
      `Learn ${definition.label} stick`,
      (e) => {
        e.stopPropagation();
        expandedRowKey = rowKey;
        options.onLearn(definition.id, 'axis');
      },
      `Reset ${definition.label} stick`,
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
      const summary = binding.kind === 'none' ? 'Disabled' : `Axis ${binding.axisIndex}`;
      const hint = isLearning ? 'Move a stick or trigger\u2026' : `Currently: ${summary}`;
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
      `${definition.label} (D-pad)`,
      badgeText,
      dpadFallback === 'none',
      isExpanded,
      `Learn ${definition.label} D-pad`,
      (e) => {
        e.stopPropagation();
        expandedRowKey = rowKey;
        options.onDpadLearn(definition.id);
      },
      `Reset ${definition.label} D-pad`,
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
        ? 'Press a D-pad direction\u2026'
        : `Currently: ${DPAD_FALLBACK_LABELS[dpadFallback]}`;
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
    warningText?: string,
  ): HTMLDivElement {
    const row = document.createElement('div');
    row.className = 'controller-config-row';
    if (isExpanded) row.classList.add('expanded');

    const label = document.createElement('div');
    label.className = 'controller-config-label';
    label.textContent = labelText;
    if (warningText) {
      row.classList.add('has-warning');
      const warning = document.createElement('div');
      warning.className = 'controller-config-warning';
      warning.textContent = `\u26a0 ${warningText}`;
      label.appendChild(warning);
    }

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
