import type { Keybinding } from '../../types';
import type { ShortcutsConfig } from '../../types';
import { SPECIAL_COMMANDS } from '../../config/definitions/shared';
import type { ModalStateReader, RendererContext } from '../context';

type SessionHelpBindingInfo = {
  bindingKey: 'KeyH' | 'KeyK';
  fallbackUsed: boolean;
  fallbackUnavailable: boolean;
};

type SessionHelpItem = {
  shortcut: string;
  action: string;
  color?: string;
};

type SessionHelpSection = {
  title: string;
  rows: SessionHelpItem[];
};
type RuntimeShortcutConfig = Omit<Required<ShortcutsConfig>, 'multiCopyTimeoutMs'>;

const HEX_COLOR_RE = /^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/;

// Fallbacks mirror the session overlay's default subtitle/word color scheme.
const FALLBACK_COLORS = {
  knownWordColor: '#a6da95',
  nPlusOneColor: '#c6a0f6',
  nameMatchColor: '#f5bde6',
  jlptN1Color: '#ed8796',
  jlptN2Color: '#f5a97f',
  jlptN3Color: '#f9e2af',
  jlptN4Color: '#a6e3a1',
  jlptN5Color: '#8aadf4',
};

const KEY_NAME_MAP: Record<string, string> = {
  Space: 'Space',
  ArrowUp: '↑',
  ArrowDown: '↓',
  ArrowLeft: '←',
  ArrowRight: '→',
  Escape: 'Esc',
  Tab: 'Tab',
  Enter: 'Enter',
  BracketLeft: '[',
  BracketRight: ']',
  CommandOrControl: 'Cmd/Ctrl',
  Ctrl: 'Ctrl',
  Control: 'Ctrl',
  Command: 'Cmd',
  Cmd: 'Cmd',
  Shift: 'Shift',
  Alt: 'Alt',
  Super: 'Meta',
  Meta: 'Meta',
  Backspace: 'Backspace',
};

function normalizeColor(value: unknown, fallback: string): string {
  if (typeof value !== 'string') return fallback;
  const next = value.trim();
  return HEX_COLOR_RE.test(next) ? next : fallback;
}

function normalizeKeyToken(token: string): string {
  if (KEY_NAME_MAP[token]) return KEY_NAME_MAP[token];
  if (token.startsWith('Key')) return token.slice(3);
  if (token.startsWith('Digit')) return token.slice(5);
  if (token.startsWith('Numpad')) return token.slice(6);
  return token;
}

function formatKeybinding(rawBinding: string): string {
  const parts = rawBinding.split('+');
  const key = parts.pop();
  if (!key) return rawBinding;
  const normalized = [...parts, normalizeKeyToken(key)];
  return normalized.join(' + ');
}

const OVERLAY_SHORTCUTS: Array<{
  key: keyof RuntimeShortcutConfig;
  label: string;
}> = [
  { key: 'copySubtitle', label: 'Copy subtitle' },
  { key: 'copySubtitleMultiple', label: 'Copy subtitle (multi)' },
  {
    key: 'updateLastCardFromClipboard',
    label: 'Update last card from clipboard',
  },
  { key: 'triggerFieldGrouping', label: 'Trigger field grouping' },
  { key: 'triggerSubsync', label: 'Open subtitle sync controls' },
  { key: 'mineSentence', label: 'Mine sentence' },
  { key: 'mineSentenceMultiple', label: 'Mine sentence (multi)' },
  { key: 'toggleSecondarySub', label: 'Toggle secondary subtitle mode' },
  { key: 'markAudioCard', label: 'Mark audio card' },
  { key: 'openCharacterDictionary', label: 'Open character dictionary anime selector' },
  { key: 'openRuntimeOptions', label: 'Open runtime options' },
  { key: 'openJimaku', label: 'Open jimaku' },
  { key: 'openSessionHelp', label: 'Open session help' },
  { key: 'openControllerSelect', label: 'Open controller select' },
  { key: 'openControllerDebug', label: 'Open controller debug' },
  { key: 'toggleSubtitleSidebar', label: 'Toggle subtitle sidebar' },
  { key: 'toggleVisibleOverlayGlobal', label: 'Show/hide visible overlay' },
];

function buildOverlayShortcutSections(shortcuts: RuntimeShortcutConfig): SessionHelpSection[] {
  const rows: SessionHelpItem[] = [];

  for (const shortcut of OVERLAY_SHORTCUTS) {
    const keybind = shortcuts[shortcut.key];

    rows.push({
      shortcut:
        typeof keybind === 'string' && keybind.trim().length > 0
          ? formatKeybinding(keybind)
          : 'Unbound',
      action: shortcut.label,
    });
  }

  if (rows.length === 0) return [];
  return [{ title: 'Overlay shortcuts', rows }];
}

function describeCommand(command: (string | number)[]): string {
  const first = command[0];
  if (typeof first !== 'string') return 'Unknown action';

  if (first === 'cycle' && command[1] === 'pause') return 'Toggle playback';
  if (first === 'seek' && typeof command[1] === 'number') {
    return `Seek ${command[1] > 0 ? '+' : ''}${command[1]} second(s)`;
  }
  if (first === 'sub-seek' && typeof command[1] === 'number') {
    if (command[1] > 0) return 'Jump to next subtitle';
    if (command[1] < 0) return 'Jump to previous subtitle';
    return 'Reload current subtitle timing';
  }
  if (first === SPECIAL_COMMANDS.SUBSYNC_TRIGGER) return 'Open subtitle sync controls';
  if (first === SPECIAL_COMMANDS.RUNTIME_OPTIONS_OPEN) return 'Open runtime options';
  if (first === SPECIAL_COMMANDS.JIMAKU_OPEN) return 'Open jimaku';
  if (first === SPECIAL_COMMANDS.PLAYLIST_BROWSER_OPEN) return 'Open playlist browser';
  if (first === SPECIAL_COMMANDS.REPLAY_SUBTITLE) return 'Replay current subtitle';
  if (first === SPECIAL_COMMANDS.PLAY_NEXT_SUBTITLE) return 'Play next subtitle';
  if (first === SPECIAL_COMMANDS.SHIFT_SUB_DELAY_TO_NEXT_SUBTITLE_START) {
    return 'Shift subtitle delay to next cue';
  }
  if (first === SPECIAL_COMMANDS.SHIFT_SUB_DELAY_TO_PREVIOUS_SUBTITLE_START) {
    return 'Shift subtitle delay to previous cue';
  }
  if (first.startsWith(SPECIAL_COMMANDS.RUNTIME_OPTION_CYCLE_PREFIX)) {
    const [, rawId, rawDirection] = first.split(':');
    return `Cycle runtime option ${rawId || 'option'} ${rawDirection === 'prev' ? 'previous' : 'next'}`;
  }

  return `MPV command: ${command.map((entry) => String(entry)).join(' ')}`;
}

export {
  describeCommand as describeSessionHelpCommand,
  formatKeybinding as formatSessionHelpKeybinding,
};

function sectionForCommand(command: (string | number)[]): string {
  const first = command[0];
  if (typeof first !== 'string') return 'Other shortcuts';

  if (
    first === 'cycle' ||
    first === 'seek' ||
    first === 'sub-seek' ||
    first === SPECIAL_COMMANDS.REPLAY_SUBTITLE ||
    first === SPECIAL_COMMANDS.PLAY_NEXT_SUBTITLE
  ) {
    return 'Playback and navigation';
  }

  if (first === 'show-text' || first === 'show-progress' || first.startsWith('osd')) {
    return 'Visual feedback';
  }

  if (first === SPECIAL_COMMANDS.SUBSYNC_TRIGGER) {
    return 'Subtitle sync';
  }

  if (
    first === SPECIAL_COMMANDS.RUNTIME_OPTIONS_OPEN ||
    first === SPECIAL_COMMANDS.JIMAKU_OPEN ||
    first === SPECIAL_COMMANDS.PLAYLIST_BROWSER_OPEN ||
    first.startsWith(SPECIAL_COMMANDS.RUNTIME_OPTION_CYCLE_PREFIX)
  ) {
    return 'Runtime settings';
  }

  if (first === 'quit') return 'System actions';
  return 'Other shortcuts';
}

function buildBindingSections(keybindings: Keybinding[]): SessionHelpSection[] {
  const grouped = new Map<string, SessionHelpItem[]>();

  for (const binding of keybindings) {
    const section = sectionForCommand(binding.command ?? []);
    const row: SessionHelpItem = {
      shortcut: formatKeybinding(binding.key),
      action: describeCommand(binding.command ?? []),
    };
    grouped.set(section, [...(grouped.get(section) ?? []), row]);
  }

  const sectionOrder = [
    'Playback and navigation',
    'Visual feedback',
    'Subtitle sync',
    'Runtime settings',
    'System actions',
    'Other shortcuts',
  ];
  const sectionEntries = Array.from(grouped.entries()).sort((a, b) => {
    const aIdx = sectionOrder.indexOf(a[0]);
    const bIdx = sectionOrder.indexOf(b[0]);
    if (aIdx === -1 && bIdx === -1) return a[0].localeCompare(b[0]);
    if (aIdx === -1) return 1;
    if (bIdx === -1) return -1;
    return aIdx - bIdx;
  });

  return sectionEntries.map(([title, rows]) => ({ title, rows }));
}

function buildColorSection(style: {
  knownWordColor?: unknown;
  nPlusOneColor?: unknown;
  nameMatchColor?: unknown;
  jlptColors?: {
    N1?: unknown;
    N2?: unknown;
    N3?: unknown;
    N4?: unknown;
    N5?: unknown;
  };
}): SessionHelpSection {
  return {
    title: 'Color legend',
    rows: [
      {
        shortcut: 'Known words',
        action: normalizeColor(style.knownWordColor, FALLBACK_COLORS.knownWordColor),
        color: normalizeColor(style.knownWordColor, FALLBACK_COLORS.knownWordColor),
      },
      {
        shortcut: 'N+1 words',
        action: normalizeColor(style.nPlusOneColor, FALLBACK_COLORS.nPlusOneColor),
        color: normalizeColor(style.nPlusOneColor, FALLBACK_COLORS.nPlusOneColor),
      },
      {
        shortcut: 'Character names',
        action: normalizeColor(style.nameMatchColor, FALLBACK_COLORS.nameMatchColor),
        color: normalizeColor(style.nameMatchColor, FALLBACK_COLORS.nameMatchColor),
      },
      {
        shortcut: 'JLPT N1',
        action: normalizeColor(style.jlptColors?.N1, FALLBACK_COLORS.jlptN1Color),
        color: normalizeColor(style.jlptColors?.N1, FALLBACK_COLORS.jlptN1Color),
      },
      {
        shortcut: 'JLPT N2',
        action: normalizeColor(style.jlptColors?.N2, FALLBACK_COLORS.jlptN2Color),
        color: normalizeColor(style.jlptColors?.N2, FALLBACK_COLORS.jlptN2Color),
      },
      {
        shortcut: 'JLPT N3',
        action: normalizeColor(style.jlptColors?.N3, FALLBACK_COLORS.jlptN3Color),
        color: normalizeColor(style.jlptColors?.N3, FALLBACK_COLORS.jlptN3Color),
      },
      {
        shortcut: 'JLPT N4',
        action: normalizeColor(style.jlptColors?.N4, FALLBACK_COLORS.jlptN4Color),
        color: normalizeColor(style.jlptColors?.N4, FALLBACK_COLORS.jlptN4Color),
      },
      {
        shortcut: 'JLPT N5',
        action: normalizeColor(style.jlptColors?.N5, FALLBACK_COLORS.jlptN5Color),
        color: normalizeColor(style.jlptColors?.N5, FALLBACK_COLORS.jlptN5Color),
      },
    ],
  };
}

function filterSections(sections: SessionHelpSection[], query: string): SessionHelpSection[] {
  const normalize = (value: string): string =>
    value
      .toLowerCase()
      .replace(/commandorcontrol/gu, 'ctrl')
      .replace(/cmd\/ctrl/gu, 'ctrl')
      .replace(/[\s+\-_/]/gu, '');
  const normalized = normalize(query);
  if (!normalized) return sections;

  return sections
    .map((section) => {
      if (normalize(section.title).includes(normalized)) {
        return section;
      }

      const rows = section.rows.filter(
        (row) =>
          normalize(row.shortcut).includes(normalized) ||
          normalize(row.action).includes(normalized),
      );
      if (rows.length === 0) return null;
      return { ...section, rows };
    })
    .filter((section): section is SessionHelpSection => section !== null)
    .filter((section) => section.rows.length > 0);
}

function formatBindingHint(info: SessionHelpBindingInfo): string {
  if (info.bindingKey === 'KeyK' && info.fallbackUsed) {
    return info.fallbackUnavailable ? 'Y-K (fallback and conflict noted)' : 'Y-K (fallback)';
  }
  return 'Y-H';
}

function createShortcutRow(row: SessionHelpItem, globalIndex: number): HTMLButtonElement {
  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'session-help-item';
  button.tabIndex = -1;
  button.dataset.sessionHelpIndex = String(globalIndex);

  const left = document.createElement('div');
  left.className = 'session-help-item-left';
  const shortcut = document.createElement('span');
  shortcut.className = 'session-help-key';
  shortcut.textContent = row.shortcut;
  left.appendChild(shortcut);

  const right = document.createElement('div');
  right.className = 'session-help-item-right';
  const action = document.createElement('span');
  action.className = 'session-help-action';
  action.textContent = row.action;
  right.appendChild(action);

  if (row.color) {
    const dot = document.createElement('span');
    dot.className = 'session-help-color-dot';
    dot.style.backgroundColor = row.color;
    right.insertBefore(dot, action);
  }

  button.appendChild(left);
  button.appendChild(right);
  return button;
}

const SECTION_ICON: Record<string, string> = {
  'MPV shortcuts': '⚙',
  'Playback and navigation': '▶',
  'Visual feedback': '◉',
  'Subtitle sync': '⟲',
  'Runtime settings': '⚙',
  'System actions': '◆',
  'Other shortcuts': '…',
  'Overlay shortcuts (configurable)': '✦',
  'Overlay shortcuts': '✦',
  'Color legend': '◈',
};

function createSectionNode(
  section: SessionHelpSection,
  sectionIndex: number,
  globalIndexMap: number[],
): HTMLElement {
  const sectionNode = document.createElement('section');
  sectionNode.className = 'session-help-section';

  const title = document.createElement('h3');
  title.className = 'session-help-section-title';
  const icon = SECTION_ICON[section.title] ?? '•';
  title.textContent = `${icon} ${section.title}`;
  sectionNode.appendChild(title);

  const list = document.createElement('div');
  list.className = 'session-help-item-list';

  section.rows.forEach((row, rowIndex) => {
    const globalIndex = (globalIndexMap[sectionIndex] ?? 0) + rowIndex;
    const button = createShortcutRow(row, globalIndex);
    list.appendChild(button);
  });

  sectionNode.appendChild(list);
  return sectionNode;
}

export function createSessionHelpModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  let priorFocus: Element | null = null;
  let openBinding: SessionHelpBindingInfo = {
    bindingKey: 'KeyH',
    fallbackUsed: false,
    fallbackUnavailable: false,
  };
  let helpFilterValue = '';
  let helpSections: SessionHelpSection[] = [];
  let focusGuard: ((event: FocusEvent) => void) | null = null;
  let windowFocusGuard: (() => void) | null = null;
  let modalPointerFocusGuard: ((event: Event) => void) | null = null;
  let isRecoveringModalFocus = false;
  let lastFocusRecoveryAt = 0;

  function getItems(): HTMLButtonElement[] {
    return Array.from(
      ctx.dom.sessionHelpContent.querySelectorAll('.session-help-item'),
    ) as HTMLButtonElement[];
  }

  function setSelected(index: number): void {
    const items = getItems();
    if (items.length === 0) return;

    const wrappedIndex = index % items.length;
    const next = wrappedIndex < 0 ? wrappedIndex + items.length : wrappedIndex;
    ctx.state.sessionHelpSelectedIndex = next;

    items.forEach((item, idx) => {
      item.classList.toggle('active', idx === next);
      item.tabIndex = idx === next ? 0 : -1;
    });
    const activeItem = items[next];
    if (!activeItem) return;
    activeItem.focus({ preventScroll: true });
    activeItem.scrollIntoView({
      block: 'nearest',
      inline: 'nearest',
    });
  }

  function isSessionHelpModalFocusTarget(target: EventTarget | null): boolean {
    return target instanceof Element && ctx.dom.sessionHelpModal.contains(target);
  }

  function focusFallbackTarget(): boolean {
    if (!ctx.platform.isModalLayer) {
      void window.electronAPI.focusMainWindow();
    }
    const items = getItems();
    const firstItem = items.find((item) => item.offsetParent !== null);
    if (firstItem) {
      firstItem.focus({ preventScroll: true });
      return document.activeElement === firstItem;
    }

    if (ctx.dom.sessionHelpClose instanceof HTMLElement) {
      ctx.dom.sessionHelpClose.focus({ preventScroll: true });
      return document.activeElement === ctx.dom.sessionHelpClose;
    }

    window.focus();
    return false;
  }

  function enforceModalFocus(): void {
    if (!ctx.state.sessionHelpModalOpen) return;
    if (!isSessionHelpModalFocusTarget(document.activeElement)) {
      if (isRecoveringModalFocus) return;

      const now = Date.now();
      if (now - lastFocusRecoveryAt < 120) return;

      isRecoveringModalFocus = true;
      lastFocusRecoveryAt = now;
      focusFallbackTarget();

      window.setTimeout(() => {
        isRecoveringModalFocus = false;
      }, 120);
    }
  }

  function isFilterInputFocused(): boolean {
    return document.activeElement === ctx.dom.sessionHelpFilter;
  }

  function focusFilterInput(): void {
    ctx.dom.sessionHelpFilter.focus({ preventScroll: true });
    ctx.dom.sessionHelpFilter.select();
  }

  function applyFilterAndRender(): void {
    const sections = filterSections(helpSections, helpFilterValue);
    const indexOffsets: number[] = [];
    let running = 0;
    for (const section of sections) {
      indexOffsets.push(running);
      running += section.rows.length;
    }

    ctx.dom.sessionHelpContent.innerHTML = '';
    sections.forEach((section, sectionIndex) => {
      const sectionNode = createSectionNode(section, sectionIndex, indexOffsets);
      ctx.dom.sessionHelpContent.appendChild(sectionNode);
    });

    if (getItems().length === 0) {
      ctx.dom.sessionHelpContent.classList.add('session-help-content-no-results');
      ctx.dom.sessionHelpContent.textContent = helpFilterValue
        ? 'No matching shortcuts found.'
        : 'No active session shortcuts found.';
      ctx.state.sessionHelpSelectedIndex = 0;
      return;
    }

    ctx.dom.sessionHelpContent.classList.remove('session-help-content-no-results');

    if (isFilterInputFocused()) return;

    setSelected(0);
  }

  function requestOverlayFocus(): void {
    if (!ctx.platform.isModalLayer) {
      void window.electronAPI.focusMainWindow();
    }
  }

  function addPointerFocusListener(): void {
    if (modalPointerFocusGuard) return;

    modalPointerFocusGuard = () => {
      requestOverlayFocus();
      enforceModalFocus();
    };
    ctx.dom.sessionHelpModal.addEventListener('pointerdown', modalPointerFocusGuard);
    ctx.dom.sessionHelpModal.addEventListener('click', modalPointerFocusGuard);
  }

  function removePointerFocusListener(): void {
    if (!modalPointerFocusGuard) return;
    ctx.dom.sessionHelpModal.removeEventListener('pointerdown', modalPointerFocusGuard);
    ctx.dom.sessionHelpModal.removeEventListener('click', modalPointerFocusGuard);
    modalPointerFocusGuard = null;
  }

  function startFocusRecoveryGuards(): void {
    if (windowFocusGuard) return;

    windowFocusGuard = () => {
      requestOverlayFocus();
      enforceModalFocus();
    };
    window.addEventListener('blur', windowFocusGuard);
    window.addEventListener('focus', windowFocusGuard);
  }

  function stopFocusRecoveryGuards(): void {
    if (!windowFocusGuard) return;
    window.removeEventListener('blur', windowFocusGuard);
    window.removeEventListener('focus', windowFocusGuard);
    windowFocusGuard = null;
  }

  function showRenderError(message: string): void {
    helpSections = [];
    helpFilterValue = '';
    ctx.dom.sessionHelpFilter.value = '';
    ctx.dom.sessionHelpContent.classList.add('session-help-content-no-results');
    ctx.dom.sessionHelpContent.textContent = message;
    ctx.state.sessionHelpSelectedIndex = 0;
  }

  async function render(): Promise<boolean> {
    try {
      const [keybindings, styleConfig, shortcuts] = await Promise.all([
        window.electronAPI.getKeybindings(),
        window.electronAPI.getSubtitleStyle(),
        window.electronAPI.getConfiguredShortcuts(),
      ]);

      const bindingSections = buildBindingSections(keybindings);
      if (bindingSections.length > 0) {
        const playback = bindingSections.find(
          (section) => section.title === 'Playback and navigation',
        );
        if (playback) {
          playback.title = 'MPV shortcuts';
        }
      }

      const shortcutSections = buildOverlayShortcutSections(shortcuts);
      if (shortcutSections.length > 0) {
        shortcutSections[0]!.title = 'Overlay shortcuts (configurable)';
      }
      const colorSection = buildColorSection(styleConfig ?? {});
      helpSections = [...bindingSections, ...shortcutSections, colorSection];
      applyFilterAndRender();
      return true;
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unable to load session help data.';
      showRenderError(`Session help failed to load: ${message}`);
      return false;
    }
  }

  function openSessionHelpModal(opening: SessionHelpBindingInfo): void {
    openBinding = opening;
    priorFocus = document.activeElement;

    ctx.state.sessionHelpModalOpen = true;
    helpSections = [];
    helpFilterValue = '';
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.sessionHelpModal.classList.remove('hidden');
    ctx.dom.sessionHelpModal.setAttribute('aria-hidden', 'false');
    ctx.dom.sessionHelpModal.setAttribute('tabindex', '-1');
    ctx.dom.sessionHelpFilter.value = '';
    ctx.state.sessionHelpSelectedIndex = 0;
    ctx.dom.sessionHelpContent.innerHTML = '';
    ctx.dom.sessionHelpContent.classList.remove('session-help-content-no-results');
    if (ctx.platform.shouldToggleMouseIgnore) {
      window.electronAPI.setIgnoreMouseEvents(false);
    }

    ctx.dom.sessionHelpShortcut.textContent = `Session help opened with ${formatBindingHint(openBinding)}`;
    if (openBinding.fallbackUnavailable) {
      ctx.dom.sessionHelpWarning.textContent =
        'Both Y-H and Y-K are bound; Y-K remains the fallback for this session.';
    } else if (openBinding.fallbackUsed) {
      ctx.dom.sessionHelpWarning.textContent = 'Y-H is already bound; using Y-K as fallback.';
    } else {
      ctx.dom.sessionHelpWarning.textContent = '';
    }
    ctx.dom.sessionHelpStatus.textContent = 'Loading session help data...';

    if (focusGuard === null) {
      focusGuard = (event: FocusEvent) => {
        if (!ctx.state.sessionHelpModalOpen) return;
        if (!isSessionHelpModalFocusTarget(event.target)) {
          event.preventDefault();
          enforceModalFocus();
        }
      };
      document.addEventListener('focusin', focusGuard);
    }

    addPointerFocusListener();
    startFocusRecoveryGuards();
    requestOverlayFocus();
    window.focus();
    enforceModalFocus();

    void render().then((dataLoaded) => {
      if (!ctx.state.sessionHelpModalOpen) return;
      if (dataLoaded) {
        ctx.dom.sessionHelpStatus.textContent =
          'Use Arrow keys, J/K/H/L, mouse, click, or / then type to filter. Esc closes.';
      } else {
        ctx.dom.sessionHelpStatus.textContent =
          'Session help data is unavailable right now. Press Esc to close.';
        ctx.dom.sessionHelpWarning.textContent =
          'Unable to load latest shortcut settings from the runtime.';
      }
    });
  }

  function closeSessionHelpModal(): void {
    if (!ctx.state.sessionHelpModalOpen) return;

    ctx.state.sessionHelpModalOpen = false;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.sessionHelpModal.classList.add('hidden');
    ctx.dom.sessionHelpModal.setAttribute('aria-hidden', 'true');
    window.electronAPI.notifyOverlayModalClosed('session-help');
    if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
      ctx.dom.overlay.classList.remove('interactive');
    }

    if (focusGuard) {
      document.removeEventListener('focusin', focusGuard);
      focusGuard = null;
    }
    removePointerFocusListener();
    stopFocusRecoveryGuards();

    if (priorFocus instanceof HTMLElement && priorFocus.isConnected) {
      priorFocus.focus({ preventScroll: true });
      return;
    }

    if (ctx.dom.overlay instanceof HTMLElement) {
      // Overlay remains `tabindex="-1"` to allow programmatic focus for fallback.
      ctx.dom.overlay.focus({ preventScroll: true });
    }
    if (ctx.platform.shouldToggleMouseIgnore) {
      if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
        window.electronAPI.setIgnoreMouseEvents(true, { forward: true });
      } else {
        window.electronAPI.setIgnoreMouseEvents(false);
      }
    }
    ctx.dom.sessionHelpFilter.value = '';
    helpFilterValue = '';
    window.focus();
  }

  function handleSessionHelpKeydown(e: KeyboardEvent): boolean {
    if (!ctx.state.sessionHelpModalOpen) return false;

    if (isFilterInputFocused()) {
      if (e.key === 'Escape') {
        e.preventDefault();
        if (!helpFilterValue) {
          closeSessionHelpModal();
          return true;
        }

        helpFilterValue = '';
        ctx.dom.sessionHelpFilter.value = '';
        applyFilterAndRender();
        focusFallbackTarget();
        return true;
      }
      return false;
    }

    if (e.key === 'Escape') {
      e.preventDefault();
      closeSessionHelpModal();
      return true;
    }

    const items = getItems();
    if (items.length === 0) return true;

    if (e.key === '/' && !e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey) {
      e.preventDefault();
      focusFilterInput();
      return true;
    }

    const key = e.key.toLowerCase();

    if (key === 'arrowdown' || key === 'j' || key === 'l') {
      e.preventDefault();
      setSelected(ctx.state.sessionHelpSelectedIndex + 1);
      return true;
    }

    if (key === 'arrowup' || key === 'k' || key === 'h') {
      e.preventDefault();
      setSelected(ctx.state.sessionHelpSelectedIndex - 1);
      return true;
    }

    return true;
  }

  function wireDomEvents(): void {
    ctx.dom.sessionHelpFilter.addEventListener('input', () => {
      helpFilterValue = ctx.dom.sessionHelpFilter.value;
      applyFilterAndRender();
    });

    ctx.dom.sessionHelpFilter.addEventListener('keydown', (event: KeyboardEvent) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        focusFallbackTarget();
      }
    });

    ctx.dom.sessionHelpContent.addEventListener('click', (event: MouseEvent) => {
      const target = event.target;
      if (!(target instanceof Element)) return;
      const row = target.closest('.session-help-item') as HTMLElement | null;
      if (!row) return;
      const index = Number.parseInt(row.dataset.sessionHelpIndex ?? '', 10);
      if (!Number.isFinite(index)) return;
      setSelected(index);
    });

    ctx.dom.sessionHelpClose.addEventListener('click', () => {
      closeSessionHelpModal();
    });
  }

  return {
    closeSessionHelpModal,
    handleSessionHelpKeydown,
    openSessionHelpModal,
    wireDomEvents,
  };
}
