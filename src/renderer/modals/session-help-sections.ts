import type {
  CompiledSessionBinding,
  SessionActionId,
  SessionKeyModifier,
  SessionKeySpec,
} from '../../types';
import { SPECIAL_COMMANDS } from '../../config/definitions/shared';
import { i18n } from '../../i18n/index.js';
import { buildColorSection, type SessionHelpSubtitleStyle } from './session-help-colors';

export type SessionHelpItem = {
  shortcut: string;
  action: string;
  color?: string;
};

export type SessionHelpSection = {
  title: string;
  rows: SessionHelpItem[];
};

export type SessionHelpTabId = 'essentials' | 'playback' | 'mining' | 'tools' | 'reference';

export type SessionHelpTab = {
  id: SessionHelpTabId;
  label: string;
};

export function getSessionHelpTabs(): SessionHelpTab[] {
  return [
    { id: 'essentials', label: i18n.t('sessionHelp.tab.essentials') },
    { id: 'playback', label: i18n.t('sessionHelp.tab.playback') },
    { id: 'mining', label: i18n.t('sessionHelp.tab.mining') },
    { id: 'tools', label: i18n.t('sessionHelp.tab.tools') },
    { id: 'reference', label: i18n.t('sessionHelp.tab.reference') },
  ];
}

const KEY_NAME_MAP: Record<string, string> = {
  Space: 'Space',
  ArrowUp: '↑',
  ArrowDown: '↓',
  ArrowLeft: '←',
  ArrowRight: '→',
  Escape: 'Esc',
  Tab: 'Tab',
  Enter: 'Enter',
  Slash: '/',
  Backslash: '\\',
  Backquote: '`',
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
  MBTN_LEFT: 'Mouse Left',
  MBTN_MID: 'Mouse Middle',
  MBTN_RIGHT: 'Mouse Right',
  MBTN_BACK: 'Mouse Back',
  MBTN_FORWARD: 'Mouse Forward',
};

function normalizeKeyToken(token: string): string {
  if (KEY_NAME_MAP[token]) return KEY_NAME_MAP[token];
  if (token.startsWith('Key')) return token.slice(3);
  if (token.startsWith('Digit')) return token.slice(5);
  if (token.startsWith('Numpad')) return token.slice(6);
  return token;
}

function formatKeybinding(rawBinding: string): string {
  const parts = rawBinding
    .split('+')
    .map((part) => part.trim())
    .filter(Boolean);
  const key = parts.pop();
  if (!key) return rawBinding;
  const normalized = [...parts.map(normalizeKeyToken), normalizeKeyToken(key)];
  return normalized.join(' + ');
}

function describeCommand(command: (string | number)[]): string {
  const first = command[0];
  if (typeof first !== 'string') return i18n.t('sessionHelp.unknownAction');

  if (first === 'cycle' && command[1] === 'pause') return i18n.t('sessionHelp.action.togglePlayback');
  if (first === 'seek' && typeof command[1] === 'number') {
    return i18n.t('sessionHelp.action.seekSubtitle', {
      sign: command[1] > 0 ? '+' : '',
      value: String(command[1]),
    });
  }
  if (first === 'sub-seek' && typeof command[1] === 'number') {
    if (command[1] > 0) return i18n.t('sessionHelp.action.jumpToNextSubtitle');
    if (command[1] < 0) return i18n.t('sessionHelp.action.jumpToPreviousSubtitle');
    return i18n.t('sessionHelp.action.reloadSubtitleTiming');
  }
  if (first === 'sub-step' && typeof command[1] === 'number') {
    if (command[1] > 0) return i18n.t('sessionHelp.action.shiftDelayNext');
    if (command[1] < 0) return i18n.t('sessionHelp.action.shiftDelayPrev');
    return i18n.t('sessionHelp.action.reloadSubtitleTiming');
  }
  if (first === SPECIAL_COMMANDS.SUBSYNC_TRIGGER) return i18n.t('sessionHelp.action.openSubsync');
  if (first === SPECIAL_COMMANDS.RUNTIME_OPTIONS_OPEN)
    return i18n.t('sessionHelp.action.openRuntimeOptions');
  if (first === SPECIAL_COMMANDS.JIMAKU_OPEN) return i18n.t('sessionHelp.action.openJimaku');
  if (first === SPECIAL_COMMANDS.PLAYLIST_BROWSER_OPEN)
    return i18n.t('sessionHelp.action.openPlaylistBrowser');
  if (first === SPECIAL_COMMANDS.REPLAY_SUBTITLE)
    return i18n.t('sessionHelp.action.replaySubtitle');
  if (first === SPECIAL_COMMANDS.PLAY_NEXT_SUBTITLE)
    return i18n.t('sessionHelp.action.playNextSubtitle');
  if (first.startsWith(SPECIAL_COMMANDS.RUNTIME_OPTION_CYCLE_PREFIX)) {
    const [, rawId, rawDirection] = first.split(':');
    return i18n.t(
      rawDirection === 'prev' ? 'sessionHelp.action.cycleOptionPrev' : 'sessionHelp.action.cycleOptionNext',
      { id: rawId || 'option' },
    );
  }

  return i18n.t('sessionHelp.action.mpvCommand', {
    command: command.map((entry) => String(entry)).join(' '),
  });
}

export {
  describeCommand as describeSessionHelpCommand,
  formatKeybinding as formatSessionHelpKeybinding,
};

function sectionForCommand(command: (string | number)[]): string {
  const first = command[0];
  if (typeof first !== 'string') return i18n.t('sessionHelp.sections.other');

  if (
    first === 'cycle' ||
    first === 'seek' ||
    first === 'sub-seek' ||
    first === 'sub-step' ||
    first === SPECIAL_COMMANDS.REPLAY_SUBTITLE ||
    first === SPECIAL_COMMANDS.PLAY_NEXT_SUBTITLE
  ) {
    return i18n.t('sessionHelp.sections.playback');
  }

  if (first === 'show-text' || first === 'show-progress' || first.startsWith('osd')) {
    return i18n.t('sessionHelp.sections.visual');
  }

  if (first === SPECIAL_COMMANDS.SUBSYNC_TRIGGER) {
    return i18n.t('sessionHelp.sections.sync');
  }

  if (
    first === SPECIAL_COMMANDS.RUNTIME_OPTIONS_OPEN ||
    first === SPECIAL_COMMANDS.JIMAKU_OPEN ||
    first === SPECIAL_COMMANDS.PLAYLIST_BROWSER_OPEN ||
    first.startsWith(SPECIAL_COMMANDS.RUNTIME_OPTION_CYCLE_PREFIX)
  ) {
    return i18n.t('sessionHelp.sections.runtime');
  }

  if (first === 'quit') return i18n.t('sessionHelp.sections.system');
  return i18n.t('sessionHelp.sections.other');
}

const MODIFIER_LABELS: Record<SessionKeyModifier, string> = {
  ctrl: 'Ctrl',
  alt: 'Alt',
  shift: 'Shift',
  meta: 'Meta',
};

function formatSessionKeySpec(key: SessionKeySpec): string {
  return [
    ...key.modifiers.map((modifier) => MODIFIER_LABELS[modifier]),
    normalizeKeyToken(key.code),
  ]
    .filter(Boolean)
    .join(' + ');
}

function describeSessionAction(
  actionId: SessionActionId,
  payload?: { runtimeOptionId?: string; direction?: 1 | -1 },
): string {
  switch (actionId) {
    case 'toggleStatsOverlay':
      return i18n.t('sessionHelp.actions.toggleStats');
    case 'toggleVisibleOverlay':
      return i18n.t('sessionHelp.actions.toggleOverlay');
    case 'copySubtitle':
      return i18n.t('sessionHelp.actions.copySubtitle');
    case 'copySubtitleMultiple':
      return i18n.t('sessionHelp.actions.copySubtitleMultiple');
    case 'updateLastCardFromClipboard':
      return i18n.t('sessionHelp.action.updateLastCard');
    case 'triggerFieldGrouping':
      return i18n.t('sessionHelp.action.triggerFieldGrouping');
    case 'triggerSubsync':
      return i18n.t('sessionHelp.actions.openSubsync');
    case 'mineSentence':
      return i18n.t('sessionHelp.actions.mineSentence');
    case 'mineSentenceMultiple':
      return i18n.t('sessionHelp.actions.mineSentenceMultiple');
    case 'toggleSecondarySub':
      return i18n.t('sessionHelp.actions.toggleSecondarySub');
    case 'toggleSubtitleSidebar':
      return i18n.t('sessionHelp.actions.toggleSubtitleSidebar');
    case 'toggleNotificationHistory':
      return i18n.t('sessionHelp.actions.toggleNotificationHistory');
    case 'markAudioCard':
      return i18n.t('sessionHelp.actions.markAudioCard');
    case 'markWatched':
      return i18n.t('sessionHelp.action.markVideoWatched');
    case 'openRuntimeOptions':
      return i18n.t('sessionHelp.actions.openRuntimeOptions');
    case 'openSessionHelp':
      return i18n.t('sessionHelp.actions.openSessionHelp');
    case 'openCharacterDictionaryManager':
      return i18n.t('sessionHelp.actions.openCharacterDictionary');
    case 'openControllerSelect':
      return i18n.t('sessionHelp.actions.openControllerSelect');
    case 'openControllerDebug':
      return i18n.t('sessionHelp.actions.openControllerDebug');
    case 'openJimaku':
      return i18n.t('sessionHelp.actions.openJimaku');
    case 'openYoutubePicker':
      return i18n.t('sessionHelp.action.openYoutubePicker');
    case 'openPlaylistBrowser':
      return i18n.t('sessionHelp.action.openPlaylistBrowserAction');
    case 'replayCurrentSubtitle':
      return i18n.t('sessionHelp.action.replaySubtitle');
    case 'playNextSubtitle':
      return i18n.t('sessionHelp.action.playNextSubtitle');
    case 'cycleRuntimeOption':
      return i18n.t(
        payload?.direction === -1
          ? 'sessionHelp.action.cycleOptionPrev'
          : 'sessionHelp.action.cycleOptionNext',
        { id: payload?.runtimeOptionId ?? 'option' },
      );
  }
}

function sectionForSessionBinding(binding: CompiledSessionBinding): string {
  if (binding.actionType === 'mpv-command') return sectionForCommand(binding.command);

  switch (binding.actionId) {
    case 'copySubtitle':
    case 'copySubtitleMultiple':
    case 'updateLastCardFromClipboard':
    case 'triggerFieldGrouping':
    case 'mineSentence':
    case 'mineSentenceMultiple':
    case 'markAudioCard':
      return i18n.t('sessionHelp.section.mining');
    case 'toggleStatsOverlay':
    case 'markWatched':
      return i18n.t('sessionHelp.section.stats');
    case 'toggleVisibleOverlay':
    case 'toggleSecondarySub':
    case 'toggleSubtitleSidebar':
    case 'toggleNotificationHistory':
      return i18n.t('sessionHelp.section.overlayControls');
    case 'triggerSubsync':
      return i18n.t('sessionHelp.sections.sync');
    case 'openRuntimeOptions':
    case 'openJimaku':
    case 'openCharacterDictionaryManager':
    case 'openControllerSelect':
    case 'openControllerDebug':
    case 'openYoutubePicker':
    case 'openPlaylistBrowser':
    case 'openSessionHelp':
      return i18n.t('sessionHelp.section.modals');
    case 'replayCurrentSubtitle':
    case 'playNextSubtitle':
      return i18n.t('sessionHelp.sections.playback');
    case 'cycleRuntimeOption':
      return i18n.t('sessionHelp.sections.runtime');
  }
}

function buildSessionBindingSections(
  sessionBindings: CompiledSessionBinding[],
): SessionHelpSection[] {
  const grouped = new Map<string, SessionHelpItem[]>();

  for (const binding of sessionBindings) {
    const section = sectionForSessionBinding(binding);
    const row: SessionHelpItem = {
      shortcut: formatSessionKeySpec(binding.key),
      action:
        binding.actionType === 'mpv-command'
          ? describeCommand(binding.command)
          : describeSessionAction(binding.actionId, binding.payload),
    };
    grouped.set(section, [...(grouped.get(section) ?? []), row]);
  }

  const sectionOrder = [
    i18n.t('sessionHelp.sections.playback'),
    i18n.t('sessionHelp.section.mining'),
    i18n.t('sessionHelp.section.stats'),
    i18n.t('sessionHelp.section.overlayControls'),
    i18n.t('sessionHelp.sections.sync'),
    i18n.t('sessionHelp.sections.runtime'),
    i18n.t('sessionHelp.section.modals'),
    i18n.t('sessionHelp.sections.visual'),
    i18n.t('sessionHelp.sections.system'),
    i18n.t('sessionHelp.sections.other'),
  ];
  return Array.from(grouped.entries())
    .sort((a, b) => {
      const aIdx = sectionOrder.indexOf(a[0]);
      const bIdx = sectionOrder.indexOf(b[0]);
      if (aIdx === -1 && bIdx === -1) return a[0].localeCompare(b[0]);
      if (aIdx === -1) return 1;
      if (bIdx === -1) return -1;
      return aIdx - bIdx;
    })
    .map(([title, rows]) => ({ title, rows }));
}

function buildConfiguredOverlaySections(input: {
  markWatchedKey?: string | null;
  subtitleSidebarToggleKey?: string | null;
}): SessionHelpSection[] {
  const statsRows: SessionHelpItem[] = [];
  if (input.markWatchedKey) {
    statsRows.push({
      shortcut: formatKeybinding(input.markWatchedKey),
      action: i18n.t('sessionHelp.action.markVideoWatched'),
    });
  }

  const overlayRows: SessionHelpItem[] = [];
  if (input.subtitleSidebarToggleKey) {
    overlayRows.push({
      shortcut: formatKeybinding(input.subtitleSidebarToggleKey),
      action: i18n.t('sessionHelp.action.toggleSubtitleSidebar'),
    });
  }

  return [
    ...(statsRows.length > 0 ? [{ title: i18n.t('sessionHelp.section.stats'), rows: statsRows }] : []),
    ...(overlayRows.length > 0
      ? [{ title: i18n.t('sessionHelp.section.overlayControls'), rows: overlayRows }]
      : []),
  ];
}

function buildFixedOverlaySections(): SessionHelpSection[] {
  return [
    {
      title: i18n.t('sessionHelp.section.fixedOverlay'),
      rows: [
        {
          shortcut: 'V',
          action: i18n.t('sessionHelp.fixed.togglePrimarySubtitle'),
        },
        {
          shortcut: 'Ctrl/Cmd + A',
          action: i18n.t('sessionHelp.fixed.appendClipboardVideo'),
        },
        {
          shortcut: 'Right-click',
          action: i18n.t('sessionHelp.fixed.togglePlaybackOutsideSubtitle'),
        },
        {
          shortcut: 'Right-click + drag',
          action: i18n.t('sessionHelp.fixed.repositionSubtitles'),
        },
      ],
    },
    {
      title: i18n.t('sessionHelp.section.yChords'),
      rows: [
        { shortcut: 'Y then Y', action: i18n.t('sessionHelp.fixed.openSubMinerMenu') },
        { shortcut: 'Y then S', action: i18n.t('sessionHelp.fixed.startOverlay') },
        { shortcut: 'Y then Shift + S', action: i18n.t('sessionHelp.fixed.stopOverlay') },
        { shortcut: 'Y then T', action: i18n.t('sessionHelp.fixed.toggleVisibleOverlay') },
        { shortcut: 'Y then O', action: i18n.t('sessionHelp.fixed.openYomitanSettings') },
        { shortcut: 'Y then R', action: i18n.t('sessionHelp.fixed.restartOverlay') },
        { shortcut: 'Y then C', action: i18n.t('sessionHelp.fixed.checkOverlayStatus') },
        { shortcut: 'Y then H/K', action: i18n.t('sessionHelp.fixed.openSessionHelp') },
        { shortcut: 'Y then D', action: i18n.t('sessionHelp.fixed.toggleDevTools') },
      ],
    },
    {
      title: i18n.t('sessionHelp.section.globalShortcuts'),
      rows: [
        {
          shortcut: 'Alt + Shift + Y',
          action: i18n.t('sessionHelp.fixed.openYomitanSettingsGlobal'),
        },
      ],
    },
  ];
}

function mergeSectionsByTitle(sections: SessionHelpSection[]): SessionHelpSection[] {
  const merged: SessionHelpSection[] = [];
  const byTitle = new Map<string, SessionHelpSection>();

  for (const section of sections) {
    const existing = byTitle.get(section.title);
    if (existing) {
      existing.rows.push(...section.rows);
      continue;
    }

    const next = { title: section.title, rows: [...section.rows] };
    byTitle.set(section.title, next);
    merged.push(next);
  }

  return merged;
}

export function buildSessionHelpSections(input: {
  sessionBindings: CompiledSessionBinding[];
  markWatchedKey?: string | null;
  subtitleSidebarToggleKey?: string | null;
  subtitleStyle: SessionHelpSubtitleStyle | null | undefined;
}): SessionHelpSection[] {
  const sessionBindings = input.sessionBindings.filter((binding) => {
    if (binding.actionType !== 'session-action') return true;
    if (input.markWatchedKey && binding.actionId === 'markWatched') return false;
    if (input.subtitleSidebarToggleKey && binding.actionId === 'toggleSubtitleSidebar') {
      return false;
    }
    return true;
  });

  return mergeSectionsByTitle([
    ...buildSessionBindingSections(sessionBindings),
    ...buildConfiguredOverlaySections({
      markWatchedKey: input.markWatchedKey,
      subtitleSidebarToggleKey: input.subtitleSidebarToggleKey,
    }),
    ...buildFixedOverlaySections(),
    buildColorSection(input.subtitleStyle ?? {}),
  ]);
}

export function getSessionHelpSectionTabId(section: SessionHelpSection): SessionHelpTabId {
  switch (section.title) {
    case i18n.t('sessionHelp.section.stats'):
    case i18n.t('sessionHelp.section.overlayControls'):
    case i18n.t('sessionHelp.section.fixedOverlay'):
    case i18n.t('sessionHelp.section.globalShortcuts'):
      return 'essentials';
    case i18n.t('sessionHelp.sections.playback'):
    case i18n.t('sessionHelp.sections.sync'):
    case i18n.t('sessionHelp.sections.visual'):
    case i18n.t('sessionHelp.sections.system'):
      return 'playback';
    case i18n.t('sessionHelp.section.mining'):
      return 'mining';
    case i18n.t('sessionHelp.section.modals'):
    case i18n.t('sessionHelp.sections.runtime'):
      return 'tools';
    case i18n.t('sessionHelp.section.yChords'):
    case i18n.t('sessionHelp.colorLegend'):
    case i18n.t('sessionHelp.sections.other'):
    default:
      return 'reference';
  }
}

export function filterSessionHelpSections(
  sections: SessionHelpSection[],
  query: string,
): SessionHelpSection[] {
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
