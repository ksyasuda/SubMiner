import type {
  CompiledSessionBinding,
  SessionActionId,
  SessionKeyModifier,
  SessionKeySpec,
} from '../../types';
import { SPECIAL_COMMANDS } from '../../config/definitions/shared';
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

export const SESSION_HELP_TABS: SessionHelpTab[] = [
  { id: 'essentials', label: 'Essentials' },
  { id: 'playback', label: 'Playback' },
  { id: 'mining', label: 'Mining' },
  { id: 'tools', label: 'Tools' },
  { id: 'reference', label: 'Reference' },
];

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
  if (first === 'sub-step' && typeof command[1] === 'number') {
    if (command[1] > 0) return 'Shift subtitle delay to next cue';
    if (command[1] < 0) return 'Shift subtitle delay to previous cue';
    return 'Reload current subtitle timing';
  }
  if (first === SPECIAL_COMMANDS.SUBSYNC_TRIGGER) return 'Open subtitle sync controls';
  if (first === SPECIAL_COMMANDS.RUNTIME_OPTIONS_OPEN) return 'Open runtime options';
  if (first === SPECIAL_COMMANDS.JIMAKU_OPEN) return 'Open jimaku';
  if (first === SPECIAL_COMMANDS.PLAYLIST_BROWSER_OPEN) return 'Open playlist browser';
  if (first === SPECIAL_COMMANDS.REPLAY_SUBTITLE) return 'Replay current subtitle';
  if (first === SPECIAL_COMMANDS.PLAY_NEXT_SUBTITLE) return 'Play next subtitle';
  if (first.startsWith(SPECIAL_COMMANDS.RUNTIME_OPTION_CYCLE_PREFIX)) {
    const [, rawId, rawDirection] = first.split(':');
    return `Cycle runtime option ${rawId || 'option'} ${
      rawDirection === 'prev' ? 'previous' : 'next'
    }`;
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
    first === 'sub-step' ||
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
      return 'Toggle stats overlay';
    case 'toggleVisibleOverlay':
      return 'Show/hide visible overlay';
    case 'copySubtitle':
      return 'Copy subtitle';
    case 'copySubtitleMultiple':
      return 'Copy subtitle (multi)';
    case 'updateLastCardFromClipboard':
      return 'Update last card from clipboard';
    case 'triggerFieldGrouping':
      return 'Trigger field grouping';
    case 'triggerSubsync':
      return 'Open subtitle sync controls';
    case 'mineSentence':
      return 'Mine sentence';
    case 'mineSentenceMultiple':
      return 'Mine sentence (multi)';
    case 'toggleSecondarySub':
      return 'Toggle secondary subtitle mode';
    case 'toggleSubtitleSidebar':
      return 'Toggle subtitle sidebar';
    case 'toggleNotificationHistory':
      return 'Toggle notification history';
    case 'markAudioCard':
      return 'Mark audio card';
    case 'markWatched':
      return 'Mark video watched';
    case 'openRuntimeOptions':
      return 'Open runtime options';
    case 'openSessionHelp':
      return 'Open session help';
    case 'openCharacterDictionaryManager':
      return 'Open character dictionary manager';
    case 'openControllerSelect':
      return 'Open controller select';
    case 'openControllerDebug':
      return 'Open controller debug';
    case 'openJimaku':
      return 'Open jimaku';
    case 'openYoutubePicker':
      return 'Open YouTube subtitle picker';
    case 'openPlaylistBrowser':
      return 'Open playlist browser';
    case 'replayCurrentSubtitle':
      return 'Replay current subtitle';
    case 'playNextSubtitle':
      return 'Play next subtitle';
    case 'cycleRuntimeOption':
      return `Cycle runtime option ${payload?.runtimeOptionId ?? 'option'} ${
        payload?.direction === -1 ? 'previous' : 'next'
      }`;
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
      return 'Mining and capture';
    case 'toggleStatsOverlay':
    case 'markWatched':
      return 'Stats and progress';
    case 'toggleVisibleOverlay':
    case 'toggleSecondarySub':
    case 'toggleSubtitleSidebar':
    case 'toggleNotificationHistory':
      return 'Overlay controls';
    case 'triggerSubsync':
      return 'Subtitle sync';
    case 'openRuntimeOptions':
    case 'openJimaku':
    case 'openCharacterDictionaryManager':
    case 'openControllerSelect':
    case 'openControllerDebug':
    case 'openYoutubePicker':
    case 'openPlaylistBrowser':
    case 'openSessionHelp':
      return 'Modals and tools';
    case 'replayCurrentSubtitle':
    case 'playNextSubtitle':
      return 'Playback and navigation';
    case 'cycleRuntimeOption':
      return 'Runtime settings';
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
    'Playback and navigation',
    'Mining and capture',
    'Stats and progress',
    'Overlay controls',
    'Subtitle sync',
    'Runtime settings',
    'Modals and tools',
    'Visual feedback',
    'System actions',
    'Other shortcuts',
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
      action: 'Mark video watched',
    });
  }

  const overlayRows: SessionHelpItem[] = [];
  if (input.subtitleSidebarToggleKey) {
    overlayRows.push({
      shortcut: formatKeybinding(input.subtitleSidebarToggleKey),
      action: 'Toggle subtitle sidebar',
    });
  }

  return [
    ...(statsRows.length > 0 ? [{ title: 'Stats and progress', rows: statsRows }] : []),
    ...(overlayRows.length > 0 ? [{ title: 'Overlay controls', rows: overlayRows }] : []),
  ];
}

function buildFixedOverlaySections(): SessionHelpSection[] {
  return [
    {
      title: 'Fixed overlay controls',
      rows: [
        { shortcut: 'V', action: 'Toggle primary subtitle bar visibility' },
        { shortcut: 'Ctrl/Cmd + A', action: 'Append clipboard video path to playlist' },
        { shortcut: 'Right-click', action: 'Toggle playback outside subtitle area' },
        { shortcut: 'Right-click + drag', action: 'Reposition subtitles on subtitle area' },
      ],
    },
    {
      title: 'Y chords',
      rows: [
        { shortcut: 'Y then Y', action: 'Open SubMiner menu' },
        { shortcut: 'Y then S', action: 'Start overlay' },
        { shortcut: 'Y then Shift + S', action: 'Stop overlay' },
        { shortcut: 'Y then T', action: 'Toggle visible overlay' },
        { shortcut: 'Y then O', action: 'Open Yomitan settings' },
        { shortcut: 'Y then R', action: 'Restart overlay' },
        { shortcut: 'Y then C', action: 'Check overlay status' },
        { shortcut: 'Y then H/K', action: 'Open session help' },
        { shortcut: 'Y then D', action: 'Toggle DevTools' },
      ],
    },
    {
      title: 'Global shortcuts',
      rows: [{ shortcut: 'Alt + Shift + Y', action: 'Open Yomitan settings' }],
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
    case 'Stats and progress':
    case 'Overlay controls':
    case 'Fixed overlay controls':
    case 'Global shortcuts':
      return 'essentials';
    case 'Playback and navigation':
    case 'Subtitle sync':
    case 'Visual feedback':
    case 'System actions':
      return 'playback';
    case 'Mining and capture':
      return 'mining';
    case 'Modals and tools':
    case 'Runtime settings':
      return 'tools';
    case 'Y chords':
    case 'Color legend':
    case 'Other shortcuts':
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
