import type { KikuFieldGroupingChoice, KikuMergePreviewRequest } from '../../types/anki';
import type {
  TsukihimeDownloadQuery,
  TsukihimeFilesQuery,
  TsukihimeSearchQuery,
  JimakuDownloadQuery,
  JimakuFilesQuery,
  JimakuSearchQuery,
  YoutubePickerResolveRequest,
} from '../../types/integrations';
import type {
  ControllerConfigUpdate,
  ControllerPreferenceUpdate,
  SessionActionDispatchRequest,
  SubsyncManualRunRequest,
} from '../../types/runtime';
import type { RuntimeOptionId, RuntimeOptionValue } from '../../types/runtime-options';
import type { SessionActionId, SessionActionPayload } from '../../types/session-bindings';
import type { SubtitlePosition } from '../../types/subtitle';
import { OVERLAY_HOSTED_MODALS, type OverlayHostedModal } from './contracts';

const RESERVED_CONTROLLER_PROFILE_IDS = new Set(['__proto__', 'constructor', 'prototype']);

const SESSION_ACTION_IDS: SessionActionId[] = [
  'toggleStatsOverlay',
  'markWatched',
  'toggleVisibleOverlay',
  'copySubtitle',
  'copySubtitleMultiple',
  'updateLastCardFromClipboard',
  'triggerFieldGrouping',
  'triggerSubsync',
  'mineSentence',
  'mineSentenceMultiple',
  'toggleSecondarySub',
  'markAudioCard',
  'toggleSubtitleSidebar',
  'toggleNotificationHistory',
  'appendClipboardVideoToQueue',
  'openRuntimeOptions',
  'openSessionHelp',
  'openCharacterDictionaryManager',
  'openControllerSelect',
  'openControllerDebug',
  'openJimaku',
  'openTsukihime',
  'openYoutubePicker',
  'openPlaylistBrowser',
  'replayCurrentSubtitle',
  'playNextSubtitle',
  'cycleRuntimeOption',
];

const RUNTIME_OPTION_IDS: RuntimeOptionId[] = [
  'anki.autoUpdateNewCards',
  'subtitle.annotation.knownWords.highlightEnabled',
  'subtitle.annotation.knownWords.maturityEnabled',
  'subtitle.annotation.nPlusOne',
  'subtitle.annotation.jlpt',
  'subtitle.annotation.frequency',
  'anki.kikuFieldGrouping',
  'anki.nPlusOneMatchMode',
];

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value);
}

function isInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isInteger(value);
}

function isSessionActionId(value: unknown): value is SessionActionId {
  return typeof value === 'string' && SESSION_ACTION_IDS.includes(value as SessionActionId);
}

function parseSessionActionPayload(
  actionId: SessionActionId,
  value: unknown,
): SessionActionPayload | undefined | null {
  if (actionId === 'copySubtitleMultiple' || actionId === 'mineSentenceMultiple') {
    if (value === undefined) return undefined;
    if (!isObject(value)) return null;
    const keys = Object.keys(value);
    if (keys.some((key) => key !== 'count')) return null;
    if (value.count === undefined) return null;
    if (!isInteger(value.count) || value.count < 1) return null;
    return { count: value.count };
  }

  if (actionId === 'cycleRuntimeOption') {
    if (!isObject(value)) return null;
    const keys = Object.keys(value);
    if (keys.some((key) => key !== 'runtimeOptionId' && key !== 'direction')) return null;
    if (typeof value.runtimeOptionId !== 'string' || value.runtimeOptionId.trim().length === 0) {
      return null;
    }
    if (value.direction !== 1 && value.direction !== -1) {
      return null;
    }
    return {
      runtimeOptionId: value.runtimeOptionId,
      direction: value.direction,
    };
  }

  return value === undefined ? undefined : null;
}

export function parseOverlayHostedModal(value: unknown): OverlayHostedModal | null {
  if (typeof value !== 'string') return null;
  return OVERLAY_HOSTED_MODALS.includes(value as OverlayHostedModal)
    ? (value as OverlayHostedModal)
    : null;
}

export function parseSubtitlePosition(value: unknown): SubtitlePosition | null {
  if (!isObject(value) || !isFiniteNumber(value.yPercent)) {
    return null;
  }
  return {
    yPercent: value.yPercent,
  };
}

export function parseControllerPreferenceUpdate(value: unknown): ControllerPreferenceUpdate | null {
  if (!isObject(value)) return null;
  if (typeof value.preferredGamepadId !== 'string') return null;
  if (typeof value.preferredGamepadLabel !== 'string') return null;
  return {
    preferredGamepadId: value.preferredGamepadId,
    preferredGamepadLabel: value.preferredGamepadLabel,
  };
}

function parseDiscreteBinding(value: unknown) {
  if (!isObject(value) || typeof value.kind !== 'string') return null;
  if (value.kind === 'none') {
    return { kind: 'none' };
  }
  if (value.kind === 'button') {
    if (!isInteger(value.buttonIndex) || value.buttonIndex < 0) return null;
    return { kind: 'button', buttonIndex: value.buttonIndex };
  }
  if (value.kind === 'axis') {
    if (!isInteger(value.axisIndex) || value.axisIndex < 0) return null;
    if (value.direction !== 'negative' && value.direction !== 'positive') return null;
    return { kind: 'axis', axisIndex: value.axisIndex, direction: value.direction };
  }
  return null;
}

function parseAxisBinding(value: unknown) {
  if (isObject(value) && value.kind === 'none') {
    return { kind: 'none' };
  }
  if (!isObject(value) || value.kind !== 'axis') return null;
  if (!isInteger(value.axisIndex) || value.axisIndex < 0) return null;
  if (
    value.dpadFallback !== undefined &&
    value.dpadFallback !== 'none' &&
    value.dpadFallback !== 'horizontal' &&
    value.dpadFallback !== 'vertical'
  ) {
    return null;
  }
  return {
    kind: 'axis',
    axisIndex: value.axisIndex,
    ...(value.dpadFallback === undefined ? {} : { dpadFallback: value.dpadFallback }),
  };
}

function parseControllerButtonIndices(
  value: unknown,
): ControllerConfigUpdate['buttonIndices'] | null {
  if (!isObject(value)) return null;
  const buttonIndices: NonNullable<ControllerConfigUpdate['buttonIndices']> = {};
  const keys = [
    'select',
    'buttonSouth',
    'buttonEast',
    'buttonNorth',
    'buttonWest',
    'leftShoulder',
    'rightShoulder',
    'leftStickPress',
    'rightStickPress',
    'leftTrigger',
    'rightTrigger',
  ] as const;
  for (const key of keys) {
    if (value[key] === undefined) continue;
    if (!isInteger(value[key]) || value[key] < 0) return null;
    buttonIndices[key] = value[key];
  }
  return buttonIndices;
}

function parseControllerBindings(value: unknown): ControllerConfigUpdate['bindings'] | null {
  if (!isObject(value)) return null;
  const bindings: NonNullable<ControllerConfigUpdate['bindings']> = {};
  const discreteKeys = [
    'toggleLookup',
    'closeLookup',
    'toggleKeyboardOnlyMode',
    'mineCard',
    'quitMpv',
    'previousAudio',
    'nextAudio',
    'playCurrentAudio',
    'toggleMpvPause',
  ] as const;
  for (const key of discreteKeys) {
    if (value[key] === undefined) continue;
    const parsed = parseDiscreteBinding(value[key]);
    if (!parsed) return null;
    bindings[key] = parsed as NonNullable<ControllerConfigUpdate['bindings']>[typeof key];
  }
  const axisKeys = [
    'leftStickHorizontal',
    'leftStickVertical',
    'rightStickHorizontal',
    'rightStickVertical',
  ] as const;
  for (const key of axisKeys) {
    if (value[key] === undefined) continue;
    const parsed = parseAxisBinding(value[key]);
    if (!parsed) return null;
    bindings[key] = parsed as NonNullable<ControllerConfigUpdate['bindings']>[typeof key];
  }
  return bindings;
}

export function parseControllerConfigUpdate(value: unknown): ControllerConfigUpdate | null {
  if (!isObject(value)) return null;
  const update: ControllerConfigUpdate = {};

  if (value.enabled !== undefined) {
    if (typeof value.enabled !== 'boolean') return null;
    update.enabled = value.enabled;
  }
  if (value.preferredGamepadId !== undefined) {
    if (typeof value.preferredGamepadId !== 'string') return null;
    update.preferredGamepadId = value.preferredGamepadId;
  }
  if (value.preferredGamepadLabel !== undefined) {
    if (typeof value.preferredGamepadLabel !== 'string') return null;
    update.preferredGamepadLabel = value.preferredGamepadLabel;
  }
  if (value.buttonIndices !== undefined) {
    const parsed = parseControllerButtonIndices(value.buttonIndices);
    if (!parsed) return null;
    update.buttonIndices = parsed;
  }

  if (value.bindings !== undefined) {
    const parsed = parseControllerBindings(value.bindings);
    if (!parsed) return null;
    update.bindings = parsed;
  }

  if (value.profiles !== undefined) {
    if (!isObject(value.profiles)) return null;
    const profiles: NonNullable<ControllerConfigUpdate['profiles']> = Object.create(null);
    for (const [profileId, rawProfile] of Object.entries(value.profiles)) {
      if (RESERVED_CONTROLLER_PROFILE_IDS.has(profileId)) return null;
      if (!isObject(rawProfile)) return null;
      const profile: NonNullable<ControllerConfigUpdate['profiles']>[string] = {};
      if (rawProfile.label !== undefined) {
        if (typeof rawProfile.label !== 'string') return null;
        profile.label = rawProfile.label;
      }
      if (rawProfile.buttonIndices !== undefined) {
        const parsed = parseControllerButtonIndices(rawProfile.buttonIndices);
        if (!parsed) return null;
        profile.buttonIndices = parsed;
      }
      if (rawProfile.bindings !== undefined) {
        const parsed = parseControllerBindings(rawProfile.bindings);
        if (!parsed) return null;
        profile.bindings = parsed;
      }
      profiles[profileId] = profile;
    }
    update.profiles = profiles;
  }

  return update;
}

export function parseSubsyncManualRunRequest(value: unknown): SubsyncManualRunRequest | null {
  if (!isObject(value)) return null;
  const { engine, sourceTrackId } = value;
  if (engine !== 'alass' && engine !== 'ffsubsync') return null;
  if (sourceTrackId !== undefined && sourceTrackId !== null && !isInteger(sourceTrackId)) {
    return null;
  }
  return {
    engine,
    sourceTrackId: sourceTrackId === undefined ? undefined : (sourceTrackId as number | null),
  };
}

export function parseRuntimeOptionId(value: unknown): RuntimeOptionId | null {
  if (typeof value !== 'string') return null;
  return RUNTIME_OPTION_IDS.includes(value as RuntimeOptionId) ? (value as RuntimeOptionId) : null;
}

export function parseRuntimeOptionDirection(value: unknown): 1 | -1 | null {
  return value === 1 || value === -1 ? value : null;
}

export function parseRuntimeOptionValue(value: unknown): RuntimeOptionValue | null {
  return typeof value === 'boolean' || typeof value === 'string'
    ? (value as RuntimeOptionValue)
    : null;
}

export function parseSessionActionDispatchRequest(
  value: unknown,
): SessionActionDispatchRequest | null {
  if (!isObject(value)) return null;
  if (!isSessionActionId(value.actionId)) return null;

  const payload = parseSessionActionPayload(value.actionId, value.payload);
  if (payload === null) return null;
  return payload === undefined
    ? { actionId: value.actionId }
    : { actionId: value.actionId, payload };
}

export function parseMpvCommand(value: unknown): Array<string | number> | null {
  if (!Array.isArray(value)) return null;
  return value.every((entry) => typeof entry === 'string' || typeof entry === 'number')
    ? (value as Array<string | number>)
    : null;
}

export function parseOptionalForwardingOptions(value: unknown): {
  forward?: boolean;
} {
  if (!isObject(value)) return {};
  const { forward } = value;
  if (forward === undefined) return {};
  return typeof forward === 'boolean' ? { forward } : {};
}

export function parseKikuFieldGroupingChoice(value: unknown): KikuFieldGroupingChoice | null {
  if (!isObject(value)) return null;
  const { keepNoteId, deleteNoteId, deleteDuplicate, cancelled } = value;
  if (!isInteger(keepNoteId) || !isInteger(deleteNoteId)) return null;
  if (typeof deleteDuplicate !== 'boolean' || typeof cancelled !== 'boolean') return null;
  return {
    keepNoteId,
    deleteNoteId,
    deleteDuplicate,
    cancelled,
  };
}

export function parseKikuMergePreviewRequest(value: unknown): KikuMergePreviewRequest | null {
  if (!isObject(value)) return null;
  const { keepNoteId, deleteNoteId, deleteDuplicate } = value;
  if (!isInteger(keepNoteId) || !isInteger(deleteNoteId)) return null;
  if (typeof deleteDuplicate !== 'boolean') return null;
  return {
    keepNoteId,
    deleteNoteId,
    deleteDuplicate,
  };
}

export function parseJimakuSearchQuery(value: unknown): JimakuSearchQuery | null {
  if (!isObject(value) || typeof value.query !== 'string') return null;
  return { query: value.query };
}

export function parseJimakuFilesQuery(value: unknown): JimakuFilesQuery | null {
  if (!isObject(value) || !isInteger(value.entryId)) return null;
  if (value.episode !== undefined && value.episode !== null && !isInteger(value.episode)) {
    return null;
  }
  return {
    entryId: value.entryId,
    episode: (value.episode as number | null | undefined) ?? undefined,
  };
}

export function parseJimakuDownloadQuery(value: unknown): JimakuDownloadQuery | null {
  if (!isObject(value)) return null;
  if (
    !isInteger(value.entryId) ||
    typeof value.url !== 'string' ||
    typeof value.name !== 'string'
  ) {
    return null;
  }
  return {
    entryId: value.entryId,
    url: value.url,
    name: value.name,
  };
}

export function parseTsukihimeSearchQuery(value: unknown): TsukihimeSearchQuery | null {
  if (!isObject(value) || typeof value.query !== 'string') return null;
  return { query: value.query };
}

export function parseTsukihimeFilesQuery(value: unknown): TsukihimeFilesQuery | null {
  if (!isObject(value) || !isInteger(value.entryId)) return null;
  return { entryId: value.entryId };
}

export function parseTsukihimeDownloadQuery(value: unknown): TsukihimeDownloadQuery | null {
  if (!isObject(value)) return null;
  if (
    !isInteger(value.entryId) ||
    typeof value.url !== 'string' ||
    typeof value.name !== 'string'
  ) {
    return null;
  }
  if (value.lang !== undefined && typeof value.lang !== 'string') {
    return null;
  }
  return {
    entryId: value.entryId,
    url: value.url,
    name: value.name,
    lang: value.lang,
  };
}

export function parseYoutubePickerResolveRequest(
  value: unknown,
): YoutubePickerResolveRequest | null {
  if (!isObject(value)) return null;
  if (typeof value.sessionId !== 'string' || !value.sessionId.trim()) return null;
  if (value.action !== 'use-selected' && value.action !== 'continue-without-subtitles') return null;
  if (value.action === 'continue-without-subtitles') {
    if (value.primaryTrackId !== null || value.secondaryTrackId !== null) {
      return null;
    }
    return {
      sessionId: value.sessionId,
      action: 'continue-without-subtitles',
      primaryTrackId: null,
      secondaryTrackId: null,
    };
  }
  if (
    value.primaryTrackId !== null &&
    value.primaryTrackId !== undefined &&
    typeof value.primaryTrackId !== 'string'
  ) {
    return null;
  }
  if (
    value.secondaryTrackId !== null &&
    value.secondaryTrackId !== undefined &&
    typeof value.secondaryTrackId !== 'string'
  ) {
    return null;
  }
  return {
    sessionId: value.sessionId,
    action: 'use-selected',
    primaryTrackId: value.primaryTrackId ?? null,
    secondaryTrackId: value.secondaryTrackId ?? null,
  };
}
