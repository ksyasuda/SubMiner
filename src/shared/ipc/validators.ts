import type {
  JimakuDownloadQuery,
  JimakuFilesQuery,
  JimakuSearchQuery,
  KikuFieldGroupingChoice,
  KikuMergePreviewRequest,
  RuntimeOptionId,
  RuntimeOptionValue,
  SubtitlePosition,
  SubsyncManualRunRequest,
} from '../../types';
import { OVERLAY_HOSTED_MODALS, type OverlayHostedModal } from './contracts';

const RUNTIME_OPTION_IDS: RuntimeOptionId[] = [
  'anki.autoUpdateNewCards',
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
