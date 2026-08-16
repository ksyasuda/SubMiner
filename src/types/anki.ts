import type { AiFeatureConfig } from './integrations';
import type { NotificationType } from './notification';
import type { NPlusOneMatchMode } from './subtitle';

/**
 * Card that a Kiku/Lapis note generates. The note types mark this with mutually
 * exclusive `Is...Card` flag fields, so only one kind may be flagged per note.
 */
export type CardKind = 'sentence' | 'audio' | 'word-and-sentence' | 'click';

/** Card kind SubMiner flags on word cards; 'none' leaves the flag fields untouched. */
export type WordCardKind = CardKind | 'none';

export type MediaTimingReviewKind = 'word' | 'sentence' | 'audio';

export interface MediaTimingReviewRequest {
  kind: MediaTimingReviewKind;
  text: string;
  startTime: number;
  endTime: number;
  noteId?: number;
  audioPadding: number;
  maxMediaDuration: number;
}

export type MediaTimingReviewDecision =
  | { action: 'confirm'; startTime: number; endTime: number }
  | { action: 'use-original' }
  | { action: 'discard' };

export interface MediaTimingReviewOpenPayload {
  reviewId: string;
  kind: MediaTimingReviewKind;
  text: string;
  noteId?: number;
  originalStartTime: number;
  originalEndTime: number;
  selectionStartTime: number;
  selectionEndTime: number;
  timelineStartTime: number;
  timelineEndTime: number;
  mediaDuration?: number;
  maxMediaDuration: number;
}

export interface MediaTimingReviewPreviewRequest {
  reviewId: string;
  startTime: number;
  endTime: number;
}

export interface MediaTimingReviewWaveformRequest {
  reviewId: string;
  startTime: number;
  endTime: number;
}

export interface MediaTimingReviewWaveformResult extends MediaTimingReviewActionResult {
  peaks?: number[];
}

export interface MediaTimingReviewResolveRequest {
  reviewId: string;
  decision: MediaTimingReviewDecision;
}

export interface MediaTimingReviewActionResult {
  ok: boolean;
  message?: string;
}

export interface NotificationOptions {
  body?: string;
  icon?: string;
}

export interface KikuDuplicateCardInfo {
  noteId: number;
  expression: string;
  sentencePreview: string;
  hasAudio: boolean;
  hasImage: boolean;
  isOriginal: boolean;
}

export interface KikuFieldGroupingRequestData {
  original: KikuDuplicateCardInfo;
  duplicate: KikuDuplicateCardInfo;
}

export interface KikuFieldGroupingChoice {
  keepNoteId: number;
  deleteNoteId: number;
  deleteDuplicate: boolean;
  cancelled: boolean;
}

export interface KikuMergePreviewRequest {
  keepNoteId: number;
  deleteNoteId: number;
  deleteDuplicate: boolean;
}

export interface KikuMergePreviewResponse {
  ok: boolean;
  compact?: Record<string, unknown>;
  full?: Record<string, unknown>;
  error?: string;
}

export interface AnkiConnectConfig {
  enabled?: boolean;
  url?: string;
  pollingRate?: number;
  proxy?: {
    enabled?: boolean;
    host?: string;
    port?: number;
    upstreamUrl?: string;
  };
  tags?: string[];
  fields?: {
    word?: string;
    audio?: string;
    image?: string;
    sentence?: string;
    miscInfo?: string;
    translation?: string;
  };
  ai?: boolean | AiFeatureConfig;
  media?: {
    generateAudio?: boolean;
    generateImage?: boolean;
    imageType?: 'static' | 'avif';
    imageFormat?: 'jpg' | 'png' | 'webp';
    imageQuality?: number;
    imageMaxWidth?: number;
    imageMaxHeight?: number;
    animatedFps?: number;
    animatedMaxWidth?: number;
    animatedMaxHeight?: number;
    animatedCrf?: number;
    syncAnimatedImageToWordAudio?: boolean;
    normalizeAudio?: boolean;
    mirrorMpvVolume?: boolean;
    reviewTiming?: boolean;
    audioPadding?: number;
    fallbackDuration?: number;
    maxMediaDuration?: number;
  };
  knownWords?: {
    highlightEnabled?: boolean;
    maturityEnabled?: boolean;
    matureThresholdDays?: number;
    refreshMinutes?: number;
    addMinedWordsImmediately?: boolean;
    matchMode?: NPlusOneMatchMode;
    decks?: Record<string, string[]>;
    color?: string;
  };
  nPlusOne?: {
    enabled?: boolean;
    minSentenceWords?: number;
  };
  behavior?: {
    overwriteAudio?: boolean;
    overwriteImage?: boolean;
    mediaInsertMode?: 'append' | 'prepend';
    highlightWord?: boolean;
    notificationType?: NotificationType;
    autoUpdateNewCards?: boolean;
  };
  metadata?: {
    pattern?: string;
  };
  deck?: string;
  isLapis?: {
    enabled?: boolean;
    sentenceCardModel?: string;
  };
  isKiku?: {
    enabled?: boolean;
    fieldGrouping?: 'auto' | 'manual' | 'disabled';
    deleteDuplicateInAuto?: boolean;
  };
  lapisKiku?: {
    wordCardKind?: WordCardKind;
  };
}
