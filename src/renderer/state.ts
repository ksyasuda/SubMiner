import type {
  JimakuEntry,
  JimakuFileEntry,
  KikuDuplicateCardInfo,
  KikuFieldGroupingChoice,
  MpvSubtitleRenderMetrics,
  RuntimeOptionId,
  RuntimeOptionState,
  RuntimeOptionValue,
  SubtitlePosition,
  SubsyncSourceTrack,
} from '../types';

export type KikuModalStep = 'select' | 'preview';
export type KikuPreviewMode = 'compact' | 'full';

export type ChordAction =
  | { type: 'mpv'; command: string[] }
  | { type: 'electron'; action: () => void }
  | { type: 'noop' };

export type RendererState = {
  isOverSubtitle: boolean;
  isDragging: boolean;
  dragStartY: number;
  startYPercent: number;
  currentYPercent: number | null;
  persistedSubtitlePosition: SubtitlePosition;

  jimakuModalOpen: boolean;
  jimakuEntries: JimakuEntry[];
  jimakuFiles: JimakuFileEntry[];
  selectedEntryIndex: number;
  selectedFileIndex: number;
  currentEpisodeFilter: number | null;
  currentEntryId: number | null;

  kikuModalOpen: boolean;
  kikuSelectedCard: 1 | 2;
  kikuOriginalData: KikuDuplicateCardInfo | null;
  kikuDuplicateData: KikuDuplicateCardInfo | null;
  kikuModalStep: KikuModalStep;
  kikuPreviewMode: KikuPreviewMode;
  kikuPendingChoice: KikuFieldGroupingChoice | null;
  kikuPreviewCompactData: Record<string, unknown> | null;
  kikuPreviewFullData: Record<string, unknown> | null;

  runtimeOptionsModalOpen: boolean;
  runtimeOptions: RuntimeOptionState[];
  runtimeOptionSelectedIndex: number;
  runtimeOptionDraftValues: Map<RuntimeOptionId, RuntimeOptionValue>;

  subsyncModalOpen: boolean;
  subsyncSourceTracks: SubsyncSourceTrack[];
  subsyncSubmitting: boolean;

  sessionHelpModalOpen: boolean;
  sessionHelpSelectedIndex: number;

  mpvSubtitleRenderMetrics: MpvSubtitleRenderMetrics | null;
  invisiblePositionEditMode: boolean;
  invisiblePositionEditStartX: number;
  invisiblePositionEditStartY: number;
  invisibleSubtitleOffsetXPx: number;
  invisibleSubtitleOffsetYPx: number;
  invisibleLayoutBaseLeftPx: number;
  invisibleLayoutBaseBottomPx: number | null;
  invisibleLayoutBaseTopPx: number | null;
  invisiblePositionEditHud: HTMLDivElement | null;
  currentInvisibleSubtitleLineCount: number;

  lastHoverSelectionKey: string;
  lastHoverSelectionNode: Text | null;
  lastHoveredTokenIndex: number | null;
  invisibleTokenHoverSourceText: string;
  invisibleTokenHoverRanges: Array<{
    start: number;
    end: number;
    tokenIndex: number;
  }>;
  invisibleMeasuredDescentPx: number | null;

  knownWordColor: string;
  nPlusOneColor: string;
  jlptN1Color: string;
  jlptN2Color: string;
  jlptN3Color: string;
  jlptN4Color: string;
  jlptN5Color: string;
  preserveSubtitleLineBreaks: boolean;
  frequencyDictionaryEnabled: boolean;
  frequencyDictionaryTopX: number;
  frequencyDictionaryMode: 'single' | 'banded';
  frequencyDictionarySingleColor: string;
  frequencyDictionaryBand1Color: string;
  frequencyDictionaryBand2Color: string;
  frequencyDictionaryBand3Color: string;
  frequencyDictionaryBand4Color: string;
  frequencyDictionaryBand5Color: string;

  keybindingsMap: Map<string, (string | number)[]>;
  chordPending: boolean;
  chordTimeout: ReturnType<typeof setTimeout> | null;
};

export function createRendererState(): RendererState {
  return {
    isOverSubtitle: false,
    isDragging: false,
    dragStartY: 0,
    startYPercent: 0,
    currentYPercent: null,
    persistedSubtitlePosition: { yPercent: 10 },

    jimakuModalOpen: false,
    jimakuEntries: [],
    jimakuFiles: [],
    selectedEntryIndex: 0,
    selectedFileIndex: 0,
    currentEpisodeFilter: null,
    currentEntryId: null,

    kikuModalOpen: false,
    kikuSelectedCard: 1,
    kikuOriginalData: null,
    kikuDuplicateData: null,
    kikuModalStep: 'select',
    kikuPreviewMode: 'compact',
    kikuPendingChoice: null,
    kikuPreviewCompactData: null,
    kikuPreviewFullData: null,

    runtimeOptionsModalOpen: false,
    runtimeOptions: [],
    runtimeOptionSelectedIndex: 0,
    runtimeOptionDraftValues: new Map(),

    subsyncModalOpen: false,
    subsyncSourceTracks: [],
    subsyncSubmitting: false,

    sessionHelpModalOpen: false,
    sessionHelpSelectedIndex: 0,

    mpvSubtitleRenderMetrics: null,
    invisiblePositionEditMode: false,
    invisiblePositionEditStartX: 0,
    invisiblePositionEditStartY: 0,
    invisibleSubtitleOffsetXPx: 0,
    invisibleSubtitleOffsetYPx: 0,
    invisibleLayoutBaseLeftPx: 0,
    invisibleLayoutBaseBottomPx: null,
    invisibleLayoutBaseTopPx: null,
    invisiblePositionEditHud: null,
    currentInvisibleSubtitleLineCount: 1,

    lastHoverSelectionKey: '',
    lastHoverSelectionNode: null,
    lastHoveredTokenIndex: null,
    invisibleTokenHoverSourceText: '',
    invisibleTokenHoverRanges: [],
    invisibleMeasuredDescentPx: null,

    knownWordColor: '#a6da95',
    nPlusOneColor: '#c6a0f6',
    jlptN1Color: '#ed8796',
    jlptN2Color: '#f5a97f',
    jlptN3Color: '#f9e2af',
    jlptN4Color: '#a6e3a1',
    jlptN5Color: '#8aadf4',
    preserveSubtitleLineBreaks: false,
    frequencyDictionaryEnabled: false,
    frequencyDictionaryTopX: 1000,
    frequencyDictionaryMode: 'single',
    frequencyDictionarySingleColor: '#f5a97f',
    frequencyDictionaryBand1Color: '#ed8796',
    frequencyDictionaryBand2Color: '#f5a97f',
    frequencyDictionaryBand3Color: '#f9e2af',
    frequencyDictionaryBand4Color: '#a6e3a1',
    frequencyDictionaryBand5Color: '#8aadf4',

    keybindingsMap: new Map(),
    chordPending: false,
    chordTimeout: null,
  };
}
