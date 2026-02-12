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
} from "../types";

export type KikuModalStep = "select" | "preview";
export type KikuPreviewMode = "compact" | "full";

export type ChordAction =
  | { type: "mpv"; command: string[] }
  | { type: "electron"; action: () => void }
  | { type: "noop" };

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
    kikuModalStep: "select",
    kikuPreviewMode: "compact",
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

    lastHoverSelectionKey: "",
    lastHoverSelectionNode: null,

    keybindingsMap: new Map(),
    chordPending: false,
    chordTimeout: null,
  };
}
