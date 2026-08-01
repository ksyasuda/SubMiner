import { MpvSubtitleRenderMetrics } from '../../types';

export type MpvMessage = {
  event?: string;
  name?: string;
  data?: unknown;
  args?: unknown;
  request_id?: number;
  error?: string;
  reason?: unknown;
  file_error?: unknown;
};

export const MPV_REQUEST_ID_SUBTEXT = 101;
export const MPV_REQUEST_ID_PATH = 102;
export const MPV_REQUEST_ID_SECONDARY_SUBTEXT = 103;
export const MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY = 104;
export const MPV_REQUEST_ID_AID = 105;
export const MPV_REQUEST_ID_SUB_POS = 106;
export const MPV_REQUEST_ID_SUB_FONT_SIZE = 107;
export const MPV_REQUEST_ID_SUB_SCALE = 108;
export const MPV_REQUEST_ID_SUB_MARGIN_Y = 109;
export const MPV_REQUEST_ID_SUB_MARGIN_X = 110;
export const MPV_REQUEST_ID_SUB_FONT = 111;
export const MPV_REQUEST_ID_SUB_SCALE_BY_WINDOW = 112;
export const MPV_REQUEST_ID_OSD_HEIGHT = 113;
export const MPV_REQUEST_ID_OSD_DIMENSIONS = 114;
export const MPV_REQUEST_ID_SUBTEXT_ASS = 115;
export const MPV_REQUEST_ID_SUB_SPACING = 116;
export const MPV_REQUEST_ID_SUB_BOLD = 117;
export const MPV_REQUEST_ID_SUB_ITALIC = 118;
export const MPV_REQUEST_ID_SUB_BORDER_SIZE = 119;
export const MPV_REQUEST_ID_SUB_SHADOW_OFFSET = 120;
export const MPV_REQUEST_ID_SUB_ASS_OVERRIDE = 121;
export const MPV_REQUEST_ID_SUB_USE_MARGINS = 122;
export const MPV_REQUEST_ID_PAUSE = 123;
export const MPV_REQUEST_ID_TRACK_LIST_SECONDARY = 200;
export const MPV_REQUEST_ID_TRACK_LIST_AUDIO = 201;

export type MpvMessageParser = (message: MpvMessage) => void;
export type MpvParseErrorHandler = (line: string, error: unknown) => void;

export interface MpvProtocolParseResult {
  messages: MpvMessage[];
  nextBuffer: string;
}

export interface MpvProtocolHandleMessageDeps {
  getResolvedConfig: () => {
    secondarySub?: { secondarySubLanguages?: Array<string> };
  };
  getSubtitleMetrics: () => MpvSubtitleRenderMetrics;
  isVisibleOverlayVisible: () => boolean;
  emitSubtitleChange: (payload: { text: string; isOverlayVisible: boolean }) => void;
  emitSubtitleAssChange: (payload: { text: string }) => void;
  emitSubtitleTiming: (payload: { text: string; start: number; end: number }) => void;
  emitSecondarySubtitleChange: (payload: { text: string }) => void;
  emitSubtitleTrackChange: (payload: { sid: number | null }) => void;
  emitSubtitleTrackListChange: (payload: { trackList: unknown[] | null }) => void;
  getCurrentSubText: () => string;
  setCurrentSubText: (text: string) => void;
  setCurrentSubStart: (value: number) => void;
  getCurrentSubStart: () => number;
  setCurrentSubEnd: (value: number) => void;
  getCurrentSubEnd: () => number;
  emitMediaPathChange: (payload: { path: string }) => void;
  emitMediaTitleChange: (payload: { title: string | null }) => void;
  emitTimePosChange: (payload: { time: number }) => void;
  emitDurationChange: (payload: { duration: number }) => void;
  emitPauseChange: (payload: { paused: boolean }) => void;
  emitFullscreenChange: (payload: { fullscreen: boolean }) => void;
  emitSubtitleMetricsChange: (payload: Partial<MpvSubtitleRenderMetrics>) => void;
  setCurrentSecondarySubText: (text: string) => void;
  resolvePendingRequest: (requestId: number, message: MpvMessage) => boolean;
  setSecondarySubVisibility: (visible: boolean) => void;
  syncCurrentAudioStreamIndex: () => void;
  setCurrentAudioTrackId: (value: number | null) => void;
  setCurrentTimePos: (value: number) => void;
  getCurrentTimePos: () => number;
  getPendingPauseAtSubEnd: () => boolean;
  setPendingPauseAtSubEnd: (value: boolean) => void;
  getPauseAtTime: () => number | null;
  setPauseAtTime: (value: number | null) => void;
  autoLoadSecondarySubTrack: (path: string) => void;
  setCurrentVideoPath: (value: string) => void;
  emitSecondarySubtitleVisibility: (payload: { visible: boolean }) => void;
  setPreviousSecondarySubVisibility: (visible: boolean) => void;
  setCurrentAudioStreamIndex: (
    tracks: Array<{
      type?: string;
      id?: number;
      selected?: boolean;
      'ff-index'?: number;
    }>,
  ) => void;
  sendCommand: (payload: { command: unknown[]; request_id?: number }) => boolean;
  restorePreviousSecondarySubVisibility: () => void;
  shouldQuitOnMpvShutdown: () => boolean;
  requestAppQuit: () => void;
  emitClientMessage?: (payload: { args: string[] }) => void;
  emitEndFile?: (payload: { reason: string; fileError: string | null }) => void;
}

type SubtitleTrackCandidate = {
  id: number;
  lang: string;
  title: string;
  selected: boolean;
  external: boolean;
  externalFilename: string | null;
};

function normalizeSubtitleTrackCandidate(
  track: Record<string, unknown>,
): SubtitleTrackCandidate | null {
  const id =
    typeof track.id === 'number'
      ? track.id
      : typeof track.id === 'string'
        ? Number(track.id.trim())
        : Number.NaN;
  if (!Number.isInteger(id)) {
    return null;
  }

  const externalFilename =
    typeof track['external-filename'] === 'string' && track['external-filename'].trim().length > 0
      ? track['external-filename'].trim()
      : typeof track.external_filename === 'string' && track.external_filename.trim().length > 0
        ? track.external_filename.trim()
        : null;

  return {
    id,
    lang: String(track.lang || '')
      .trim()
      .toLowerCase(),
    title: String(track.title || '')
      .trim()
      .toLowerCase(),
    selected: track.selected === true,
    external: track.external === true,
    externalFilename,
  };
}

function getSubtitleTrackIdentity(track: SubtitleTrackCandidate): string {
  if (track.externalFilename) {
    return `external:${track.externalFilename.toLowerCase()}`;
  }
  if (track.title.length > 0) {
    return `title:${track.title}`;
  }
  return `id:${track.id}`;
}

function isSignsOrSongsSubtitleTrack(track: SubtitleTrackCandidate): boolean {
  const label = `${track.title} ${track.externalFilename ?? ''}`.toLowerCase();
  return /\b(signs?|songs?)\b/.test(label);
}

function pickSecondarySubtitleTrackId(
  tracks: Array<Record<string, unknown>>,
  preferredLanguages: string[],
): number | null {
  const normalizedLanguages = preferredLanguages
    .map((language) => language.trim().toLowerCase())
    .filter((language) => language.length > 0);
  if (normalizedLanguages.length === 0) {
    return null;
  }

  const subtitleTracks = tracks
    .filter((track) => track.type === 'sub')
    .map(normalizeSubtitleTrackCandidate)
    .filter((track): track is SubtitleTrackCandidate => track !== null);

  const dedupedTracks = new Map<string, SubtitleTrackCandidate>();
  for (const track of subtitleTracks) {
    const identity = getSubtitleTrackIdentity(track);
    const existing = dedupedTracks.get(identity);
    if (!existing || (track.selected && !existing.selected)) {
      dedupedTracks.set(identity, track);
    }
  }

  const uniqueTracks = [...dedupedTracks.values()];

  for (const language of normalizedLanguages) {
    const languageTracks = uniqueTracks.filter((track) => track.lang === language);
    if (languageTracks.length === 0) {
      continue;
    }
    const cleanTracks = languageTracks.filter((track) => !isSignsOrSongsSubtitleTrack(track));
    const candidateTracks = cleanTracks.length > 0 ? cleanTracks : languageTracks;

    const selectedMatch = candidateTracks.find((track) => track.selected);
    if (selectedMatch) {
      return selectedMatch.id;
    }

    const match = candidateTracks[0];
    if (match) {
      return match.id;
    }
  }

  return null;
}

export function splitMpvMessagesFromBuffer(
  buffer: string,
  onMessage?: MpvMessageParser,
  onError?: MpvParseErrorHandler,
): MpvProtocolParseResult {
  const lines = buffer.split('\n');
  const nextBuffer = lines.pop() || '';
  const messages: MpvMessage[] = [];

  for (const line of lines) {
    if (!line.trim()) continue;
    try {
      const message = JSON.parse(line) as MpvMessage;
      messages.push(message);
      if (onMessage) {
        onMessage(message);
      }
    } catch (error) {
      if (onError) {
        onError(line, error);
      }
    }
  }

  return {
    messages,
    nextBuffer,
  };
}

export async function dispatchMpvProtocolMessage(
  msg: MpvMessage,
  deps: MpvProtocolHandleMessageDeps,
): Promise<void> {
  if (msg.event === 'property-change') {
    if (msg.name === 'sub-text') {
      const nextSubText = (msg.data as string) || '';
      const overlayVisible = deps.isVisibleOverlayVisible();
      deps.emitSubtitleChange({
        text: nextSubText,
        isOverlayVisible: overlayVisible,
      });
      deps.setCurrentSubText(nextSubText);
    } else if (msg.name === 'sub-text/ass' || msg.name === 'sub-text-ass') {
      deps.emitSubtitleAssChange({ text: (msg.data as string) || '' });
    } else if (msg.name === 'sub-start') {
      deps.setCurrentSubStart((msg.data as number) || 0);
      deps.emitSubtitleTiming({
        text: deps.getCurrentSubText(),
        start: deps.getCurrentSubStart(),
        end: deps.getCurrentSubEnd(),
      });
    } else if (msg.name === 'sub-end') {
      const subEnd = (msg.data as number) || 0;
      deps.setCurrentSubEnd(subEnd);
      if (deps.getPendingPauseAtSubEnd() && subEnd > 0) {
        deps.setPauseAtTime(subEnd);
        deps.setPendingPauseAtSubEnd(false);
        deps.sendCommand({ command: ['set_property', 'pause', false] });
      }
      deps.emitSubtitleTiming({
        text: deps.getCurrentSubText(),
        start: deps.getCurrentSubStart(),
        end: deps.getCurrentSubEnd(),
      });
    } else if (msg.name === 'secondary-sub-text') {
      const nextSubText = (msg.data as string) || '';
      deps.setCurrentSecondarySubText(nextSubText);
      deps.emitSecondarySubtitleChange({ text: nextSubText });
    } else if (msg.name === 'sid') {
      const sid =
        typeof msg.data === 'number'
          ? msg.data
          : typeof msg.data === 'string'
            ? Number(msg.data)
            : null;
      deps.emitSubtitleTrackChange({ sid: sid !== null && Number.isFinite(sid) ? sid : null });
    } else if (msg.name === 'track-list') {
      deps.emitSubtitleTrackListChange({
        trackList: Array.isArray(msg.data) ? (msg.data as unknown[]) : null,
      });
    } else if (msg.name === 'aid') {
      deps.setCurrentAudioTrackId(typeof msg.data === 'number' ? (msg.data as number) : null);
      deps.syncCurrentAudioStreamIndex();
    } else if (msg.name === 'time-pos') {
      const timePos = (msg.data as number) || 0;
      deps.setCurrentTimePos(timePos);
      deps.emitTimePosChange({ time: timePos });
      if (
        deps.getPauseAtTime() !== null &&
        deps.getCurrentTimePos() >= (deps.getPauseAtTime() as number)
      ) {
        deps.setPauseAtTime(null);
        deps.sendCommand({ command: ['set_property', 'pause', true] });
      }
    } else if (msg.name === 'duration') {
      const duration = typeof msg.data === 'number' ? msg.data : 0;
      if (duration > 0) {
        deps.emitDurationChange({ duration });
      }
    } else if (msg.name === 'pause') {
      deps.emitPauseChange({ paused: asBoolean(msg.data, false) });
    } else if (msg.name === 'fullscreen') {
      deps.emitFullscreenChange({ fullscreen: asBoolean(msg.data, false) });
    } else if (msg.name === 'media-title') {
      deps.emitMediaTitleChange({
        title: typeof msg.data === 'string' ? msg.data.trim() : null,
      });
    } else if (msg.name === 'path') {
      const path = (msg.data as string) || '';
      deps.setCurrentVideoPath(path);
      deps.emitMediaPathChange({ path });
      deps.autoLoadSecondarySubTrack(path);
      deps.syncCurrentAudioStreamIndex();
    } else if (msg.name === 'sub-pos') {
      deps.emitSubtitleMetricsChange({ subPos: msg.data as number });
    } else if (msg.name === 'sub-font-size') {
      deps.emitSubtitleMetricsChange({ subFontSize: msg.data as number });
    } else if (msg.name === 'sub-scale') {
      deps.emitSubtitleMetricsChange({ subScale: msg.data as number });
    } else if (msg.name === 'sub-margin-y') {
      deps.emitSubtitleMetricsChange({ subMarginY: msg.data as number });
    } else if (msg.name === 'sub-margin-x') {
      deps.emitSubtitleMetricsChange({ subMarginX: msg.data as number });
    } else if (msg.name === 'sub-font') {
      deps.emitSubtitleMetricsChange({ subFont: msg.data as string });
    } else if (msg.name === 'sub-spacing') {
      deps.emitSubtitleMetricsChange({ subSpacing: msg.data as number });
    } else if (msg.name === 'sub-bold') {
      deps.emitSubtitleMetricsChange({
        subBold: asBoolean(msg.data, deps.getSubtitleMetrics().subBold),
      });
    } else if (msg.name === 'sub-italic') {
      deps.emitSubtitleMetricsChange({
        subItalic: asBoolean(msg.data, deps.getSubtitleMetrics().subItalic),
      });
    } else if (msg.name === 'sub-border-size') {
      deps.emitSubtitleMetricsChange({ subBorderSize: msg.data as number });
    } else if (msg.name === 'sub-shadow-offset') {
      deps.emitSubtitleMetricsChange({ subShadowOffset: msg.data as number });
    } else if (msg.name === 'sub-ass-override') {
      deps.emitSubtitleMetricsChange({ subAssOverride: msg.data as string });
    } else if (msg.name === 'sub-scale-by-window') {
      deps.emitSubtitleMetricsChange({
        subScaleByWindow: asBoolean(msg.data, deps.getSubtitleMetrics().subScaleByWindow),
      });
    } else if (msg.name === 'sub-visibility') {
      if (deps.isVisibleOverlayVisible() && asBoolean(msg.data, false)) {
        deps.sendCommand({ command: ['set_property', 'sub-visibility', false] });
      }
    } else if (msg.name === 'sub-use-margins') {
      deps.emitSubtitleMetricsChange({
        subUseMargins: asBoolean(msg.data, deps.getSubtitleMetrics().subUseMargins),
      });
    } else if (msg.name === 'osd-height') {
      deps.emitSubtitleMetricsChange({ osdHeight: msg.data as number });
    } else if (msg.name === 'osd-dimensions') {
      const dims = msg.data as Record<string, unknown> | null;
      if (!dims) {
        deps.emitSubtitleMetricsChange({ osdDimensions: null });
      } else {
        deps.emitSubtitleMetricsChange({
          osdDimensions: {
            w: asFiniteNumber(dims.w, 0),
            h: asFiniteNumber(dims.h, 0),
            ml: asFiniteNumber(dims.ml, 0),
            mr: asFiniteNumber(dims.mr, 0),
            mt: asFiniteNumber(dims.mt, 0),
            mb: asFiniteNumber(dims.mb, 0),
          },
        });
      }
    }
  } else if (msg.event === 'client-message') {
    const args = Array.isArray(msg.args)
      ? msg.args.filter((arg): arg is string => typeof arg === 'string')
      : [];
    if (args.length > 0) {
      deps.emitClientMessage?.({ args });
    }
  } else if (msg.event === 'end-file') {
    deps.emitEndFile?.({
      reason: typeof msg.reason === 'string' ? msg.reason : 'unknown',
      fileError:
        typeof msg.file_error === 'string' && msg.file_error.length > 0 ? msg.file_error : null,
    });
  } else if (msg.event === 'shutdown') {
    deps.restorePreviousSecondarySubVisibility();
    if (deps.shouldQuitOnMpvShutdown()) {
      deps.requestAppQuit();
    }
  } else if (msg.request_id) {
    if (deps.resolvePendingRequest(msg.request_id, msg)) {
      return;
    }

    if (msg.data === undefined) {
      return;
    }

    if (msg.request_id === MPV_REQUEST_ID_TRACK_LIST_SECONDARY) {
      const tracks = msg.data as Array<{
        type: string;
        lang?: string;
        id: number;
      }>;
      if (Array.isArray(tracks)) {
        const config = deps.getResolvedConfig();
        const languages = config.secondarySub?.secondarySubLanguages || [];
        const secondaryTrackId = pickSecondarySubtitleTrackId(tracks, languages);
        if (secondaryTrackId !== null) {
          deps.sendCommand({
            command: ['set_property', 'secondary-sid', secondaryTrackId],
          });
        }
      }
    } else if (msg.request_id === MPV_REQUEST_ID_TRACK_LIST_AUDIO) {
      deps.setCurrentAudioStreamIndex(
        msg.data as Array<{
          type?: string;
          id?: number;
          selected?: boolean;
          'ff-index'?: number;
        }>,
      );
    } else if (msg.request_id === MPV_REQUEST_ID_SUBTEXT) {
      const nextSubText = (msg.data as string) || '';
      deps.setCurrentSubText(nextSubText);
      deps.emitSubtitleChange({
        text: nextSubText,
        isOverlayVisible: deps.isVisibleOverlayVisible(),
      });
    } else if (msg.request_id === MPV_REQUEST_ID_SUBTEXT_ASS) {
      deps.emitSubtitleAssChange({ text: (msg.data as string) || '' });
    } else if (msg.request_id === MPV_REQUEST_ID_PATH) {
      deps.emitMediaPathChange({ path: (msg.data as string) || '' });
    } else if (msg.request_id === MPV_REQUEST_ID_AID) {
      deps.setCurrentAudioTrackId(typeof msg.data === 'number' ? (msg.data as number) : null);
      deps.syncCurrentAudioStreamIndex();
    } else if (msg.request_id === MPV_REQUEST_ID_SECONDARY_SUBTEXT) {
      const nextSubText = (msg.data as string) || '';
      deps.setCurrentSecondarySubText(nextSubText);
      deps.emitSecondarySubtitleChange({ text: nextSubText });
    } else if (msg.request_id === MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY) {
      const previous = parseVisibilityProperty(msg.data);
      if (previous !== null) {
        deps.setPreviousSecondarySubVisibility(previous);
        deps.emitSecondarySubtitleVisibility({ visible: previous });
      }
      deps.setSecondarySubVisibility(false);
    } else if (msg.request_id === MPV_REQUEST_ID_SUB_POS) {
      deps.emitSubtitleMetricsChange({ subPos: msg.data as number });
    } else if (msg.request_id === MPV_REQUEST_ID_SUB_FONT_SIZE) {
      deps.emitSubtitleMetricsChange({ subFontSize: msg.data as number });
    } else if (msg.request_id === MPV_REQUEST_ID_SUB_SCALE) {
      deps.emitSubtitleMetricsChange({ subScale: msg.data as number });
    } else if (msg.request_id === MPV_REQUEST_ID_SUB_MARGIN_Y) {
      deps.emitSubtitleMetricsChange({ subMarginY: msg.data as number });
    } else if (msg.request_id === MPV_REQUEST_ID_SUB_MARGIN_X) {
      deps.emitSubtitleMetricsChange({ subMarginX: msg.data as number });
    } else if (msg.request_id === MPV_REQUEST_ID_SUB_FONT) {
      deps.emitSubtitleMetricsChange({ subFont: msg.data as string });
    } else if (msg.request_id === MPV_REQUEST_ID_SUB_SPACING) {
      deps.emitSubtitleMetricsChange({ subSpacing: msg.data as number });
    } else if (msg.request_id === MPV_REQUEST_ID_SUB_BOLD) {
      deps.emitSubtitleMetricsChange({
        subBold: asBoolean(msg.data, deps.getSubtitleMetrics().subBold),
      });
    } else if (msg.request_id === MPV_REQUEST_ID_SUB_ITALIC) {
      deps.emitSubtitleMetricsChange({
        subItalic: asBoolean(msg.data, deps.getSubtitleMetrics().subItalic),
      });
    } else if (msg.request_id === MPV_REQUEST_ID_SUB_BORDER_SIZE) {
      deps.emitSubtitleMetricsChange({ subBorderSize: msg.data as number });
    } else if (msg.request_id === MPV_REQUEST_ID_SUB_SHADOW_OFFSET) {
      deps.emitSubtitleMetricsChange({ subShadowOffset: msg.data as number });
    } else if (msg.request_id === MPV_REQUEST_ID_SUB_ASS_OVERRIDE) {
      deps.emitSubtitleMetricsChange({ subAssOverride: msg.data as string });
    } else if (msg.request_id === MPV_REQUEST_ID_SUB_SCALE_BY_WINDOW) {
      deps.emitSubtitleMetricsChange({
        subScaleByWindow: asBoolean(msg.data, deps.getSubtitleMetrics().subScaleByWindow),
      });
    } else if (msg.request_id === MPV_REQUEST_ID_SUB_USE_MARGINS) {
      deps.emitSubtitleMetricsChange({
        subUseMargins: asBoolean(msg.data, deps.getSubtitleMetrics().subUseMargins),
      });
    } else if (msg.request_id === MPV_REQUEST_ID_PAUSE) {
      deps.emitPauseChange({ paused: asBoolean(msg.data, false) });
    } else if (msg.request_id === MPV_REQUEST_ID_OSD_HEIGHT) {
      deps.emitSubtitleMetricsChange({ osdHeight: msg.data as number });
    } else if (msg.request_id === MPV_REQUEST_ID_OSD_DIMENSIONS) {
      const dims = msg.data as Record<string, unknown> | null;
      if (!dims) {
        deps.emitSubtitleMetricsChange({ osdDimensions: null });
      } else {
        deps.emitSubtitleMetricsChange({
          osdDimensions: {
            w: asFiniteNumber(dims.w, 0),
            h: asFiniteNumber(dims.h, 0),
            ml: asFiniteNumber(dims.ml, 0),
            mr: asFiniteNumber(dims.mr, 0),
            mt: asFiniteNumber(dims.mt, 0),
            mb: asFiniteNumber(dims.mb, 0),
          },
        });
      }
    }
  }
}

export function asBoolean(value: unknown, fallback: boolean): boolean {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value !== 0;
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (['yes', 'true', '1'].includes(normalized)) return true;
    if (['no', 'false', '0'].includes(normalized)) return false;
  }
  return fallback;
}

export function asFiniteNumber(value: unknown, fallback: number): number {
  const nextValue = Number(value);
  return Number.isFinite(nextValue) ? nextValue : fallback;
}

export function parseVisibilityProperty(value: unknown): boolean | null {
  if (typeof value === 'boolean') return value;
  if (typeof value !== 'string') return null;

  const normalized = value.trim().toLowerCase();
  if (normalized === 'yes' || normalized === 'true' || normalized === '1') {
    return true;
  }
  if (normalized === 'no' || normalized === 'false' || normalized === '0') {
    return false;
  }

  return null;
}
