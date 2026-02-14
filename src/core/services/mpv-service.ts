import * as net from "net";
import { EventEmitter } from "events";
import {
  Config,
  MpvClient,
  MpvSubtitleRenderMetrics,
  SubtitleData,
} from "../../types";
import { asBoolean, asFiniteNumber } from "../utils/coerce";

interface MpvMessage {
  event?: string;
  name?: string;
  data?: unknown;
  request_id?: number;
  error?: string;
}

const MPV_REQUEST_ID_SUBTEXT = 101;
const MPV_REQUEST_ID_PATH = 102;
const MPV_REQUEST_ID_SECONDARY_SUBTEXT = 103;
export const MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY = 104;
const MPV_REQUEST_ID_AID = 105;
const MPV_REQUEST_ID_SUB_POS = 106;
const MPV_REQUEST_ID_SUB_FONT_SIZE = 107;
const MPV_REQUEST_ID_SUB_SCALE = 108;
const MPV_REQUEST_ID_SUB_MARGIN_Y = 109;
const MPV_REQUEST_ID_SUB_MARGIN_X = 110;
const MPV_REQUEST_ID_SUB_FONT = 111;
const MPV_REQUEST_ID_SUB_SCALE_BY_WINDOW = 112;
const MPV_REQUEST_ID_OSD_HEIGHT = 113;
const MPV_REQUEST_ID_OSD_DIMENSIONS = 114;
const MPV_REQUEST_ID_SUBTEXT_ASS = 115;
const MPV_REQUEST_ID_SUB_SPACING = 116;
const MPV_REQUEST_ID_SUB_BOLD = 117;
const MPV_REQUEST_ID_SUB_ITALIC = 118;
const MPV_REQUEST_ID_SUB_BORDER_SIZE = 119;
const MPV_REQUEST_ID_SUB_SHADOW_OFFSET = 120;
const MPV_REQUEST_ID_SUB_ASS_OVERRIDE = 121;
const MPV_REQUEST_ID_SUB_USE_MARGINS = 122;
const MPV_REQUEST_ID_TRACK_LIST_SECONDARY = 200;
const MPV_REQUEST_ID_TRACK_LIST_AUDIO = 201;

interface SubtitleTimingTrackerLike {
  recordSubtitle: (text: string, start: number, end: number) => void;
}

export interface MpvIpcClientProtocolDeps {
  getResolvedConfig: () => Config;
  autoStartOverlay: boolean;
  setOverlayVisible: (visible: boolean) => void;
  shouldBindVisibleOverlayToMpvSubVisibility: () => boolean;
  isVisibleOverlayVisible: () => boolean;
  getReconnectTimer: () => ReturnType<typeof setTimeout> | null;
  setReconnectTimer: (timer: ReturnType<typeof setTimeout> | null) => void;
}

export interface MpvIpcClientRuntimeDeps {
  getCurrentSubText?: () => string;
  setCurrentSubText?: (text: string) => void;
  setCurrentSubAssText?: (text: string) => void;
  getSubtitleTimingTracker?: () => SubtitleTimingTrackerLike | null;
  subtitleWsBroadcast?: (text: string) => void;
  getOverlayWindowsCount?: () => number;
  tokenizeSubtitle?: (text: string) => Promise<SubtitleData>;
  broadcastToOverlayWindows?: (channel: string, ...args: unknown[]) => void;
  updateCurrentMediaPath?: (mediaPath: unknown) => void;
  updateMpvSubtitleRenderMetrics?: (
    patch: Partial<MpvSubtitleRenderMetrics>,
  ) => void;
  getMpvSubtitleRenderMetrics?: () => MpvSubtitleRenderMetrics;
  getPreviousSecondarySubVisibility?: () => boolean | null;
  setPreviousSecondarySubVisibility?: (value: boolean | null) => void;
  showMpvOsd?: (text: string) => void;
  updateCurrentMediaTitle?: (mediaTitle: unknown) => void;
}

export interface MpvIpcClientDeps extends MpvIpcClientProtocolDeps, MpvIpcClientRuntimeDeps {}

export interface MpvIpcClientEventMap {
  "subtitle-change": { text: string; isOverlayVisible: boolean };
  "subtitle-ass-change": { text: string };
  "secondary-subtitle-change": { text: string };
  "media-path-change": { path: string };
  "media-title-change": { title: string | null };
  "subtitle-metrics-change": { patch: Partial<MpvSubtitleRenderMetrics> };
  "secondary-subtitle-visibility": { visible: boolean };
}

type MpvIpcClientEventName = keyof MpvIpcClientEventMap;

export class MpvIpcClient implements MpvClient {
  private socketPath: string;
  private deps: MpvIpcClientProtocolDeps & Required<MpvIpcClientRuntimeDeps>;
  public socket: net.Socket | null = null;
  private eventBus = new EventEmitter();
  private buffer = "";
  public connected = false;
  private connecting = false;
  private reconnectAttempt = 0;
  private firstConnection = true;
  private hasConnectedOnce = false;
  public currentVideoPath = "";
  public currentTimePos = 0;
  public currentSubStart = 0;
  public currentSubEnd = 0;
  public currentSubText = "";
  public currentSecondarySubText = "";
  public currentAudioStreamIndex: number | null = null;
  private currentAudioTrackId: number | null = null;
  private pauseAtTime: number | null = null;
  private pendingPauseAtSubEnd = false;
  private nextDynamicRequestId = 1000;
  private pendingRequests = new Map<number, (message: MpvMessage) => void>();

  constructor(socketPath: string, deps: MpvIpcClientDeps) {
    this.socketPath = socketPath;
    this.deps = {
      getCurrentSubText: () => "",
      setCurrentSubText: () => undefined,
      setCurrentSubAssText: () => undefined,
      getSubtitleTimingTracker: () => null,
      subtitleWsBroadcast: () => undefined,
      getOverlayWindowsCount: () => 0,
      tokenizeSubtitle: async (text) => ({ text, tokens: null }),
      broadcastToOverlayWindows: () => undefined,
      updateCurrentMediaPath: () => undefined,
      updateCurrentMediaTitle: () => undefined,
      updateMpvSubtitleRenderMetrics: () => undefined,
      getMpvSubtitleRenderMetrics: () => ({
        subPos: 100,
        subFontSize: 36,
        subScale: 1,
        subMarginY: 0,
        subMarginX: 0,
        subFont: "",
        subSpacing: 0,
        subBold: false,
        subItalic: false,
        subBorderSize: 0,
        subShadowOffset: 0,
        subAssOverride: "yes",
        subScaleByWindow: true,
        subUseMargins: true,
        osdHeight: 0,
        osdDimensions: null,
      }),
      getPreviousSecondarySubVisibility: () => null,
      setPreviousSecondarySubVisibility: () => undefined,
      showMpvOsd: () => undefined,
      ...deps,
    };
  }

  on<EventName extends MpvIpcClientEventName>(
    event: EventName,
    listener: (payload: MpvIpcClientEventMap[EventName]) => void,
  ): void {
    this.eventBus.on(event as string, listener);
  }

  off<EventName extends MpvIpcClientEventName>(
    event: EventName,
    listener: (payload: MpvIpcClientEventMap[EventName]) => void,
  ): void {
    this.eventBus.off(event as string, listener);
  }

  private emit<EventName extends MpvIpcClientEventName>(
    event: EventName,
    payload: MpvIpcClientEventMap[EventName],
  ): void {
    this.eventBus.emit(event as string, payload);
  }

  setSocketPath(socketPath: string): void {
    this.socketPath = socketPath;
  }

  connect(): void {
    if (this.connected || this.connecting) {
      return;
    }

    if (this.socket) {
      this.socket.destroy();
    }

    this.connecting = true;
    this.socket = new net.Socket();

    this.socket.on("connect", () => {
      console.log("Connected to MPV socket");
      this.connected = true;
      this.connecting = false;
      this.reconnectAttempt = 0;
      this.hasConnectedOnce = true;
      this.subscribeToProperties();
      this.getInitialState();

      const shouldAutoStart =
        this.deps.autoStartOverlay ||
        this.deps.getResolvedConfig().auto_start_overlay === true;
      if (this.firstConnection && shouldAutoStart) {
        console.log("Auto-starting overlay, hiding mpv subtitles");
        setTimeout(() => {
          this.deps.setOverlayVisible(true);
        }, 100);
      } else if (this.deps.shouldBindVisibleOverlayToMpvSubVisibility()) {
        this.setSubVisibility(!this.deps.isVisibleOverlayVisible());
      }

      this.firstConnection = false;
    });

    this.socket.on("data", (data: Buffer) => {
      this.buffer += data.toString();
      this.processBuffer();
    });

    this.socket.on("error", (err: Error) => {
      console.error("MPV socket error:", err.message);
      this.connected = false;
      this.connecting = false;
      this.failPendingRequests();
    });

    this.socket.on("close", () => {
      console.log("MPV socket closed");
      this.connected = false;
      this.connecting = false;
      this.failPendingRequests();
      this.scheduleReconnect();
    });

    this.socket.connect(this.socketPath);
  }

  private scheduleReconnect(): void {
    const reconnectTimer = this.deps.getReconnectTimer();
    if (reconnectTimer) {
      clearTimeout(reconnectTimer);
    }
    const attempt = this.reconnectAttempt++;
    let delay: number;
    if (this.hasConnectedOnce) {
      if (attempt < 2) {
        delay = 1000;
      } else if (attempt < 4) {
        delay = 2000;
      } else if (attempt < 7) {
        delay = 5000;
      } else {
        delay = 10000;
      }
    } else {
      if (attempt < 2) {
        delay = 200;
      } else if (attempt < 4) {
        delay = 500;
      } else if (attempt < 6) {
        delay = 1000;
      } else {
        delay = 2000;
      }
    }
    this.deps.setReconnectTimer(
      setTimeout(() => {
        console.log(
          `Attempting to reconnect to MPV (attempt ${attempt + 1}, delay ${delay}ms)...`,
        );
        this.connect();
      }, delay),
    );
  }

  private processBuffer(): void {
    const lines = this.buffer.split("\n");
    this.buffer = lines.pop() || "";

    for (const line of lines) {
      if (!line.trim()) continue;
      try {
        const msg = JSON.parse(line) as MpvMessage;
        this.handleMessage(msg);
      } catch (e) {
        console.error("Failed to parse MPV message:", line, e);
      }
    }
  }

  private async handleMessage(msg: MpvMessage): Promise<void> {
    if (msg.event === "property-change") {
      if (msg.name === "sub-text") {
        const nextSubText = (msg.data as string) || "";
        const overlayVisible = this.deps.isVisibleOverlayVisible();
        this.emit("subtitle-change", {
          text: nextSubText,
          isOverlayVisible: overlayVisible,
        });
        this.deps.setCurrentSubText(nextSubText);
        this.currentSubText = nextSubText;
        const subtitleTimingTracker = this.deps.getSubtitleTimingTracker();
        if (
          subtitleTimingTracker &&
          this.currentSubStart !== undefined &&
          this.currentSubEnd !== undefined
        ) {
          subtitleTimingTracker.recordSubtitle(
            nextSubText,
            this.currentSubStart,
            this.currentSubEnd,
          );
        }
        this.deps.subtitleWsBroadcast(nextSubText);
        if (this.deps.getOverlayWindowsCount() > 0) {
          const subtitleData = await this.deps.tokenizeSubtitle(nextSubText);
          this.deps.broadcastToOverlayWindows("subtitle:set", subtitleData);
        }
      } else if (msg.name === "sub-text-ass") {
        const nextSubAssText = (msg.data as string) || "";
        this.emit("subtitle-ass-change", {
          text: nextSubAssText,
        });
        this.deps.setCurrentSubAssText(nextSubAssText);
        this.deps.broadcastToOverlayWindows("subtitle-ass:set", nextSubAssText);
      } else if (msg.name === "sub-start") {
        this.currentSubStart = (msg.data as number) || 0;
        const subtitleTimingTracker = this.deps.getSubtitleTimingTracker();
        if (subtitleTimingTracker && this.deps.getCurrentSubText()) {
          subtitleTimingTracker.recordSubtitle(
            this.deps.getCurrentSubText(),
            this.currentSubStart,
            this.currentSubEnd,
          );
        }
      } else if (msg.name === "sub-end") {
        this.currentSubEnd = (msg.data as number) || 0;
        if (this.pendingPauseAtSubEnd && this.currentSubEnd > 0) {
          this.pauseAtTime = this.currentSubEnd;
          this.pendingPauseAtSubEnd = false;
          this.send({ command: ["set_property", "pause", false] });
        }
        const subtitleTimingTracker = this.deps.getSubtitleTimingTracker();
        if (subtitleTimingTracker && this.deps.getCurrentSubText()) {
          subtitleTimingTracker.recordSubtitle(
            this.deps.getCurrentSubText(),
            this.currentSubStart,
            this.currentSubEnd,
          );
        }
      } else if (msg.name === "secondary-sub-text") {
        this.currentSecondarySubText = (msg.data as string) || "";
        this.emit("secondary-subtitle-change", {
          text: this.currentSecondarySubText,
        });
        this.deps.broadcastToOverlayWindows(
          "secondary-subtitle:set",
          this.currentSecondarySubText,
        );
      } else if (msg.name === "aid") {
        this.currentAudioTrackId =
          typeof msg.data === "number" ? (msg.data as number) : null;
        this.syncCurrentAudioStreamIndex();
      } else if (msg.name === "time-pos") {
        this.currentTimePos = (msg.data as number) || 0;
        if (
          this.pauseAtTime !== null &&
          this.currentTimePos >= this.pauseAtTime
        ) {
          this.pauseAtTime = null;
          this.send({ command: ["set_property", "pause", true] });
        }
      } else if (msg.name === "media-title") {
        this.emit("media-title-change", {
          title: typeof msg.data === "string" ? msg.data.trim() : null,
        });
        this.deps.updateCurrentMediaTitle?.(msg.data);
      } else if (msg.name === "path") {
        this.currentVideoPath = (msg.data as string) || "";
        this.emit("media-path-change", {
          path: (msg.data as string) || "",
        });
        this.deps.updateCurrentMediaPath(msg.data);
        this.autoLoadSecondarySubTrack();
        this.syncCurrentAudioStreamIndex();
      } else if (msg.name === "sub-pos") {
        const patch = { subPos: msg.data as number };
        this.emit("subtitle-metrics-change", { patch });
        this.deps.updateMpvSubtitleRenderMetrics({ subPos: msg.data as number });
      } else if (msg.name === "sub-font-size") {
        this.deps.updateMpvSubtitleRenderMetrics({
          subFontSize: msg.data as number,
        });
      } else if (msg.name === "sub-scale") {
        this.deps.updateMpvSubtitleRenderMetrics({ subScale: msg.data as number });
      } else if (msg.name === "sub-margin-y") {
        this.deps.updateMpvSubtitleRenderMetrics({
          subMarginY: msg.data as number,
        });
      } else if (msg.name === "sub-margin-x") {
        this.deps.updateMpvSubtitleRenderMetrics({
          subMarginX: msg.data as number,
        });
      } else if (msg.name === "sub-font") {
        this.deps.updateMpvSubtitleRenderMetrics({ subFont: msg.data as string });
      } else if (msg.name === "sub-spacing") {
        this.deps.updateMpvSubtitleRenderMetrics({
          subSpacing: msg.data as number,
        });
      } else if (msg.name === "sub-bold") {
        this.deps.updateMpvSubtitleRenderMetrics({
          subBold: asBoolean(msg.data, this.deps.getMpvSubtitleRenderMetrics().subBold),
        });
      } else if (msg.name === "sub-italic") {
        this.deps.updateMpvSubtitleRenderMetrics({
          subItalic: asBoolean(
            msg.data,
            this.deps.getMpvSubtitleRenderMetrics().subItalic,
          ),
        });
      } else if (msg.name === "sub-border-size") {
        this.deps.updateMpvSubtitleRenderMetrics({
          subBorderSize: msg.data as number,
        });
      } else if (msg.name === "sub-shadow-offset") {
        this.deps.updateMpvSubtitleRenderMetrics({
          subShadowOffset: msg.data as number,
        });
      } else if (msg.name === "sub-ass-override") {
        this.deps.updateMpvSubtitleRenderMetrics({
          subAssOverride: msg.data as string,
        });
      } else if (msg.name === "sub-scale-by-window") {
        this.deps.updateMpvSubtitleRenderMetrics({
          subScaleByWindow: asBoolean(
            msg.data,
            this.deps.getMpvSubtitleRenderMetrics().subScaleByWindow,
          ),
        });
      } else if (msg.name === "sub-use-margins") {
        this.deps.updateMpvSubtitleRenderMetrics({
          subUseMargins: asBoolean(
            msg.data,
            this.deps.getMpvSubtitleRenderMetrics().subUseMargins,
          ),
        });
      } else if (msg.name === "osd-height") {
        this.deps.updateMpvSubtitleRenderMetrics({
          osdHeight: msg.data as number,
        });
      } else if (msg.name === "osd-dimensions") {
        const dims = msg.data as Record<string, unknown> | null;
        if (!dims) {
          this.deps.updateMpvSubtitleRenderMetrics({ osdDimensions: null });
        } else {
          this.deps.updateMpvSubtitleRenderMetrics({
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
    } else if (msg.event === "shutdown") {
      this.restorePreviousSecondarySubVisibility();
    } else if (msg.request_id) {
      const pending = this.pendingRequests.get(msg.request_id);
      if (pending) {
        this.pendingRequests.delete(msg.request_id);
        pending(msg);
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
          const config = this.deps.getResolvedConfig();
          const languages = config.secondarySub?.secondarySubLanguages || [];
          const subTracks = tracks.filter((t) => t.type === "sub");
          for (const lang of languages) {
            const match = subTracks.find((t) => t.lang === lang);
            if (match) {
              this.send({
                command: ["set_property", "secondary-sid", match.id],
              });
            //   this.deps.showMpvOsd(
            //     `Secondary subtitle: ${lang} (track ${match.id})`,
            //   );
              break;
            }
          }
        }
      } else if (msg.request_id === MPV_REQUEST_ID_TRACK_LIST_AUDIO) {
        this.updateCurrentAudioStreamIndex(
          msg.data as Array<{
            type?: string;
            id?: number;
            selected?: boolean;
            "ff-index"?: number;
          }>,
        );
      } else if (msg.request_id === MPV_REQUEST_ID_SUBTEXT) {
        const nextSubText = (msg.data as string) || "";
        this.deps.setCurrentSubText(nextSubText);
        this.currentSubText = nextSubText;
        this.emit("subtitle-change", {
          text: nextSubText,
          isOverlayVisible: this.deps.isVisibleOverlayVisible(),
        });
        this.deps.subtitleWsBroadcast(nextSubText);
        if (this.deps.getOverlayWindowsCount() > 0) {
          this.deps.tokenizeSubtitle(nextSubText).then((subtitleData) => {
            this.deps.broadcastToOverlayWindows("subtitle:set", subtitleData);
          });
        }
      } else if (msg.request_id === MPV_REQUEST_ID_SUBTEXT_ASS) {
        const nextSubAssText = (msg.data as string) || "";
        this.emit("subtitle-ass-change", {
          text: nextSubAssText,
        });
        this.deps.setCurrentSubAssText(nextSubAssText);
        this.deps.broadcastToOverlayWindows("subtitle-ass:set", nextSubAssText);
      } else if (msg.request_id === MPV_REQUEST_ID_PATH) {
        this.emit("media-path-change", {
          path: (msg.data as string) || "",
        });
        this.deps.updateCurrentMediaPath(msg.data);
      } else if (msg.request_id === MPV_REQUEST_ID_AID) {
        this.currentAudioTrackId =
          typeof msg.data === "number" ? (msg.data as number) : null;
        this.syncCurrentAudioStreamIndex();
      } else if (msg.request_id === MPV_REQUEST_ID_SECONDARY_SUBTEXT) {
        this.currentSecondarySubText = (msg.data as string) || "";
        this.deps.broadcastToOverlayWindows(
          "secondary-subtitle:set",
          this.currentSecondarySubText,
        );
      } else if (msg.request_id === MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY) {
        this.deps.setPreviousSecondarySubVisibility(
          msg.data === true || msg.data === "yes",
        );
        this.send({
          command: ["set_property", "secondary-sub-visibility", "no"],
        });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_POS) {
        this.deps.updateMpvSubtitleRenderMetrics({ subPos: msg.data as number });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_FONT_SIZE) {
        this.deps.updateMpvSubtitleRenderMetrics({
          subFontSize: msg.data as number,
        });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_SCALE) {
        this.deps.updateMpvSubtitleRenderMetrics({ subScale: msg.data as number });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_MARGIN_Y) {
        this.deps.updateMpvSubtitleRenderMetrics({
          subMarginY: msg.data as number,
        });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_MARGIN_X) {
        this.deps.updateMpvSubtitleRenderMetrics({
          subMarginX: msg.data as number,
        });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_FONT) {
        this.deps.updateMpvSubtitleRenderMetrics({ subFont: msg.data as string });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_SPACING) {
        this.deps.updateMpvSubtitleRenderMetrics({
          subSpacing: msg.data as number,
        });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_BOLD) {
        this.deps.updateMpvSubtitleRenderMetrics({
          subBold: asBoolean(msg.data, this.deps.getMpvSubtitleRenderMetrics().subBold),
        });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_ITALIC) {
        this.deps.updateMpvSubtitleRenderMetrics({
          subItalic: asBoolean(
            msg.data,
            this.deps.getMpvSubtitleRenderMetrics().subItalic,
          ),
        });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_BORDER_SIZE) {
        this.deps.updateMpvSubtitleRenderMetrics({
          subBorderSize: msg.data as number,
        });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_SHADOW_OFFSET) {
        this.deps.updateMpvSubtitleRenderMetrics({
          subShadowOffset: msg.data as number,
        });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_ASS_OVERRIDE) {
        this.deps.updateMpvSubtitleRenderMetrics({
          subAssOverride: msg.data as string,
        });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_SCALE_BY_WINDOW) {
        this.deps.updateMpvSubtitleRenderMetrics({
          subScaleByWindow: asBoolean(
            msg.data,
            this.deps.getMpvSubtitleRenderMetrics().subScaleByWindow,
          ),
        });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_USE_MARGINS) {
        this.deps.updateMpvSubtitleRenderMetrics({
          subUseMargins: asBoolean(
            msg.data,
            this.deps.getMpvSubtitleRenderMetrics().subUseMargins,
          ),
        });
      } else if (msg.request_id === MPV_REQUEST_ID_OSD_HEIGHT) {
        this.deps.updateMpvSubtitleRenderMetrics({
          osdHeight: msg.data as number,
        });
      } else if (msg.request_id === MPV_REQUEST_ID_OSD_DIMENSIONS) {
        const dims = msg.data as Record<string, unknown> | null;
        if (!dims) {
          this.deps.updateMpvSubtitleRenderMetrics({ osdDimensions: null });
        } else {
          this.deps.updateMpvSubtitleRenderMetrics({
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

  private autoLoadSecondarySubTrack(): void {
    const config = this.deps.getResolvedConfig();
    if (!config.secondarySub?.autoLoadSecondarySub) return;
    const languages = config.secondarySub.secondarySubLanguages;
    if (!languages || languages.length === 0) return;

    setTimeout(() => {
      this.send({
        command: ["get_property", "track-list"],
        request_id: MPV_REQUEST_ID_TRACK_LIST_SECONDARY,
      });
    }, 500);
  }

  private syncCurrentAudioStreamIndex(): void {
    this.send({
      command: ["get_property", "track-list"],
      request_id: MPV_REQUEST_ID_TRACK_LIST_AUDIO,
    });
  }

  private updateCurrentAudioStreamIndex(
    tracks: Array<{
      type?: string;
      id?: number;
      selected?: boolean;
      "ff-index"?: number;
    }>,
  ): void {
    if (!Array.isArray(tracks)) {
      this.currentAudioStreamIndex = null;
      return;
    }

    const audioTracks = tracks.filter((track) => track.type === "audio");
    const activeTrack =
      audioTracks.find((track) => track.id === this.currentAudioTrackId) ||
      audioTracks.find((track) => track.selected === true);

    const ffIndex = activeTrack?.["ff-index"];
    this.currentAudioStreamIndex =
      typeof ffIndex === "number" && Number.isInteger(ffIndex) && ffIndex >= 0
        ? ffIndex
        : null;
  }

  send(command: { command: unknown[]; request_id?: number }): boolean {
    if (!this.connected || !this.socket) {
      return false;
    }
    const msg = JSON.stringify(command) + "\n";
    this.socket.write(msg);
    return true;
  }

  request(command: unknown[]): Promise<MpvMessage> {
    return new Promise((resolve, reject) => {
      if (!this.connected || !this.socket) {
        reject(new Error("MPV not connected"));
        return;
      }

      const requestId = this.nextDynamicRequestId++;
      this.pendingRequests.set(requestId, resolve);
      const sent = this.send({ command, request_id: requestId });
      if (!sent) {
        this.pendingRequests.delete(requestId);
        reject(new Error("Failed to send MPV request"));
        return;
      }

      setTimeout(() => {
        if (this.pendingRequests.delete(requestId)) {
          reject(new Error("MPV request timed out"));
        }
      }, 4000);
    });
  }

  async requestProperty(name: string): Promise<unknown> {
    const response = await this.request(["get_property", name]);
    if (response.error && response.error !== "success") {
      throw new Error(
        `Failed to read MPV property '${name}': ${response.error}`,
      );
    }
    return response.data;
  }

  private failPendingRequests(): void {
    for (const [requestId, resolve] of this.pendingRequests.entries()) {
      resolve({ request_id: requestId, error: "disconnected" });
    }
    this.pendingRequests.clear();
  }

  private subscribeToProperties(): void {
    this.send({ command: ["observe_property", 1, "sub-text"] });
    this.send({ command: ["observe_property", 2, "path"] });
    this.send({ command: ["observe_property", 3, "sub-start"] });
    this.send({ command: ["observe_property", 4, "sub-end"] });
    this.send({ command: ["observe_property", 5, "time-pos"] });
    this.send({ command: ["observe_property", 6, "secondary-sub-text"] });
    this.send({ command: ["observe_property", 7, "aid"] });
    this.send({ command: ["observe_property", 8, "sub-pos"] });
    this.send({ command: ["observe_property", 9, "sub-font-size"] });
    this.send({ command: ["observe_property", 10, "sub-scale"] });
    this.send({ command: ["observe_property", 11, "sub-margin-y"] });
    this.send({ command: ["observe_property", 12, "sub-margin-x"] });
    this.send({ command: ["observe_property", 13, "sub-font"] });
    this.send({ command: ["observe_property", 14, "sub-spacing"] });
    this.send({ command: ["observe_property", 15, "sub-bold"] });
    this.send({ command: ["observe_property", 16, "sub-italic"] });
    this.send({ command: ["observe_property", 17, "sub-scale-by-window"] });
    this.send({ command: ["observe_property", 18, "osd-height"] });
    this.send({ command: ["observe_property", 19, "osd-dimensions"] });
    this.send({ command: ["observe_property", 20, "sub-text-ass"] });
    this.send({ command: ["observe_property", 21, "sub-border-size"] });
    this.send({ command: ["observe_property", 22, "sub-shadow-offset"] });
    this.send({ command: ["observe_property", 23, "sub-ass-override"] });
    this.send({ command: ["observe_property", 24, "sub-use-margins"] });
    this.send({ command: ["observe_property", 25, "media-title"] });
  }

  private getInitialState(): void {
    this.send({
      command: ["get_property", "sub-text"],
      request_id: MPV_REQUEST_ID_SUBTEXT,
    });
    this.send({
      command: ["get_property", "sub-text-ass"],
      request_id: MPV_REQUEST_ID_SUBTEXT_ASS,
    });
    this.send({
      command: ["get_property", "path"],
      request_id: MPV_REQUEST_ID_PATH,
    });
    this.send({
      command: ["get_property", "media-title"],
    });
    this.send({
      command: ["get_property", "secondary-sub-text"],
      request_id: MPV_REQUEST_ID_SECONDARY_SUBTEXT,
    });
    this.send({
      command: ["get_property", "secondary-sub-visibility"],
      request_id: MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY,
    });
    this.send({
      command: ["get_property", "aid"],
      request_id: MPV_REQUEST_ID_AID,
    });
    this.send({
      command: ["get_property", "sub-pos"],
      request_id: MPV_REQUEST_ID_SUB_POS,
    });
    this.send({
      command: ["get_property", "sub-font-size"],
      request_id: MPV_REQUEST_ID_SUB_FONT_SIZE,
    });
    this.send({
      command: ["get_property", "sub-scale"],
      request_id: MPV_REQUEST_ID_SUB_SCALE,
    });
    this.send({
      command: ["get_property", "sub-margin-y"],
      request_id: MPV_REQUEST_ID_SUB_MARGIN_Y,
    });
    this.send({
      command: ["get_property", "sub-margin-x"],
      request_id: MPV_REQUEST_ID_SUB_MARGIN_X,
    });
    this.send({
      command: ["get_property", "sub-font"],
      request_id: MPV_REQUEST_ID_SUB_FONT,
    });
    this.send({
      command: ["get_property", "sub-spacing"],
      request_id: MPV_REQUEST_ID_SUB_SPACING,
    });
    this.send({
      command: ["get_property", "sub-bold"],
      request_id: MPV_REQUEST_ID_SUB_BOLD,
    });
    this.send({
      command: ["get_property", "sub-italic"],
      request_id: MPV_REQUEST_ID_SUB_ITALIC,
    });
    this.send({
      command: ["get_property", "sub-scale-by-window"],
      request_id: MPV_REQUEST_ID_SUB_SCALE_BY_WINDOW,
    });
    this.send({
      command: ["get_property", "osd-height"],
      request_id: MPV_REQUEST_ID_OSD_HEIGHT,
    });
    this.send({
      command: ["get_property", "osd-dimensions"],
      request_id: MPV_REQUEST_ID_OSD_DIMENSIONS,
    });
    this.send({
      command: ["get_property", "sub-border-size"],
      request_id: MPV_REQUEST_ID_SUB_BORDER_SIZE,
    });
    this.send({
      command: ["get_property", "sub-shadow-offset"],
      request_id: MPV_REQUEST_ID_SUB_SHADOW_OFFSET,
    });
    this.send({
      command: ["get_property", "sub-ass-override"],
      request_id: MPV_REQUEST_ID_SUB_ASS_OVERRIDE,
    });
    this.send({
      command: ["get_property", "sub-use-margins"],
      request_id: MPV_REQUEST_ID_SUB_USE_MARGINS,
    });
  }

  setSubVisibility(visible: boolean): void {
    this.send({
      command: ["set_property", "sub-visibility", visible ? "yes" : "no"],
    });
  }

  replayCurrentSubtitle(): void {
    this.pendingPauseAtSubEnd = true;
    this.send({ command: ["sub-seek", 0] });
  }

  playNextSubtitle(): void {
    this.pendingPauseAtSubEnd = true;
    this.send({ command: ["sub-seek", 1] });
  }

  restorePreviousSecondarySubVisibility(): void {
    const previous = this.deps.getPreviousSecondarySubVisibility();
    if (previous === null) return;
    this.send({
      command: ["set_property", "secondary-sub-visibility", previous ? "yes" : "no"],
    });
    this.deps.setPreviousSecondarySubVisibility(null);
  }
}
