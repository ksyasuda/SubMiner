import * as net from "net";
import { EventEmitter } from "events";
import {
  Config,
  MpvClient,
  MpvSubtitleRenderMetrics,
} from "../../types";
import {
  dispatchMpvProtocolMessage,
  MPV_REQUEST_ID_AID,
  MPV_REQUEST_ID_OSD_DIMENSIONS,
  MPV_REQUEST_ID_OSD_HEIGHT,
  MPV_REQUEST_ID_PATH,
  MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY,
  MPV_REQUEST_ID_SECONDARY_SUBTEXT,
  MPV_REQUEST_ID_SUB_ASS_OVERRIDE,
  MPV_REQUEST_ID_SUB_BOLD,
  MPV_REQUEST_ID_SUB_BORDER_SIZE,
  MPV_REQUEST_ID_SUB_FONT,
  MPV_REQUEST_ID_SUB_FONT_SIZE,
  MPV_REQUEST_ID_SUB_ITALIC,
  MPV_REQUEST_ID_SUB_MARGIN_X,
  MPV_REQUEST_ID_SUB_MARGIN_Y,
  MPV_REQUEST_ID_SUB_POS,
  MPV_REQUEST_ID_SUB_SCALE,
  MPV_REQUEST_ID_SUB_SCALE_BY_WINDOW,
  MPV_REQUEST_ID_SUB_SHADOW_OFFSET,
  MPV_REQUEST_ID_SUB_SPACING,
  MPV_REQUEST_ID_SUBTEXT,
  MPV_REQUEST_ID_SUBTEXT_ASS,
  MPV_REQUEST_ID_SUB_USE_MARGINS,
  MPV_REQUEST_ID_TRACK_LIST_AUDIO,
  MPV_REQUEST_ID_TRACK_LIST_SECONDARY,
  MpvMessage,
  MpvProtocolHandleMessageDeps,
  splitMpvMessagesFromBuffer,
} from "./mpv-protocol";

export {
  MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY,
} from "./mpv-protocol";

export interface MpvIpcClientProtocolDeps {
  getResolvedConfig: () => Config;
  autoStartOverlay: boolean;
  setOverlayVisible: (visible: boolean) => void;
  shouldBindVisibleOverlayToMpvSubVisibility: () => boolean;
  isVisibleOverlayVisible: () => boolean;
  getReconnectTimer: () => ReturnType<typeof setTimeout> | null;
  setReconnectTimer: (timer: ReturnType<typeof setTimeout> | null) => void;
}

export interface MpvIpcClientDeps extends MpvIpcClientProtocolDeps {}

export interface MpvIpcClientEventMap {
  "subtitle-change": { text: string; isOverlayVisible: boolean };
  "subtitle-ass-change": { text: string };
  "subtitle-timing": { text: string; start: number; end: number };
  "secondary-subtitle-change": { text: string };
  "media-path-change": { path: string };
  "media-title-change": { title: string | null };
  "subtitle-metrics-change": { patch: Partial<MpvSubtitleRenderMetrics> };
  "secondary-subtitle-visibility": { visible: boolean };
}

type MpvIpcClientEventName = keyof MpvIpcClientEventMap;

export class MpvIpcClient implements MpvClient {
  private socketPath: string;
  private deps: MpvIpcClientProtocolDeps;
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
  private mpvSubtitleRenderMetrics: MpvSubtitleRenderMetrics = {
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
  };
  private previousSecondarySubVisibility: boolean | null = null;
  private pauseAtTime: number | null = null;
  private pendingPauseAtSubEnd = false;
  private nextDynamicRequestId = 1000;
  private pendingRequests = new Map<number, (message: MpvMessage) => void>();

  constructor(
    socketPath: string,
    deps: MpvIpcClientDeps,
  ) {
    this.socketPath = socketPath;
    this.deps = deps;
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

  private emitSubtitleMetricsChange(
    patch: Partial<MpvSubtitleRenderMetrics>,
  ): void {
    this.mpvSubtitleRenderMetrics = {
      ...this.mpvSubtitleRenderMetrics,
      ...patch,
    };
    this.emit("subtitle-metrics-change", { patch });
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
      this.setSecondarySubVisibility(false);
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
    const parsed = splitMpvMessagesFromBuffer(
      this.buffer,
      (message) => {
        this.handleMessage(message);
      },
      (line, error) => {
        console.error("Failed to parse MPV message:", line, error);
      },
    );
    this.buffer = parsed.nextBuffer;
  }

  private async handleMessage(msg: MpvMessage): Promise<void> {
    await dispatchMpvProtocolMessage(msg, this.createProtocolMessageDeps());
  }

  private createProtocolMessageDeps(): MpvProtocolHandleMessageDeps {
    return {
      getResolvedConfig: () => this.deps.getResolvedConfig(),
      getSubtitleMetrics: () => this.mpvSubtitleRenderMetrics,
      isVisibleOverlayVisible: () => this.deps.isVisibleOverlayVisible(),
      emitSubtitleChange: (payload) => {
        this.emit("subtitle-change", payload);
      },
      emitSubtitleAssChange: (payload) => {
        this.emit("subtitle-ass-change", payload);
      },
      emitSubtitleTiming: (payload) => {
        this.emit("subtitle-timing", payload);
      },
      emitSecondarySubtitleChange: (payload) => {
        this.emit("secondary-subtitle-change", payload);
      },
      getCurrentSubText: () => this.currentSubText,
      setCurrentSubText: (text: string) => {
        this.currentSubText = text;
      },
      setCurrentSubStart: (value: number) => {
        this.currentSubStart = value;
      },
      getCurrentSubStart: () => this.currentSubStart,
      setCurrentSubEnd: (value: number) => {
        this.currentSubEnd = value;
      },
      getCurrentSubEnd: () => this.currentSubEnd,
      emitMediaPathChange: (payload) => {
        this.emit("media-path-change", payload);
      },
      emitMediaTitleChange: (payload) => {
        this.emit("media-title-change", payload);
      },
      emitSubtitleMetricsChange: (patch) => {
        this.emitSubtitleMetricsChange(patch);
      },
      setCurrentSecondarySubText: (text: string) => {
        this.currentSecondarySubText = text;
      },
      resolvePendingRequest: (requestId: number, message: MpvMessage) =>
        this.tryResolvePendingRequest(requestId, message),
      setSecondarySubVisibility: (visible: boolean) =>
        this.setSecondarySubVisibility(visible),
      syncCurrentAudioStreamIndex: () => {
        this.syncCurrentAudioStreamIndex();
      },
      setCurrentAudioTrackId: (value: number | null) => {
        this.currentAudioTrackId = value;
      },
      setCurrentTimePos: (value: number) => {
        this.currentTimePos = value;
      },
      getCurrentTimePos: () => this.currentTimePos,
      getPendingPauseAtSubEnd: () => this.pendingPauseAtSubEnd,
      setPendingPauseAtSubEnd: (value: boolean) => {
        this.pendingPauseAtSubEnd = value;
      },
      getPauseAtTime: () => this.pauseAtTime,
      setPauseAtTime: (value: number | null) => {
        this.pauseAtTime = value;
      },
      autoLoadSecondarySubTrack: () => {
        this.autoLoadSecondarySubTrack();
      },
      setCurrentVideoPath: (value: string) => {
        this.currentVideoPath = value;
      },
      emitSecondarySubtitleVisibility: (payload) => {
        this.emit("secondary-subtitle-visibility", payload);
      },
      setCurrentAudioStreamIndex: (tracks) => {
        this.updateCurrentAudioStreamIndex(tracks);
      },
      sendCommand: (payload) => this.send(payload),
      restorePreviousSecondarySubVisibility: () => {
        this.restorePreviousSecondarySubVisibility();
      },
    };
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

  private tryResolvePendingRequest(
    requestId: number,
    message: MpvMessage,
  ): boolean {
    const pending = this.pendingRequests.get(requestId);
    if (!pending) {
      return false;
    }
    this.pendingRequests.delete(requestId);
    pending(message);
    return true;
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
    const previous = this.previousSecondarySubVisibility;
    if (previous === null) return;
    this.send({
      command: ["set_property", "secondary-sub-visibility", previous ? "yes" : "no"],
    });
    this.previousSecondarySubVisibility = null;
  }

  private setSecondarySubVisibility(visible: boolean): void {
    this.send({
      command: [
        "set_property",
        "secondary-sub-visibility",
        visible ? "yes" : "no",
      ],
    });
  }
}
