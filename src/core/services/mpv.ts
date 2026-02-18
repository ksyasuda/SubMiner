import { EventEmitter } from "events";
import { Config, MpvClient, MpvSubtitleRenderMetrics } from "../../types";
import {
  dispatchMpvProtocolMessage,
  MPV_REQUEST_ID_TRACK_LIST_AUDIO,
  MPV_REQUEST_ID_TRACK_LIST_SECONDARY,
  MpvMessage,
  MpvProtocolHandleMessageDeps,
  splitMpvMessagesFromBuffer,
} from "./mpv-protocol";
import {
  requestMpvInitialState,
  subscribeToMpvProperties,
} from "./mpv-properties";
import { scheduleMpvReconnect, MpvSocketTransport } from "./mpv-transport";
import { createLogger } from "../../logger";

const logger = createLogger("main:mpv");

export type MpvTrackProperty = {
  type?: string;
  id?: number;
  selected?: boolean;
  "ff-index"?: number;
};

export function resolveCurrentAudioStreamIndex(
  tracks: Array<MpvTrackProperty> | null | undefined,
  currentAudioTrackId: number | null,
): number | null {
  if (!Array.isArray(tracks)) {
    return null;
  }

  const audioTracks = tracks.filter((track) => track.type === "audio");
  const activeTrack =
    audioTracks.find((track) => track.id === currentAudioTrackId) ||
    audioTracks.find((track) => track.selected === true);

  const ffIndex = activeTrack?.["ff-index"];
  return typeof ffIndex === "number" &&
    Number.isInteger(ffIndex) &&
    ffIndex >= 0
    ? ffIndex
    : null;
}

export interface MpvRuntimeClientLike {
  connected: boolean;
  send: (payload: { command: (string | number)[] }) => void;
  replayCurrentSubtitle?: () => void;
  playNextSubtitle?: () => void;
  setSubVisibility?: (visible: boolean) => void;
}

export function showMpvOsdRuntime(
  mpvClient: MpvRuntimeClientLike | null,
  text: string,
  fallbackLog: (text: string) => void = (line) => logger.info(line),
): void {
  if (mpvClient && mpvClient.connected) {
    mpvClient.send({ command: ["show-text", text, "3000"] });
    return;
  }
  fallbackLog(`OSD (MPV not connected): ${text}`);
}

export function replayCurrentSubtitleRuntime(
  mpvClient: MpvRuntimeClientLike | null,
): void {
  if (!mpvClient?.replayCurrentSubtitle) return;
  mpvClient.replayCurrentSubtitle();
}

export function playNextSubtitleRuntime(
  mpvClient: MpvRuntimeClientLike | null,
): void {
  if (!mpvClient?.playNextSubtitle) return;
  mpvClient.playNextSubtitle();
}

export function sendMpvCommandRuntime(
  mpvClient: MpvRuntimeClientLike | null,
  command: (string | number)[],
): void {
  if (!mpvClient) return;
  mpvClient.send({ command });
}

export function setMpvSubVisibilityRuntime(
  mpvClient: MpvRuntimeClientLike | null,
  visible: boolean,
): void {
  if (!mpvClient?.setSubVisibility) return;
  mpvClient.setSubVisibility(visible);
}

export { MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY } from "./mpv-protocol";

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
  "connection-change": { connected: boolean };
  "subtitle-change": { text: string; isOverlayVisible: boolean };
  "subtitle-ass-change": { text: string };
  "subtitle-timing": { text: string; start: number; end: number };
  "time-pos-change": { time: number };
  "pause-change": { paused: boolean };
  "secondary-subtitle-change": { text: string };
  "media-path-change": { path: string };
  "media-title-change": { title: string | null };
  "subtitle-metrics-change": { patch: Partial<MpvSubtitleRenderMetrics> };
  "secondary-subtitle-visibility": { visible: boolean };
}

type MpvIpcClientEventName = keyof MpvIpcClientEventMap;

export class MpvIpcClient implements MpvClient {
  private deps: MpvIpcClientProtocolDeps;
  private transport: MpvSocketTransport;
  public socket: ReturnType<MpvSocketTransport["getSocket"]> = null;
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

  constructor(socketPath: string, deps: MpvIpcClientDeps) {
    this.deps = deps;

    this.transport = new MpvSocketTransport({
      socketPath,
      onConnect: () => {
        logger.debug("Connected to MPV socket");
        this.connected = true;
        this.connecting = false;
        this.socket = this.transport.getSocket();
        this.emit("connection-change", { connected: true });
        this.reconnectAttempt = 0;
        this.hasConnectedOnce = true;
        this.setSecondarySubVisibility(false);
        subscribeToMpvProperties(this.send.bind(this));
        requestMpvInitialState(this.send.bind(this));

        const shouldAutoStart =
          this.deps.autoStartOverlay ||
          this.deps.getResolvedConfig().auto_start_overlay === true;
        if (this.firstConnection && shouldAutoStart) {
          logger.debug("Auto-starting overlay, hiding mpv subtitles");
          setTimeout(() => {
            this.deps.setOverlayVisible(true);
          }, 100);
        } else if (this.deps.shouldBindVisibleOverlayToMpvSubVisibility()) {
          this.setSubVisibility(!this.deps.isVisibleOverlayVisible());
        }

        this.firstConnection = false;
      },
      onData: (data) => {
        this.buffer += data.toString();
        this.processBuffer();
      },
      onError: (err: Error) => {
        logger.debug("MPV socket error:", err.message);
        this.failPendingRequests();
      },
      onClose: () => {
        logger.debug("MPV socket closed");
        this.connected = false;
        this.connecting = false;
        this.socket = null;
        this.emit("connection-change", { connected: false });
        this.failPendingRequests();
        this.scheduleReconnect();
      },
    });
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
    this.transport.setSocketPath(socketPath);
  }

  connect(): void {
    if (this.connected || this.connecting) {
      logger.debug(
        `MPV IPC connect request skipped; connected=${this.connected}, connecting=${this.connecting}`,
      );
      return;
    }

    logger.info("MPV IPC connect requested.");
    this.connecting = true;
    this.transport.connect();
  }

  private scheduleReconnect(): void {
    this.reconnectAttempt = scheduleMpvReconnect({
      attempt: this.reconnectAttempt,
      hasConnectedOnce: this.hasConnectedOnce,
      getReconnectTimer: () => this.deps.getReconnectTimer(),
      setReconnectTimer: (timer) => this.deps.setReconnectTimer(timer),
      onReconnectAttempt: (attempt, delay) => {
        logger.debug(
          `Attempting to reconnect to MPV (attempt ${attempt}, delay ${delay}ms)...`,
        );
      },
      connect: () => {
        this.connect();
      },
    });
  }

  private processBuffer(): void {
    const parsed = splitMpvMessagesFromBuffer(
      this.buffer,
      (message) => {
        this.handleMessage(message);
      },
      (line, error) => {
        logger.error("Failed to parse MPV message:", line, error);
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
      emitTimePosChange: (payload) => {
        this.emit("time-pos-change", payload);
      },
      emitPauseChange: (payload) => {
        this.emit("pause-change", payload);
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
      setPreviousSecondarySubVisibility: (visible: boolean) => {
        this.previousSecondarySubVisibility = visible;
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
    this.currentAudioStreamIndex = resolveCurrentAudioStreamIndex(
      tracks,
      this.currentAudioTrackId,
    );
  }

  send(command: { command: unknown[]; request_id?: number }): boolean {
    if (!this.connected || !this.socket) {
      return false;
    }
    return this.transport.send(command);
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
      command: [
        "set_property",
        "secondary-sub-visibility",
        previous ? "yes" : "no",
      ],
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
