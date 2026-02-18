import WebSocket from 'ws';

export interface JellyfinRemoteSessionMessage {
  MessageType?: string;
  Data?: unknown;
}

export interface JellyfinTimelinePlaybackState {
  itemId: string;
  mediaSourceId?: string;
  positionTicks?: number;
  playbackStartTimeTicks?: number;
  isPaused?: boolean;
  isMuted?: boolean;
  canSeek?: boolean;
  volumeLevel?: number;
  playbackRate?: number;
  playMethod?: string;
  audioStreamIndex?: number | null;
  subtitleStreamIndex?: number | null;
  playlistItemId?: string | null;
  eventName?: string;
}

export interface JellyfinTimelinePayload {
  ItemId: string;
  MediaSourceId?: string;
  PositionTicks: number;
  PlaybackStartTimeTicks: number;
  IsPaused: boolean;
  IsMuted: boolean;
  CanSeek: boolean;
  VolumeLevel: number;
  PlaybackRate: number;
  PlayMethod: string;
  AudioStreamIndex?: number | null;
  SubtitleStreamIndex?: number | null;
  PlaylistItemId?: string | null;
  EventName: string;
}

interface JellyfinRemoteSocket {
  on(event: 'open', listener: () => void): this;
  on(event: 'close', listener: () => void): this;
  on(event: 'error', listener: (error: Error) => void): this;
  on(event: 'message', listener: (data: unknown) => void): this;
  close(): void;
}

type JellyfinRemoteSocketHeaders = Record<string, string>;

export interface JellyfinRemoteSessionServiceOptions {
  serverUrl: string;
  accessToken: string;
  deviceId: string;
  capabilities?: {
    PlayableMediaTypes?: string;
    SupportedCommands?: string;
    SupportsMediaControl?: boolean;
  };
  onPlay?: (payload: unknown) => void;
  onPlaystate?: (payload: unknown) => void;
  onGeneralCommand?: (payload: unknown) => void;
  fetchImpl?: typeof fetch;
  webSocketFactory?: (url: string) => JellyfinRemoteSocket;
  socketHeadersFactory?: (
    url: string,
    headers: JellyfinRemoteSocketHeaders,
  ) => JellyfinRemoteSocket;
  setTimer?: typeof setTimeout;
  clearTimer?: typeof clearTimeout;
  reconnectBaseDelayMs?: number;
  reconnectMaxDelayMs?: number;
  clientName?: string;
  clientVersion?: string;
  deviceName?: string;
  onConnected?: () => void;
  onDisconnected?: () => void;
}

function normalizeServerUrl(serverUrl: string): string {
  return serverUrl.trim().replace(/\/+$/, '');
}

function clampVolume(value: number | undefined): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) return 100;
  return Math.max(0, Math.min(100, Math.round(value)));
}

function normalizeTicks(value: number | undefined): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) return 0;
  return Math.max(0, Math.floor(value));
}

function parseMessageData(value: unknown): unknown {
  if (typeof value !== 'string') return value;
  const trimmed = value.trim();
  if (!trimmed) return value;
  try {
    return JSON.parse(trimmed);
  } catch {
    return value;
  }
}

function parseInboundMessage(rawData: unknown): JellyfinRemoteSessionMessage | null {
  const serialized =
    typeof rawData === 'string'
      ? rawData
      : Buffer.isBuffer(rawData)
        ? rawData.toString('utf8')
        : null;
  if (!serialized) return null;
  try {
    const parsed = JSON.parse(serialized) as JellyfinRemoteSessionMessage;
    if (!parsed || typeof parsed !== 'object') return null;
    return parsed;
  } catch {
    return null;
  }
}

function asNullableInteger(value: number | null | undefined): number | null {
  if (typeof value !== 'number' || !Number.isInteger(value)) return null;
  return value;
}

function createDefaultCapabilities(): {
  PlayableMediaTypes: string;
  SupportedCommands: string;
  SupportsMediaControl: boolean;
} {
  return {
    PlayableMediaTypes: 'Video,Audio',
    SupportedCommands:
      'Play,Playstate,PlayMediaSource,SetAudioStreamIndex,SetSubtitleStreamIndex,Mute,Unmute,SetVolume,DisplayContent',
    SupportsMediaControl: true,
  };
}

function buildAuthorizationHeader(params: {
  clientName: string;
  deviceName: string;
  clientVersion: string;
  deviceId: string;
  accessToken: string;
}): string {
  return `MediaBrowser Client="${params.clientName}", Device="${params.deviceName}", DeviceId="${params.deviceId}", Version="${params.clientVersion}", Token="${params.accessToken}"`;
}

export function buildJellyfinTimelinePayload(
  state: JellyfinTimelinePlaybackState,
): JellyfinTimelinePayload {
  return {
    ItemId: state.itemId,
    MediaSourceId: state.mediaSourceId,
    PositionTicks: normalizeTicks(state.positionTicks),
    PlaybackStartTimeTicks: normalizeTicks(state.playbackStartTimeTicks),
    IsPaused: state.isPaused === true,
    IsMuted: state.isMuted === true,
    CanSeek: state.canSeek !== false,
    VolumeLevel: clampVolume(state.volumeLevel),
    PlaybackRate:
      typeof state.playbackRate === 'number' && Number.isFinite(state.playbackRate)
        ? state.playbackRate
        : 1,
    PlayMethod: state.playMethod || 'DirectPlay',
    AudioStreamIndex: asNullableInteger(state.audioStreamIndex),
    SubtitleStreamIndex: asNullableInteger(state.subtitleStreamIndex),
    PlaylistItemId: state.playlistItemId,
    EventName: state.eventName || 'timeupdate',
  };
}

export class JellyfinRemoteSessionService {
  private readonly serverUrl: string;
  private readonly accessToken: string;
  private readonly deviceId: string;
  private readonly fetchImpl: typeof fetch;
  private readonly webSocketFactory?: (url: string) => JellyfinRemoteSocket;
  private readonly socketHeadersFactory?: (
    url: string,
    headers: JellyfinRemoteSocketHeaders,
  ) => JellyfinRemoteSocket;
  private readonly setTimer: typeof setTimeout;
  private readonly clearTimer: typeof clearTimeout;
  private readonly onPlay?: (payload: unknown) => void;
  private readonly onPlaystate?: (payload: unknown) => void;
  private readonly onGeneralCommand?: (payload: unknown) => void;
  private readonly capabilities: {
    PlayableMediaTypes: string;
    SupportedCommands: string;
    SupportsMediaControl: boolean;
  };
  private readonly authHeader: string;
  private readonly onConnected?: () => void;
  private readonly onDisconnected?: () => void;

  private readonly reconnectBaseDelayMs: number;
  private readonly reconnectMaxDelayMs: number;
  private socket: JellyfinRemoteSocket | null = null;
  private running = false;
  private connected = false;
  private reconnectAttempt = 0;
  private reconnectTimer: ReturnType<typeof setTimeout> | null = null;

  constructor(options: JellyfinRemoteSessionServiceOptions) {
    this.serverUrl = normalizeServerUrl(options.serverUrl);
    this.accessToken = options.accessToken;
    this.deviceId = options.deviceId;
    this.fetchImpl = options.fetchImpl ?? fetch;
    this.webSocketFactory = options.webSocketFactory;
    this.socketHeadersFactory = options.socketHeadersFactory;
    this.setTimer = options.setTimer ?? setTimeout;
    this.clearTimer = options.clearTimer ?? clearTimeout;
    this.onPlay = options.onPlay;
    this.onPlaystate = options.onPlaystate;
    this.onGeneralCommand = options.onGeneralCommand;
    this.capabilities = {
      ...createDefaultCapabilities(),
      ...(options.capabilities ?? {}),
    };
    const clientName = options.clientName || 'SubMiner';
    const clientVersion = options.clientVersion || '0.1.0';
    const deviceName = options.deviceName || clientName;
    this.authHeader = buildAuthorizationHeader({
      clientName,
      deviceName,
      clientVersion,
      deviceId: this.deviceId,
      accessToken: this.accessToken,
    });
    this.onConnected = options.onConnected;
    this.onDisconnected = options.onDisconnected;
    this.reconnectBaseDelayMs = Math.max(100, options.reconnectBaseDelayMs ?? 500);
    this.reconnectMaxDelayMs = Math.max(
      this.reconnectBaseDelayMs,
      options.reconnectMaxDelayMs ?? 10_000,
    );
  }

  public start(): void {
    if (this.running) return;
    this.running = true;
    this.reconnectAttempt = 0;
    this.connectSocket();
  }

  public stop(): void {
    this.running = false;
    this.connected = false;
    if (this.reconnectTimer) {
      this.clearTimer(this.reconnectTimer);
      this.reconnectTimer = null;
    }
    if (this.socket) {
      this.socket.close();
      this.socket = null;
    }
  }

  public isConnected(): boolean {
    return this.connected;
  }

  public async advertiseNow(): Promise<boolean> {
    await this.postCapabilities();
    return this.isRegisteredOnServer();
  }

  public async reportPlaying(state: JellyfinTimelinePlaybackState): Promise<boolean> {
    return this.postTimeline('/Sessions/Playing', {
      ...buildJellyfinTimelinePayload(state),
      EventName: state.eventName || 'start',
    });
  }

  public async reportProgress(state: JellyfinTimelinePlaybackState): Promise<boolean> {
    return this.postTimeline('/Sessions/Playing/Progress', buildJellyfinTimelinePayload(state));
  }

  public async reportStopped(state: JellyfinTimelinePlaybackState): Promise<boolean> {
    return this.postTimeline('/Sessions/Playing/Stopped', {
      ...buildJellyfinTimelinePayload(state),
      EventName: state.eventName || 'stop',
    });
  }

  private connectSocket(): void {
    if (!this.running) return;
    if (this.reconnectTimer) {
      this.clearTimer(this.reconnectTimer);
      this.reconnectTimer = null;
    }
    const socket = this.createSocket(this.createSocketUrl());
    this.socket = socket;
    let disconnected = false;

    socket.on('open', () => {
      if (this.socket !== socket || !this.running) return;
      this.connected = true;
      this.reconnectAttempt = 0;
      this.onConnected?.();
      void this.postCapabilities();
    });

    socket.on('message', (rawData) => {
      this.handleInboundMessage(rawData);
    });

    const handleDisconnect = () => {
      if (disconnected) return;
      disconnected = true;
      if (this.socket === socket) {
        this.socket = null;
      }
      this.connected = false;
      this.onDisconnected?.();
      if (this.running) {
        this.scheduleReconnect();
      }
    };

    socket.on('close', handleDisconnect);
    socket.on('error', handleDisconnect);
  }

  private scheduleReconnect(): void {
    const delay = Math.min(
      this.reconnectMaxDelayMs,
      this.reconnectBaseDelayMs * 2 ** this.reconnectAttempt,
    );
    this.reconnectAttempt += 1;
    if (this.reconnectTimer) {
      this.clearTimer(this.reconnectTimer);
    }
    this.reconnectTimer = this.setTimer(() => {
      this.reconnectTimer = null;
      this.connectSocket();
    }, delay);
  }

  private createSocketUrl(): string {
    const baseUrl = new URL(`${this.serverUrl}/`);
    const socketUrl = new URL('/socket', baseUrl);
    socketUrl.protocol = baseUrl.protocol === 'https:' ? 'wss:' : 'ws:';
    socketUrl.searchParams.set('api_key', this.accessToken);
    socketUrl.searchParams.set('deviceId', this.deviceId);
    return socketUrl.toString();
  }

  private createSocket(url: string): JellyfinRemoteSocket {
    const headers: JellyfinRemoteSocketHeaders = {
      Authorization: this.authHeader,
      'X-Emby-Authorization': this.authHeader,
      'X-Emby-Token': this.accessToken,
    };
    if (this.socketHeadersFactory) {
      return this.socketHeadersFactory(url, headers);
    }
    if (this.webSocketFactory) {
      return this.webSocketFactory(url);
    }
    return new WebSocket(url, { headers }) as unknown as JellyfinRemoteSocket;
  }

  private async postCapabilities(): Promise<void> {
    const payload = this.capabilities;
    const fullEndpointOk = await this.postJson('/Sessions/Capabilities/Full', payload);
    if (fullEndpointOk) return;
    await this.postJson('/Sessions/Capabilities', payload);
  }

  private async isRegisteredOnServer(): Promise<boolean> {
    try {
      const response = await this.fetchImpl(`${this.serverUrl}/Sessions`, {
        method: 'GET',
        headers: {
          Authorization: this.authHeader,
          'X-Emby-Authorization': this.authHeader,
          'X-Emby-Token': this.accessToken,
        },
      });
      if (!response.ok) return false;
      const sessions = (await response.json()) as Array<Record<string, unknown>>;
      return sessions.some((session) => String(session.DeviceId || '') === this.deviceId);
    } catch {
      return false;
    }
  }

  private async postTimeline(path: string, payload: JellyfinTimelinePayload): Promise<boolean> {
    return this.postJson(path, payload);
  }

  private async postJson(path: string, payload: unknown): Promise<boolean> {
    try {
      const response = await this.fetchImpl(`${this.serverUrl}${path}`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: this.authHeader,
          'X-Emby-Authorization': this.authHeader,
          'X-Emby-Token': this.accessToken,
        },
        body: JSON.stringify(payload),
      });
      return response.ok;
    } catch {
      return false;
    }
  }

  private handleInboundMessage(rawData: unknown): void {
    const message = parseInboundMessage(rawData);
    if (!message) return;
    const messageType = message.MessageType;
    const payload = parseMessageData(message.Data);
    if (messageType === 'Play') {
      this.onPlay?.(payload);
      return;
    }
    if (messageType === 'Playstate') {
      this.onPlaystate?.(payload);
      return;
    }
    if (messageType === 'GeneralCommand') {
      this.onGeneralCommand?.(payload);
    }
  }
}
