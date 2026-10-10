import WebSocket from 'ws';
import { JellyfinPlaybackReporter } from './jellyfin-playback-reporter';

export interface JellyfinRemoteSessionMessage {
  MessageType?: string;
  Data?: unknown;
}

interface JellyfinRemoteSocket {
  on(event: 'open', listener: () => void): this;
  on(event: 'close', listener: () => void): this;
  on(event: 'error', listener: (error: Error) => void): this;
  on(event: 'message', listener: (data: unknown) => void): this;
  send(data: string): void;
  terminate?(): void;
  close(): void;
}

// Jellyfin advertises its keep-alive timeout in the ForceKeepAlive message (60s by default),
// drops sockets that stay silent past it, and since 12.0 also detaches the session's remote
// controller when that happens. The drop never reaches the client as a close frame, so the
// client has to keep sending KeepAlive and treat missing replies as a dead connection.
const DEFAULT_KEEP_ALIVE_TIMEOUT_MS = 60_000;
const KEEP_ALIVE_LOST_FACTOR = 1.5;

function unrefTimer(timer: ReturnType<typeof setTimeout>): void {
  (timer as unknown as { unref?: () => void }).unref?.();
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
  logWarn?: (message: string, details?: unknown) => void;
  keepAliveTimeoutMs?: number;
  getNow?: () => number;
}

function normalizeServerUrl(serverUrl: string): string {
  return serverUrl.trim().replace(/\/+$/, '');
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
  // Shares this device's identity for capability posts; playback reports use their own reporter.
  private readonly http: JellyfinPlaybackReporter;
  private readonly onConnected?: () => void;
  private readonly onDisconnected?: () => void;
  private readonly logWarn?: (message: string, details?: unknown) => void;
  private readonly now: () => number;
  private keepAliveTimeoutMs: number;
  private keepAliveTimer: ReturnType<typeof setTimeout> | null = null;
  private lastInboundAtMs = 0;

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
    this.logWarn = options.logWarn;
    this.http = new JellyfinPlaybackReporter({
      serverUrl: this.serverUrl,
      accessToken: this.accessToken,
      deviceId: this.deviceId,
      clientName: options.clientName,
      clientVersion: options.clientVersion,
      deviceName: options.deviceName,
      fetchImpl: this.fetchImpl,
      logWarn: this.logWarn,
    });
    this.authHeader = this.http.authHeader;
    this.onConnected = options.onConnected;
    this.onDisconnected = options.onDisconnected;
    this.now = options.getNow ?? Date.now;
    this.keepAliveTimeoutMs = Math.max(
      1000,
      options.keepAliveTimeoutMs ?? DEFAULT_KEEP_ALIVE_TIMEOUT_MS,
    );
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
    this.stopKeepAlive();
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
      this.lastInboundAtMs = this.now();
      this.startKeepAlive(socket, this.keepAliveTimeoutMs);
      this.onConnected?.();
      void this.postCapabilities();
    });

    socket.on('message', (rawData) => {
      if (this.socket !== socket || !this.running) return;
      this.lastInboundAtMs = this.now();
      this.handleInboundMessage(socket, rawData);
    });

    const handleDisconnect = () => {
      if (disconnected) return;
      disconnected = true;
      if (this.socket === socket) {
        this.socket = null;
        this.stopKeepAlive();
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

  private startKeepAlive(socket: JellyfinRemoteSocket, timeoutMs: number): void {
    this.stopKeepAlive();
    this.keepAliveTimeoutMs = timeoutMs;
    this.sendKeepAlive(socket);
    this.scheduleKeepAliveTick(socket);
  }

  private scheduleKeepAliveTick(socket: JellyfinRemoteSocket): void {
    const intervalMs = Math.max(1000, Math.floor(this.keepAliveTimeoutMs / 2));
    const timer = this.setTimer(() => {
      this.keepAliveTimer = null;
      if (this.socket !== socket || !this.running) return;
      const silentForMs = this.now() - this.lastInboundAtMs;
      if (silentForMs >= this.keepAliveTimeoutMs * KEEP_ALIVE_LOST_FACTOR) {
        this.logWarn?.('Jellyfin remote websocket stopped answering keep-alives; reconnecting.');
        // Dropping the socket raises 'close', which schedules the reconnect.
        if (socket.terminate) {
          socket.terminate();
        } else {
          socket.close();
        }
        return;
      }
      this.sendKeepAlive(socket);
      this.scheduleKeepAliveTick(socket);
    }, intervalMs);
    unrefTimer(timer);
    this.keepAliveTimer = timer;
  }

  private stopKeepAlive(): void {
    if (this.keepAliveTimer) {
      this.clearTimer(this.keepAliveTimer);
      this.keepAliveTimer = null;
    }
  }

  private sendKeepAlive(socket: JellyfinRemoteSocket): void {
    try {
      socket.send(JSON.stringify({ MessageType: 'KeepAlive' }));
    } catch (error) {
      this.logWarn?.('Failed to send Jellyfin remote keep-alive.', error);
    }
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
    socketUrl.searchParams.set('ApiKey', this.accessToken);
    socketUrl.searchParams.set('deviceId', this.deviceId);
    return socketUrl.toString();
  }

  private createSocket(url: string): JellyfinRemoteSocket {
    const headers: JellyfinRemoteSocketHeaders = {
      Authorization: this.authHeader,
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
    const fullEndpointOk = await this.http.postJson('/Sessions/Capabilities/Full', payload);
    if (fullEndpointOk) return;
    await this.http.postJson('/Sessions/Capabilities', payload);
  }

  private async isRegisteredOnServer(): Promise<boolean> {
    try {
      const response = await this.fetchImpl(`${this.serverUrl}/Sessions`, {
        method: 'GET',
        headers: {
          Authorization: this.authHeader,
        },
      });
      if (!response.ok) return false;
      const sessions = (await response.json()) as Array<Record<string, unknown>>;
      return sessions.some((session) => String(session.DeviceId || '') === this.deviceId);
    } catch {
      return false;
    }
  }

  private handleInboundMessage(socket: JellyfinRemoteSocket, rawData: unknown): void {
    const message = parseInboundMessage(rawData);
    if (!message) return;
    const messageType = message.MessageType;
    if (messageType === 'ForceKeepAlive') {
      const seconds = Number(message.Data);
      const timeoutMs =
        Number.isFinite(seconds) && seconds > 0 ? seconds * 1000 : this.keepAliveTimeoutMs;
      this.startKeepAlive(socket, timeoutMs);
      return;
    }
    if (messageType === 'KeepAlive') return;
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
