import * as net from 'net';

export function getMpvReconnectDelay(attempt: number, hasConnectedOnce: boolean): number {
  if (hasConnectedOnce) {
    if (attempt < 2) {
      return 1000;
    }
    if (attempt < 4) {
      return 2000;
    }
    if (attempt < 7) {
      return 5000;
    }
    return 10000;
  }

  if (attempt < 2) {
    return 200;
  }
  if (attempt < 4) {
    return 500;
  }
  if (attempt < 6) {
    return 1000;
  }
  return 2000;
}

export interface MpvReconnectSchedulerDeps {
  attempt: number;
  hasConnectedOnce: boolean;
  getReconnectTimer: () => ReturnType<typeof setTimeout> | null;
  setReconnectTimer: (timer: ReturnType<typeof setTimeout> | null) => void;
  onReconnectAttempt: (attempt: number, delay: number) => void;
  connect: () => void;
}

export function scheduleMpvReconnect(deps: MpvReconnectSchedulerDeps): number {
  const reconnectTimer = deps.getReconnectTimer();
  if (reconnectTimer) {
    clearTimeout(reconnectTimer);
  }
  const delay = getMpvReconnectDelay(deps.attempt, deps.hasConnectedOnce);
  deps.setReconnectTimer(
    setTimeout(() => {
      deps.onReconnectAttempt(deps.attempt + 1, delay);
      deps.connect();
    }, delay),
  );
  return deps.attempt + 1;
}

export interface MpvSocketMessagePayload {
  command: unknown[];
  request_id?: number;
}

interface MpvSocketTransportEvents {
  onConnect: () => void;
  onData: (data: Buffer) => void;
  onError: (error: Error) => void;
  onClose: () => void;
}

export const MPV_CONNECT_TIMEOUT_MS = 5000;

export interface MpvSocketTransportOptions {
  socketPath: string;
  onConnect: () => void;
  onData: (data: Buffer) => void;
  onError: (error: Error) => void;
  onClose: () => void;
  socketFactory?: () => net.Socket;
  connectTimeoutMs?: number;
}

export class MpvSocketTransport {
  private socketPath: string;
  private readonly callbacks: MpvSocketTransportEvents;
  private readonly socketFactory: () => net.Socket;
  private readonly connectTimeoutMs: number;
  private socketRef: net.Socket | null = null;
  private connectTimer: ReturnType<typeof setTimeout> | null = null;
  public socket: net.Socket | null = null;
  public connected = false;
  public connecting = false;

  constructor(options: MpvSocketTransportOptions) {
    this.socketPath = options.socketPath;
    this.socketFactory = options.socketFactory ?? (() => new net.Socket());
    this.connectTimeoutMs = options.connectTimeoutMs ?? MPV_CONNECT_TIMEOUT_MS;
    this.callbacks = {
      onConnect: options.onConnect,
      onData: options.onData,
      onError: options.onError,
      onClose: options.onClose,
    };
  }

  private clearConnectTimeout(): void {
    if (this.connectTimer) {
      clearTimeout(this.connectTimer);
      this.connectTimer = null;
    }
  }

  // A named-pipe/socket dial that neither connects nor errors would otherwise
  // latch `connecting` forever and silently block every future connect().
  private armConnectTimeout(socket: net.Socket): void {
    this.clearConnectTimeout();
    this.connectTimer = setTimeout(() => {
      this.connectTimer = null;
      if (this.socketRef !== socket || this.connected) return;
      this.connecting = false;
      this.callbacks.onError(
        new Error(`MPV IPC connect timed out after ${this.connectTimeoutMs}ms: ${this.socketPath}`),
      );
      // Destroying the socket emits 'close', which drives the normal
      // disconnect path (including reconnect scheduling) upstream.
      socket.destroy();
    }, this.connectTimeoutMs);
    this.connectTimer.unref?.();
  }

  setSocketPath(socketPath: string): void {
    this.socketPath = socketPath;
  }

  connect(): void {
    if (this.connected || this.connecting) {
      return;
    }

    if (this.socketRef) {
      this.socketRef.destroy();
    }

    this.connecting = true;
    const socket = this.socketFactory();
    this.socketRef = socket;
    this.socket = socket;

    socket.on('connect', () => {
      if (this.socketRef !== socket) return;
      this.clearConnectTimeout();
      this.connected = true;
      this.connecting = false;
      this.callbacks.onConnect();
    });

    socket.on('data', (data: Buffer) => {
      if (this.socketRef !== socket) return;
      this.callbacks.onData(data);
    });

    socket.on('error', (error: Error) => {
      if (this.socketRef !== socket) return;
      this.clearConnectTimeout();
      this.connected = false;
      this.connecting = false;
      this.callbacks.onError(error);
    });

    socket.on('close', () => {
      if (this.socketRef !== socket) return;
      this.clearConnectTimeout();
      this.connected = false;
      this.connecting = false;
      this.callbacks.onClose();
    });

    socket.connect(this.socketPath);
    this.armConnectTimeout(socket);
  }

  send(payload: MpvSocketMessagePayload): boolean {
    if (!this.connected || !this.socketRef) {
      return false;
    }

    const message = JSON.stringify(payload) + '\n';
    this.socketRef.write(message);
    return true;
  }

  shutdown(): void {
    this.clearConnectTimeout();
    const socket = this.socketRef;
    this.socketRef = null;
    this.socket = null;
    this.connected = false;
    this.connecting = false;
    if (socket) {
      socket.destroy();
    }
  }

  getSocket(): net.Socket | null {
    return this.socketRef;
  }

  get isConnected(): boolean {
    return this.connected;
  }

  get isConnecting(): boolean {
    return this.connecting;
  }
}
