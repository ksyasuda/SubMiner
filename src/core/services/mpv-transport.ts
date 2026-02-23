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

export interface MpvSocketTransportOptions {
  socketPath: string;
  onConnect: () => void;
  onData: (data: Buffer) => void;
  onError: (error: Error) => void;
  onClose: () => void;
  socketFactory?: () => net.Socket;
}

export class MpvSocketTransport {
  private socketPath: string;
  private readonly callbacks: MpvSocketTransportEvents;
  private readonly socketFactory: () => net.Socket;
  private socketRef: net.Socket | null = null;
  public socket: net.Socket | null = null;
  public connected = false;
  public connecting = false;

  constructor(options: MpvSocketTransportOptions) {
    this.socketPath = options.socketPath;
    this.socketFactory = options.socketFactory ?? (() => new net.Socket());
    this.callbacks = {
      onConnect: options.onConnect,
      onData: options.onData,
      onError: options.onError,
      onClose: options.onClose,
    };
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
    this.socketRef = this.socketFactory();
    this.socket = this.socketRef;

    this.socketRef.on('connect', () => {
      this.connected = true;
      this.connecting = false;
      this.callbacks.onConnect();
    });

    this.socketRef.on('data', (data: Buffer) => {
      this.callbacks.onData(data);
    });

    this.socketRef.on('error', (error: Error) => {
      this.connected = false;
      this.connecting = false;
      this.callbacks.onError(error);
    });

    this.socketRef.on('close', () => {
      this.connected = false;
      this.connecting = false;
      this.callbacks.onClose();
    });

    this.socketRef.connect(this.socketPath);
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
    if (this.socketRef) {
      this.socketRef.destroy();
    }
    this.socketRef = null;
    this.socket = null;
    this.connected = false;
    this.connecting = false;
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
