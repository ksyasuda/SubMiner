import http, { IncomingMessage, ServerResponse } from 'node:http';
import axios, { AxiosInstance } from 'axios';

interface StartProxyOptions {
  host: string;
  port: number;
  upstreamUrl: string;
}

interface AnkiConnectEnvelope {
  result: unknown;
  error: unknown;
}

export interface AnkiConnectProxyServerDeps {
  shouldAutoUpdateNewCards: () => boolean;
  processNewCard: (noteId: number) => Promise<void>;
  logInfo: (message: string, ...args: unknown[]) => void;
  logWarn: (message: string, ...args: unknown[]) => void;
  logError: (message: string, ...args: unknown[]) => void;
}

export class AnkiConnectProxyServer {
  private server: http.Server | null = null;
  private client: AxiosInstance;
  private pendingNoteIds: number[] = [];
  private pendingNoteIdSet = new Set<number>();
  private inFlightNoteIds = new Set<number>();
  private processingQueue = false;

  constructor(private readonly deps: AnkiConnectProxyServerDeps) {
    this.client = axios.create({
      timeout: 15000,
      validateStatus: () => true,
      responseType: 'arraybuffer',
    });
  }

  get isRunning(): boolean {
    return this.server !== null;
  }

  start(options: StartProxyOptions): void {
    this.stop();

    if (this.isSelfReferentialProxy(options)) {
      this.deps.logError(
        '[anki-proxy] Proxy upstream points to proxy host/port; refusing to start to avoid loop.',
      );
      return;
    }

    this.server = http.createServer((req, res) => {
      void this.handleRequest(req, res, options.upstreamUrl);
    });

    this.server.on('error', (error) => {
      this.deps.logError('[anki-proxy] Server error:', (error as Error).message);
    });

    this.server.listen(options.port, options.host, () => {
      this.deps.logInfo(
        `[anki-proxy] Listening on http://${options.host}:${options.port} -> ${options.upstreamUrl}`,
      );
    });
  }

  stop(): void {
    if (this.server) {
      this.server.close();
      this.server = null;
      this.deps.logInfo('[anki-proxy] Stopped');
    }
    this.pendingNoteIds = [];
    this.pendingNoteIdSet.clear();
    this.inFlightNoteIds.clear();
    this.processingQueue = false;
  }

  private isSelfReferentialProxy(options: StartProxyOptions): boolean {
    try {
      const upstream = new URL(options.upstreamUrl);
      const normalizedUpstreamHost = upstream.hostname.toLowerCase();
      const normalizedBindHost = options.host.toLowerCase();
      const upstreamPort =
        upstream.port.length > 0
          ? Number(upstream.port)
          : upstream.protocol === 'https:'
            ? 443
            : 80;
      const hostMatches =
        normalizedUpstreamHost === normalizedBindHost ||
        (normalizedUpstreamHost === 'localhost' && normalizedBindHost === '127.0.0.1') ||
        (normalizedUpstreamHost === '127.0.0.1' && normalizedBindHost === 'localhost');
      return hostMatches && upstreamPort === options.port;
    } catch {
      return false;
    }
  }

  private async handleRequest(
    req: IncomingMessage,
    res: ServerResponse<IncomingMessage>,
    upstreamUrl: string,
  ): Promise<void> {
    this.setCorsHeaders(res);

    if (req.method === 'OPTIONS') {
      res.statusCode = 204;
      res.end();
      return;
    }

    if (!req.method || (req.method !== 'GET' && req.method !== 'POST')) {
      res.statusCode = 405;
      res.end('Method Not Allowed');
      return;
    }

    let rawBody: Buffer = Buffer.alloc(0);
    if (req.method === 'POST') {
      rawBody = await this.readRequestBody(req);
    }

    let requestJson: Record<string, unknown> | null = null;
    if (req.method === 'POST' && rawBody.length > 0) {
      requestJson = this.tryParseJson(rawBody);
    }

    try {
      const targetUrl = new URL(req.url || '/', upstreamUrl).toString();
      const contentType =
        typeof req.headers['content-type'] === 'string'
          ? req.headers['content-type']
          : 'application/json';
      const upstreamResponse = await this.client.request<ArrayBuffer>({
        url: targetUrl,
        method: req.method,
        data: req.method === 'POST' ? rawBody : undefined,
        headers: {
          'content-type': contentType,
        },
      });

      const responseBody: Buffer = Buffer.isBuffer(upstreamResponse.data)
        ? upstreamResponse.data
        : Buffer.from(new Uint8Array(upstreamResponse.data));
      this.copyUpstreamHeaders(res, upstreamResponse.headers as Record<string, unknown>);
      res.statusCode = upstreamResponse.status;
      res.end(responseBody);

      if (req.method === 'POST') {
        this.maybeEnqueueFromRequest(requestJson, responseBody);
      }
    } catch (error) {
      this.deps.logWarn('[anki-proxy] Failed to forward request:', (error as Error).message);
      res.statusCode = 502;
      res.end('Bad Gateway');
    }
  }

  private maybeEnqueueFromRequest(
    requestJson: Record<string, unknown> | null,
    responseBody: Buffer,
  ): void {
    if (!requestJson || !this.deps.shouldAutoUpdateNewCards()) {
      return;
    }

    const action =
      typeof requestJson.action === 'string' ? requestJson.action : String(requestJson.action ?? '');
    if (action !== 'addNote' && action !== 'addNotes') {
      return;
    }

    const responseJson = this.tryParseJson(responseBody) as AnkiConnectEnvelope | null;
    if (!responseJson || responseJson.error !== null) {
      return;
    }

    const noteIds =
      action === 'addNote'
        ? this.collectSingleResultId(responseJson.result)
        : this.collectBatchResultIds(responseJson.result);
    if (noteIds.length === 0) {
      return;
    }

    this.enqueueNotes(noteIds);
  }

  private collectSingleResultId(value: unknown): number[] {
    if (typeof value === 'number' && Number.isInteger(value) && value > 0) {
      return [value];
    }
    return [];
  }

  private collectBatchResultIds(value: unknown): number[] {
    if (!Array.isArray(value)) {
      return [];
    }
    return value.filter((entry): entry is number => {
      return typeof entry === 'number' && Number.isInteger(entry) && entry > 0;
    });
  }

  private enqueueNotes(noteIds: number[]): void {
    let enqueuedCount = 0;
    for (const noteId of noteIds) {
      if (this.pendingNoteIdSet.has(noteId) || this.inFlightNoteIds.has(noteId)) {
        continue;
      }
      this.pendingNoteIds.push(noteId);
      this.pendingNoteIdSet.add(noteId);
      enqueuedCount += 1;
    }

    if (enqueuedCount === 0) {
      return;
    }

    this.deps.logInfo(`[anki-proxy] Enqueued ${enqueuedCount} note(s) for enrichment`);
    this.processQueue();
  }

  private processQueue(): void {
    if (this.processingQueue) {
      return;
    }
    this.processingQueue = true;

    void (async () => {
      try {
        while (this.pendingNoteIds.length > 0) {
          const noteId = this.pendingNoteIds.shift();
          if (noteId === undefined) {
            continue;
          }
          this.pendingNoteIdSet.delete(noteId);

          if (!this.deps.shouldAutoUpdateNewCards()) {
            continue;
          }

          this.inFlightNoteIds.add(noteId);
          try {
            await this.deps.processNewCard(noteId);
          } catch (error) {
            this.deps.logWarn(
              `[anki-proxy] Failed to auto-enrich note ${noteId}:`,
              (error as Error).message,
            );
          } finally {
            this.inFlightNoteIds.delete(noteId);
          }
        }
      } finally {
        this.processingQueue = false;
        if (this.pendingNoteIds.length > 0) {
          this.processQueue();
        }
      }
    })();
  }

  private async readRequestBody(req: IncomingMessage): Promise<Buffer> {
    const chunks: Buffer[] = [];
    for await (const chunk of req) {
      chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    }
    return Buffer.concat(chunks);
  }

  private tryParseJson(rawBody: Buffer): Record<string, unknown> | null {
    if (rawBody.length === 0) {
      return null;
    }
    try {
      const parsed = JSON.parse(rawBody.toString('utf8'));
      return parsed && typeof parsed === 'object' ? (parsed as Record<string, unknown>) : null;
    } catch {
      return null;
    }
  }

  private setCorsHeaders(res: ServerResponse<IncomingMessage>): void {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
  }

  private copyUpstreamHeaders(
    res: ServerResponse<IncomingMessage>,
    headers: Record<string, unknown>,
  ): void {
    for (const [key, value] of Object.entries(headers)) {
      if (value === undefined) {
        continue;
      }
      if (key.toLowerCase() === 'content-length') {
        continue;
      }
      if (Array.isArray(value)) {
        res.setHeader(
          key,
          value.map((entry) => String(entry)),
        );
      } else {
        res.setHeader(key, String(value));
      }
    }
  }
}
