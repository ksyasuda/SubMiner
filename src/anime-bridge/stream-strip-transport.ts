import http from 'node:http';

const UPSTREAM_TIMEOUT_MS = 15_000;

interface ClientRequestLifecycle {
  activeUpstreamRequest: http.ClientRequest | null;
  closed: boolean;
}

const DROPPED_REQUEST_HEADERS = new Set([
  'accept-encoding',
  'connection',
  'keep-alive',
  'transfer-encoding',
  'content-length',
]);

export function forwardableRequestHeaders(
  headers: http.IncomingHttpHeaders,
): http.OutgoingHttpHeaders {
  const result: http.OutgoingHttpHeaders = {};
  for (const [name, value] of Object.entries(headers)) {
    if (value === undefined || DROPPED_REQUEST_HEADERS.has(name.toLowerCase())) continue;
    result[name] = value;
  }
  result['accept-encoding'] = 'identity';
  return result;
}

export function requestUpstream(options: {
  req: http.IncomingMessage;
  res: http.ServerResponse;
  upstreamUrl: URL;
  requestHeaders: http.OutgoingHttpHeaders;
  retryDelayMs: number;
  log: (message: string) => void;
  handleResponse: (upstream: http.IncomingMessage) => void;
}): void {
  const lifecycle: ClientRequestLifecycle = { activeUpstreamRequest: null, closed: false };
  const destroyUpstreamOnClientClose = (): void => {
    lifecycle.closed = true;
    lifecycle.activeUpstreamRequest?.destroy();
  };
  options.req.once('aborted', destroyUpstreamOnClientClose);
  options.res.once('close', destroyUpstreamOnClientClose);
  options.res.once('finish', () => {
    options.req.off('aborted', destroyUpstreamOnClientClose);
    options.res.off('close', destroyUpstreamOnClientClose);
    lifecycle.activeUpstreamRequest = null;
  });

  requestAttempt(options, lifecycle, 0);
}

function requestAttempt(
  options: Parameters<typeof requestUpstream>[0],
  lifecycle: ClientRequestLifecycle,
  attempt: number,
): void {
  const { req, res, upstreamUrl, requestHeaders, retryDelayMs, log } = options;
  if (lifecycle.closed || res.destroyed) return;
  const mayRetry = req.method === 'GET' && attempt === 0;
  let handedOff = false;
  const retry = (): void => {
    setTimeout(() => requestAttempt(options, lifecycle, attempt + 1), retryDelayMs);
  };

  const upstreamRequest = http.request(
    upstreamUrl,
    { method: req.method, headers: requestHeaders, timeout: UPSTREAM_TIMEOUT_MS },
    (upstream) => {
      const clearActiveRequest = (): void => {
        if (lifecycle.activeUpstreamRequest === upstreamRequest) {
          lifecycle.activeUpstreamRequest = null;
        }
      };
      upstream.once('end', clearActiveRequest);
      upstream.once('close', clearActiveRequest);
      const status = upstream.statusCode ?? 502;
      if (status === 404 || status >= 500) {
        if (mayRetry) {
          log(`[stream-proxy] upstream ${status} for ${upstreamUrl.pathname}; retrying once`);
          upstream.resume();
          retry();
          return;
        }
        log(`[stream-proxy] upstream ${status} for ${upstreamUrl.pathname}`);
      }
      handedOff = true;
      options.handleResponse(upstream);
    },
  );
  lifecycle.activeUpstreamRequest = upstreamRequest;
  upstreamRequest.on('timeout', () => {
    upstreamRequest.destroy(new Error(`upstream silent for ${UPSTREAM_TIMEOUT_MS}ms`));
  });
  upstreamRequest.on('error', (error) => {
    if (handedOff) {
      log(`[stream-proxy] upstream failed mid-response: ${String(error)}`);
      res.destroy();
      return;
    }
    if (mayRetry) {
      log(`[stream-proxy] upstream request failed: ${String(error)}; retrying once`);
      retry();
      return;
    }
    log(`[stream-proxy] upstream request failed: ${String(error)}`);
    if (!res.headersSent) res.writeHead(502);
    res.end();
  });
  upstreamRequest.end();
}
