export type HachidoriHostStatus =
  | { kind: 'local' }
  | { kind: 'connected'; address: string; name: string; dictionaryCount: number }
  | { kind: 'disconnected'; address: string; message: string }
  | { kind: 'unavailable'; message: string };

export type HachidoriSharingRequest =
  | { type: 'hd_sharing_status' }
  | { type: 'hd_sharing_client_link'; address: string }
  | { type: 'hd_sharing_client_unlink' };

export function parseHachidoriManagementUrl(value: unknown): string {
  if (typeof value !== 'string') throw new Error('Expected an HTTP(S) origin or an empty string.');
  if (!value.trim()) return '';
  const url = new URL(value.trim());
  if (
    !['http:', 'https:'].includes(url.protocol) ||
    url.username ||
    url.password ||
    url.pathname !== '/' ||
    url.search ||
    url.hash
  ) {
    throw new Error('Expected an HTTP(S) origin without credentials, a path, query, or fragment.');
  }
  return url.origin;
}

function object(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

export function parseHachidoriHostStatus(reply: unknown): HachidoriHostStatus {
  if (!object(reply) || reply.ok !== true || !object(reply.sharing)) {
    throw new Error(
      object(reply) && typeof reply.error === 'string'
        ? reply.error
        : 'Hachidori returned an invalid sharing status.',
    );
  }
  const client = reply.sharing.client;
  if (!object(client) || typeof client.linked !== 'boolean') {
    throw new Error('Hachidori returned an invalid host connection.');
  }
  if (!client.linked) return { kind: 'local' };
  if (typeof client.address !== 'string') throw new Error('Missing Hachidori host address.');
  if (client.connected !== true) {
    return {
      kind: 'disconnected',
      address: client.address,
      message:
        typeof client.error === 'string'
          ? client.error
          : 'The dictionary host is not reachable. Start it, then refresh status.',
    };
  }
  const host = client.host;
  if (
    !object(host) ||
    typeof host.dictionaryCount !== 'number' ||
    !Number.isSafeInteger(host.dictionaryCount) ||
    host.dictionaryCount < 0
  ) {
    throw new Error('Hachidori returned an invalid dictionary count.');
  }
  return {
    kind: 'connected',
    address: client.address,
    name: typeof host.name === 'string' ? host.name : 'Hachidori',
    dictionaryCount: host.dictionaryCount,
  };
}

export function buildHachidoriSharingScript(request: HachidoriSharingRequest): string {
  return `(async () => {
    const request = ${JSON.stringify(request)};
    const requestWithTimeout = async fields => {
      let timer;
      try {
        return await Promise.race([
          chrome.runtime.sendMessage({ requestId: crypto.randomUUID(), ...fields }),
          new Promise((_, reject) => { timer = setTimeout(() => reject(new Error('The dictionary host did not respond. Check the host and refresh status.')), 5000); }),
        ]);
      } finally { clearTimeout(timer); }
    };
    const send = fields => requestWithTimeout({ target: 'hachidori-sharing', ...fields });
    const read = (target, type) => requestWithTimeout({ target, type });
    let reply = await send(request);
    if (!reply?.ok) throw new Error(reply?.error || 'Hachidori could not update the host connection.');
    if (reply.sharing?.client?.linked && (reply.sharing.client.connected || request.type === 'hd_sharing_client_link')) {
      try {
        const engine = await read('hoshidicts-offscreen', 'hd_status');
        reply = await send({type: 'hd_sharing_status'});
        if (engine?.ok && engine.ready && !engine.loading && reply.sharing?.client?.connected) {
          const inventory = await read('hoshidicts-worker', 'hd_state_read');
          if (!inventory?.ok || !Array.isArray(inventory.state?.dictionaries)) throw new Error('Could not read the dictionary host library.');
          reply.sharing.client.host.dictionaryCount = inventory.state.dictionaries.length;
        } else if (reply.sharing?.client?.linked) {
          reply.sharing.client.connected = false;
          reply.sharing.client.error = engine?.loading ? 'The dictionary host is loading. Refresh status when it is ready.' : (engine?.error || 'The dictionary host is not ready.');
        }
      } catch (error) {
        reply.sharing.client.connected = false;
        reply.sharing.client.error = error.message;
      }
    }
    return reply;
  })()`;
}
