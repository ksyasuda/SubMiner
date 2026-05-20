import type { FetchLike, FetchResponseLike } from './release-assets';

export interface ElectronNetFetchLike {
  fetch: (url: string, init?: Record<string, unknown>) => Promise<FetchResponseLike>;
}

export type GlobalFetchLike = (url: string, init?: RequestInit) => Promise<FetchResponseLike>;

export function createElectronNetFetch(net: ElectronNetFetchLike): FetchLike {
  return (url, init) => net.fetch(url, init);
}

function getGlobalFetch(): GlobalFetchLike {
  if (typeof globalThis.fetch !== 'function') {
    throw new Error('Global fetch is not available for updater requests.');
  }
  return globalThis.fetch.bind(globalThis) as GlobalFetchLike;
}

export function createGlobalFetch(fetchImpl?: GlobalFetchLike): FetchLike {
  return (url, init) => (fetchImpl ?? getGlobalFetch())(url, init as RequestInit);
}
