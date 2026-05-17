import type { FetchLike, FetchResponseLike } from './release-assets';

export interface ElectronNetFetchLike {
  fetch: (url: string, init?: Record<string, unknown>) => Promise<FetchResponseLike>;
}

export function createElectronNetFetch(net: ElectronNetFetchLike): FetchLike {
  return (url, init) => net.fetch(url, init);
}
