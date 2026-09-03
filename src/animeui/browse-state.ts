import type { AnimeBrowserEntry } from '../types/anime-browser';

export interface BrowseState {
  requestId: number;
  query: string;
  page: number;
  loading: boolean;
  hasNextPage: boolean;
}

export interface BrowseRequest {
  id: number;
  query: string;
  page: number;
  append: boolean;
}

export interface StartedBrowse {
  state: BrowseState;
  request: BrowseRequest;
}

export function createBrowseState(): BrowseState {
  return { requestId: 0, query: '', page: 0, loading: false, hasNextPage: false };
}

export function beginBrowse(state: BrowseState, query: string): StartedBrowse {
  const request: BrowseRequest = {
    id: state.requestId + 1,
    query,
    page: 1,
    append: false,
  };
  return {
    state: {
      requestId: request.id,
      query,
      page: request.page,
      loading: true,
      hasNextPage: false,
    },
    request,
  };
}

export function beginNextPage(state: BrowseState): StartedBrowse | null {
  if (state.loading || !state.hasNextPage) return null;
  const request: BrowseRequest = {
    id: state.requestId + 1,
    query: state.query,
    page: state.page + 1,
    append: true,
  };
  return {
    state: {
      ...state,
      requestId: request.id,
      page: request.page,
      loading: true,
      hasNextPage: false,
    },
    request,
  };
}

export function soleBrowseRequest(
  inFlight: ReadonlyMap<number, BrowseRequest>,
): BrowseRequest | null {
  if (inFlight.size !== 1) return null;
  return inFlight.values().next().value ?? null;
}

export function finishBrowse(
  state: BrowseState,
  requestId: number,
  hasNextPage: boolean,
): BrowseState {
  if (requestId !== state.requestId) return state;
  return { ...state, loading: false, hasNextPage };
}

export function failBrowse(state: BrowseState, request: BrowseRequest): BrowseState {
  if (request.id !== state.requestId) return state;
  return {
    ...state,
    page: request.append ? request.page - 1 : request.page,
    loading: false,
    hasNextPage: request.append,
  };
}

export class LatestRequest {
  private current = 0;

  begin(): number {
    return ++this.current;
  }

  cancel(): void {
    this.current += 1;
  }

  isCurrent(request: number): boolean {
    return request === this.current;
  }
}

export function safeUploadDate(uploadedAt: number): string | null {
  const date = new Date(uploadedAt);
  return Number.isFinite(date.getTime()) ? date.toISOString().slice(0, 10) : null;
}

export type Captured<T> = { ok: true; value: T } | { ok: false; error: unknown };

export async function capture<T>(operation: () => Promise<T>): Promise<Captured<T>> {
  try {
    return { ok: true, value: await operation() };
  } catch (error) {
    return { ok: false, error };
  }
}

export function takeUnseenEntries(
  entries: AnimeBrowserEntry[],
  seen: Set<string>,
): AnimeBrowserEntry[] {
  return entries.filter((entry) => {
    const key = `${entry.sourceId}\0${entry.url}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}
