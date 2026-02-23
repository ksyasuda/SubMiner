import type { AnilistRetryQueueState, AnilistSecretResolutionState } from '../state';

type AnilistQueueSnapshot = Pick<AnilistRetryQueueState, 'pending' | 'ready' | 'deadLetter'>;

type AnilistStatusSnapshot = {
  tokenStatus: AnilistSecretResolutionState['status'];
  tokenSource: AnilistSecretResolutionState['source'];
  tokenMessage: string | null;
  tokenResolvedAt: number | null;
  tokenErrorAt: number | null;
  queuePending: number;
  queueReady: number;
  queueDeadLetter: number;
  queueLastAttemptAt: number | null;
  queueLastError: string | null;
};

export type AnilistStateRuntimeDeps = {
  getClientSecretState: () => AnilistSecretResolutionState;
  setClientSecretState: (next: AnilistSecretResolutionState) => void;
  getRetryQueueState: () => AnilistRetryQueueState;
  setRetryQueueState: (next: AnilistRetryQueueState) => void;
  getUpdateQueueSnapshot: () => AnilistQueueSnapshot;
  clearStoredToken: () => void;
  clearCachedAccessToken: () => void;
};

export function createAnilistStateRuntime(deps: AnilistStateRuntimeDeps): {
  setClientSecretState: (partial: Partial<AnilistSecretResolutionState>) => void;
  refreshRetryQueueState: () => void;
  getStatusSnapshot: () => AnilistStatusSnapshot;
  getQueueStatusSnapshot: () => AnilistRetryQueueState;
  clearTokenState: () => void;
} {
  const setClientSecretState = (partial: Partial<AnilistSecretResolutionState>): void => {
    deps.setClientSecretState({
      ...deps.getClientSecretState(),
      ...partial,
    });
  };

  const refreshRetryQueueState = (): void => {
    deps.setRetryQueueState({
      ...deps.getRetryQueueState(),
      ...deps.getUpdateQueueSnapshot(),
    });
  };

  const getStatusSnapshot = (): AnilistStatusSnapshot => {
    const client = deps.getClientSecretState();
    const queue = deps.getRetryQueueState();
    return {
      tokenStatus: client.status,
      tokenSource: client.source,
      tokenMessage: client.message,
      tokenResolvedAt: client.resolvedAt,
      tokenErrorAt: client.errorAt,
      queuePending: queue.pending,
      queueReady: queue.ready,
      queueDeadLetter: queue.deadLetter,
      queueLastAttemptAt: queue.lastAttemptAt,
      queueLastError: queue.lastError,
    };
  };

  const getQueueStatusSnapshot = (): AnilistRetryQueueState => {
    refreshRetryQueueState();
    const queue = deps.getRetryQueueState();
    return {
      pending: queue.pending,
      ready: queue.ready,
      deadLetter: queue.deadLetter,
      lastAttemptAt: queue.lastAttemptAt,
      lastError: queue.lastError,
    };
  };

  const clearTokenState = (): void => {
    deps.clearStoredToken();
    deps.clearCachedAccessToken();
    setClientSecretState({
      status: 'not_checked',
      source: 'none',
      message: 'stored token cleared',
      resolvedAt: null,
      errorAt: null,
    });
  };

  return {
    setClientSecretState,
    refreshRetryQueueState,
    getStatusSnapshot,
    getQueueStatusSnapshot,
    clearTokenState,
  };
}
