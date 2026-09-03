import type { AnimeBrowserIpcSender } from './anime-browser-ipc-handlers';

export interface AnimeBrowserSessionRegistry {
  get: (sessionId: string) => AnimeBrowserIpcSender | undefined;
  register: (sessionId: string, sender: AnimeBrowserIpcSender) => void;
  values: () => IterableIterator<AnimeBrowserIpcSender>;
}

export function createAnimeBrowserSessionRegistry(
  releaseSession: (sessionId: string) => void,
): AnimeBrowserSessionRegistry {
  const sessions = new Map<string, AnimeBrowserIpcSender>();

  const register = (sessionId: string, sender: AnimeBrowserIpcSender): void => {
    if (sessions.get(sessionId) === sender) return;
    for (const [previousSessionId, registeredSender] of sessions) {
      if (previousSessionId === sessionId || registeredSender !== sender) continue;
      sessions.delete(previousSessionId);
      releaseSession(previousSessionId);
    }
    sessions.set(sessionId, sender);
    sender.once('destroyed', () => {
      if (sessions.get(sessionId) !== sender) return;
      sessions.delete(sessionId);
      releaseSession(sessionId);
    });
  };

  return {
    get: (sessionId) => sessions.get(sessionId),
    register,
    values: () => sessions.values(),
  };
}
