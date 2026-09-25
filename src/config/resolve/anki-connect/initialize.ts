import type { ResolveContext } from '../context';

export function initializeAnkiConnectResolution(context: ResolveContext): void {
  context.resolved.ankiConnect = {
    ...context.resolved.ankiConnect,
    fields: {
      ...context.resolved.ankiConnect.fields,
    },
    media: {
      ...context.resolved.ankiConnect.media,
    },
    knownWords: {
      ...context.resolved.ankiConnect.knownWords,
    },
    behavior: {
      ...context.resolved.ankiConnect.behavior,
    },
    proxy: {
      ...context.resolved.ankiConnect.proxy,
    },
    metadata: {
      ...context.resolved.ankiConnect.metadata,
    },
    isLapis: {
      ...context.resolved.ankiConnect.isLapis,
    },
    isKiku: {
      ...context.resolved.ankiConnect.isKiku,
    },
    isSenren: {
      ...context.resolved.ankiConnect.isSenren,
    },
    lapisKiku: {
      ...context.resolved.ankiConnect.lapisKiku,
    },
  };
}
