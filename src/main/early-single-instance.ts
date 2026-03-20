interface ElectronSecondInstanceAppLike {
  requestSingleInstanceLock: () => boolean;
  on: (event: 'second-instance', listener: (_event: unknown, argv: string[]) => void) => unknown;
}

export function shouldBypassSingleInstanceLockForArgv(argv: readonly string[]): boolean {
  return argv.includes('--stats-background') || argv.includes('--stats-stop');
}

let cachedSingleInstanceLock: boolean | null = null;
let secondInstanceListenerAttached = false;
const secondInstanceArgvHistory: string[][] = [];
const secondInstanceHandlers = new Set<(_event: unknown, argv: string[]) => void>();

function attachSecondInstanceListener(app: ElectronSecondInstanceAppLike): void {
  if (secondInstanceListenerAttached) return;
  app.on('second-instance', (event, argv) => {
    const clonedArgv = [...argv];
    secondInstanceArgvHistory.push(clonedArgv);
    for (const handler of secondInstanceHandlers) {
      handler(event, [...clonedArgv]);
    }
  });
  secondInstanceListenerAttached = true;
}

export function requestSingleInstanceLockEarly(app: ElectronSecondInstanceAppLike): boolean {
  attachSecondInstanceListener(app);
  if (cachedSingleInstanceLock !== null) {
    return cachedSingleInstanceLock;
  }
  cachedSingleInstanceLock = app.requestSingleInstanceLock();
  return cachedSingleInstanceLock;
}

export function registerSecondInstanceHandlerEarly(
  app: ElectronSecondInstanceAppLike,
  handler: (_event: unknown, argv: string[]) => void,
): () => void {
  attachSecondInstanceListener(app);
  secondInstanceHandlers.add(handler);
  for (const argv of secondInstanceArgvHistory) {
    handler(undefined, [...argv]);
  }
  return () => {
    secondInstanceHandlers.delete(handler);
  };
}

export function resetEarlySingleInstanceStateForTests(): void {
  cachedSingleInstanceLock = null;
  secondInstanceListenerAttached = false;
  secondInstanceArgvHistory.length = 0;
  secondInstanceHandlers.clear();
}
