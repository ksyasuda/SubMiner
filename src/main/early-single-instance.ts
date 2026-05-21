interface ElectronSecondInstanceAppLike {
  requestSingleInstanceLock: (additionalData?: Record<string, unknown>) => boolean;
  on: (
    event: 'second-instance',
    listener: (
      _event: unknown,
      argv: string[],
      workingDirectory?: string,
      additionalData?: unknown,
    ) => void,
  ) => unknown;
}

const SECOND_INSTANCE_ARGV_KEY = 'subminerArgv';

export function shouldBypassSingleInstanceLockForArgv(argv: readonly string[]): boolean {
  return argv.includes('--stats-background') || argv.includes('--stats-stop');
}

let cachedSingleInstanceLock: boolean | null = null;
let secondInstanceListenerAttached = false;
const secondInstanceArgvHistory: string[][] = [];
const secondInstanceHandlers = new Set<(_event: unknown, argv: string[]) => void>();

function normalizeSecondInstanceArgv(argv: string[], additionalData: unknown): string[] {
  if (
    additionalData &&
    typeof additionalData === 'object' &&
    Array.isArray((additionalData as { subminerArgv?: unknown }).subminerArgv) &&
    (additionalData as { subminerArgv: unknown[] }).subminerArgv.every(
      (value) => typeof value === 'string',
    )
  ) {
    return [...(additionalData as { subminerArgv: string[] }).subminerArgv];
  }
  return [...argv];
}

function attachSecondInstanceListener(app: ElectronSecondInstanceAppLike): void {
  if (secondInstanceListenerAttached) return;
  app.on('second-instance', (event, argv, _workingDirectory, additionalData) => {
    const clonedArgv = normalizeSecondInstanceArgv(argv, additionalData);
    secondInstanceArgvHistory.push(clonedArgv);
    for (const handler of secondInstanceHandlers) {
      handler(event, [...clonedArgv]);
    }
  });
  secondInstanceListenerAttached = true;
}

export function requestSingleInstanceLockEarly(
  app: ElectronSecondInstanceAppLike,
  argv: readonly string[] = process.argv,
): boolean {
  attachSecondInstanceListener(app);
  if (cachedSingleInstanceLock !== null) {
    return cachedSingleInstanceLock;
  }
  cachedSingleInstanceLock = app.requestSingleInstanceLock({
    [SECOND_INSTANCE_ARGV_KEY]: [...argv],
  });
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
