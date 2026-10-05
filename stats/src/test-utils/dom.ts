import { Window } from 'happy-dom';

/**
 * Installs a happy-dom window as the global DOM for React tests and returns a
 * function that restores the previous globals.
 */
export function installDom(): () => void {
  const window = new Window();
  const properties = {
    window,
    document: window.document,
    HTMLElement: window.HTMLElement,
    ResizeObserver: window.ResizeObserver,
    IS_REACT_ACT_ENVIRONMENT: true,
  };
  const previous = Object.keys(properties).map(
    (key) => [key, Object.getOwnPropertyDescriptor(globalThis, key)] as const,
  );
  for (const [key, value] of Object.entries(properties)) {
    Object.defineProperty(globalThis, key, { value, configurable: true, writable: true });
  }
  return () => {
    for (const [key, descriptor] of previous) {
      if (descriptor) Object.defineProperty(globalThis, key, descriptor);
      else Reflect.deleteProperty(globalThis, key);
    }
    window.happyDOM.abort();
  };
}

/** Installs an in-memory `localStorage` seeded with `initial`; returns the restore function. */
export function installLocalStorage(initial: Record<string, string> = {}): () => void {
  const previous = Object.getOwnPropertyDescriptor(globalThis, 'localStorage');
  const values = new Map(Object.entries(initial));
  Object.defineProperty(globalThis, 'localStorage', {
    configurable: true,
    value: {
      get length() {
        return values.size;
      },
      clear: () => values.clear(),
      getItem: (key: string) => values.get(key) ?? null,
      key: (index: number) => Array.from(values.keys())[index] ?? null,
      removeItem: (key: string) => values.delete(key),
      setItem: (key: string, value: string) => values.set(key, value),
    },
  });
  return () => {
    if (previous) Object.defineProperty(globalThis, 'localStorage', previous);
    else Reflect.deleteProperty(globalThis, 'localStorage');
  };
}

/** Runs `run` with an in-memory `localStorage` seeded with `initial`. */
export function withLocalStorage<T>(initial: Record<string, string>, run: () => T): T {
  const restore = installLocalStorage(initial);
  try {
    return run();
  } finally {
    restore();
  }
}
