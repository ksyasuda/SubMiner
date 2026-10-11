export type WaitOptions = {
  description: string;
  timeoutMs?: number;
  intervalMs?: number;
};

export type Truthy<T> = Exclude<T, false | 0 | '' | null | undefined>;

/** Polls `probe` until it returns a truthy value, then returns that value. */
export async function waitUntil<T>(
  probe: () => T | Promise<T>,
  { description, timeoutMs = 15_000, intervalMs = 100 }: WaitOptions,
): Promise<Truthy<T>> {
  const deadline = Date.now() + timeoutMs;
  for (;;) {
    const value = await probe();
    if (value) return value as Truthy<T>;
    if (Date.now() >= deadline) {
      throw new Error(`Timed out after ${timeoutMs}ms waiting for ${description}`);
    }
    await new Promise((resolve) => setTimeout(resolve, intervalMs));
  }
}
