/**
 * Runs `run` with the given process env values (`undefined` deletes a key) and restores the
 * originals afterwards, even if `run` throws or rejects.
 */
export async function withEnv<T>(
  overrides: Record<string, string | undefined>,
  run: () => T | Promise<T>,
): Promise<T> {
  const originals = Object.fromEntries(
    Object.keys(overrides).map((key) => [key, process.env[key]]),
  );
  const apply = (values: Record<string, string | undefined>): void => {
    for (const [key, value] of Object.entries(values)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
  };

  apply(overrides);
  try {
    return await run();
  } finally {
    apply(originals);
  }
}
