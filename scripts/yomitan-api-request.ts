const DEFAULT_YOMITAN_API_TIMEOUT_MS = 10_000;

export async function fetchYomitanApi(
  input: RequestInfo | URL,
  init?: RequestInit,
  timeoutMs = DEFAULT_YOMITAN_API_TIMEOUT_MS,
): Promise<Response> {
  const controller = new AbortController();
  const requestSignal = input instanceof Request ? input.signal : undefined;
  const signals = [controller.signal, init?.signal, requestSignal].filter(
    (signal): signal is AbortSignal => signal != null,
  );
  const signal = signals.length === 1 ? controller.signal : AbortSignal.any(signals);
  const timeout = setTimeout(() => controller.abort(), timeoutMs);

  try {
    return await fetch(input, { ...init, signal });
  } catch (error) {
    if (controller.signal.aborted) {
      throw new Error(`Yomitan API request timed out after ${timeoutMs}ms`, { cause: error });
    }
    throw error;
  } finally {
    clearTimeout(timeout);
  }
}
