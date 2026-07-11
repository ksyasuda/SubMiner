const DEFAULT_YOMITAN_API_TIMEOUT_MS = 10_000;

export async function fetchYomitanApi(
  input: RequestInfo | URL,
  init?: RequestInit,
  timeoutMs = DEFAULT_YOMITAN_API_TIMEOUT_MS,
): Promise<Response> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);

  try {
    return await fetch(input, { ...init, signal: controller.signal });
  } catch (error) {
    if (controller.signal.aborted) {
      throw new Error(`Yomitan API request timed out after ${timeoutMs}ms`, { cause: error });
    }
    throw error;
  } finally {
    clearTimeout(timeout);
  }
}
