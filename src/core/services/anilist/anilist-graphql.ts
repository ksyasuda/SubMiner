import type { AnilistRateLimiter } from './rate-limiter';

const ANILIST_GRAPHQL_URL = 'https://graphql.anilist.co';

export class AnilistRateLimitedError extends Error {
  constructor() {
    super('Anilist rate limit reached');
    this.name = 'AnilistRateLimitedError';
  }
}

/**
 * Runs one anonymous AniList GraphQL query through the shared rate limiter.
 * Throws AnilistRateLimitedError on HTTP 429 (the limiter has already recorded
 * the retry window) and a plain Error for other HTTP or GraphQL failures.
 */
export async function executeAnilistQuery<T>(
  rateLimiter: AnilistRateLimiter,
  query: string,
  variables: Record<string, unknown>,
): Promise<T> {
  await rateLimiter.acquire();

  const res = await fetch(ANILIST_GRAPHQL_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
    body: JSON.stringify({ query, variables }),
  });

  rateLimiter.recordResponse(res.headers);

  if (res.status === 429) {
    throw new AnilistRateLimitedError();
  }

  if (!res.ok) {
    throw new Error(`Anilist search failed: ${res.status} ${res.statusText}`);
  }

  const json = (await res.json()) as { data?: T; errors?: Array<{ message?: string }> };
  const firstError = json.errors?.find((entry) => Boolean(entry?.message));
  if (firstError?.message) {
    throw new Error(firstError.message);
  }
  if (!json.data) {
    throw new Error('Anilist response missing data');
  }
  return json.data;
}
