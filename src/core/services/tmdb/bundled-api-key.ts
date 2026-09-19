import fs from 'node:fs';
import path from 'node:path';

/**
 * Release builds stage a project-owned TMDB key into dist/ (see
 * scripts/bundled-integration-keys.mjs). Source checkouts and CI builds have no
 * such file, and TMDB lookups then depend on the user's own `tmdb.apiKey`.
 */
export const BUNDLED_INTEGRATION_KEYS_FILENAME = 'bundled-integration-keys.json';

export function readBundledTmdbApiKey(distDir: string): string | null {
  try {
    const raw = fs.readFileSync(path.join(distDir, BUNDLED_INTEGRATION_KEYS_FILENAME), 'utf8');
    const parsed = JSON.parse(raw) as { tmdbApiKey?: unknown };
    const key = typeof parsed.tmdbApiKey === 'string' ? parsed.tmdbApiKey.trim() : '';
    return key.length > 0 ? key : null;
  } catch {
    return null;
  }
}
