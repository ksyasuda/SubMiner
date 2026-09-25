import fs from 'node:fs';
import path from 'node:path';

/**
 * Release builds carry a project-owned TMDB key so live-action lookups work
 * without user setup. The key is injected from the SUBMINER_TMDB_API_KEY
 * environment variable at build time (a GitHub Actions secret in CI) and never
 * lives in the repository. The runtime reader is
 * src/core/services/tmdb/bundled-api-key.ts; keep the file name in sync.
 */
export const BUNDLED_INTEGRATION_KEYS_FILENAME = 'bundled-integration-keys.json';
export const TMDB_API_KEY_ENV = 'SUBMINER_TMDB_API_KEY';

/**
 * Write the bundled keys file into `distDir`, or remove a stale one when no
 * key is present so a keyless build never ships an older key by accident.
 * Returns the names of the keys staged.
 */
export function stageBundledIntegrationKeys(distDir, env = process.env) {
  const outputPath = path.join(distDir, BUNDLED_INTEGRATION_KEYS_FILENAME);
  const tmdbApiKey = env[TMDB_API_KEY_ENV]?.trim() ?? '';
  if (!tmdbApiKey) {
    fs.rmSync(outputPath, { force: true });
    return [];
  }
  fs.mkdirSync(distDir, { recursive: true });
  fs.writeFileSync(outputPath, `${JSON.stringify({ tmdbApiKey })}\n`, { mode: 0o644 });
  return ['tmdb'];
}
