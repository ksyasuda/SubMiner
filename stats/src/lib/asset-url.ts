/**
 * Resolve a bundled public asset against Vite's base URL.
 *
 * The in-player stats window is loaded with `loadFile`, so the document lives on
 * `file://`. A root-absolute path like `/favicon.png` resolves to the filesystem
 * root there and 404s, while the HTTP-served web app resolves it fine. Vite
 * rewrites asset refs in `index.html` but not string literals in JSX, so build
 * the URL from the configured base instead of hardcoding a leading slash.
 */
export function resolveAssetUrl(path: string, base: string): string {
  const normalizedBase = base.endsWith('/') ? base : `${base}/`;
  return `${normalizedBase}${path.replace(/^\/+/, '')}`;
}

function currentBase(): string {
  // Vite injects BASE_URL at build time ('./' per vite.config.ts) and serves '/'
  // in dev. Outside a Vite bundle (tests) there is no env, so fall back to './'.
  const env = (import.meta as { env?: Record<string, string | undefined> }).env;
  return env?.BASE_URL || './';
}

export function assetUrl(path: string): string {
  return resolveAssetUrl(path, currentBase());
}
