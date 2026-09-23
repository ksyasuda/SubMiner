/**
 * Icons for the extension rows.
 *
 * An Aniyomi repository publishes each extension's icon — the streaming site's
 * own favicon — next to its APK, so the repository index is the only place an
 * icon can come from. Installed extensions are listed from disk and carry no
 * icon of their own; they borrow the one their package has in the catalogue,
 * which means an extension whose repository was removed, or one dropped in by
 * hand, falls back to a monogram.
 */

interface IconSource {
  pkg: string;
  iconUrl: string;
}

/** Package name to icon URL, for looking up an installed extension's icon. */
export function buildIconIndex(extensions: IconSource[]): Map<string, string> {
  const index = new Map<string, string>();
  for (const extension of extensions) {
    if (typeof extension.iconUrl === 'string' && extension.iconUrl.length > 0) {
      index.set(extension.pkg, extension.iconUrl);
    }
  }
  return index;
}

/** Only https icons are loaded; a repo index is content the user pointed us at. */
export function isSafeIconUrl(url: string | null | undefined): url is string {
  return typeof url === 'string' && /^https:\/\/\S+$/.test(url);
}

/**
 * The favicon of the host serving a repository index.
 *
 * A repository publishes no icon for itself, so the host's own favicon stands
 * in. Nothing depends on it: a host that serves none simply leaves the row on
 * its monogram.
 */
export function repoFaviconUrl(indexUrl: string): string | null {
  try {
    const url = new URL(indexUrl.trim());
    if (url.protocol !== 'https:') return null;
    return `https://${url.host}/favicon.ico`;
  } catch {
    return null;
  }
}

/**
 * The letter shown while an icon loads, or in place of one that never does.
 *
 * Leading punctuation and whitespace are skipped, and a name that carries no
 * letter or digit at all still gets a placeholder rather than an empty box.
 */
export function iconMonogram(name: string): string {
  for (const char of name.trim()) {
    if (/[\p{L}\p{N}]/u.test(char)) return char.toUpperCase();
  }
  return '?';
}
