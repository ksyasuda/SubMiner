/**
 * The language filter for the available-extensions list.
 *
 * A repository index carries every language it knows about, which is far more
 * than any one user reads. The selection is a set of language codes; the empty
 * set means "All", so picking a language always replaces "All" rather than
 * sitting alongside it.
 */

/** The code a repository uses for an extension whose sources span languages. */
export const MULTI_LANG = 'all';

let displayNames: Intl.DisplayNames | null | undefined;

function languageDisplayNames(): Intl.DisplayNames | null {
  if (displayNames === undefined) {
    try {
      displayNames = new Intl.DisplayNames(['en'], { type: 'language' });
    } catch {
      displayNames = null;
    }
  }
  return displayNames;
}

/** `ja` → `Japanese`, `pt-BR` → `Brazilian Portuguese`, `all` → `Multi-language`. */
export function languageLabel(code: string): string {
  if (code === MULTI_LANG) return 'Multi-language';
  try {
    const name = languageDisplayNames()?.of(code);
    // Intl echoes the input back when it knows no name for the tag.
    if (name && name !== code) return name;
  } catch {
    // An invalid tag throws; fall through to the raw code.
  }
  return code.toUpperCase();
}

/**
 * The languages offered, ordered for the chip row: multi-language first
 * because it is the odd one out, then alphabetically by display name.
 */
export function collectLanguages(extensions: ReadonlyArray<{ lang: string }>): string[] {
  const codes = [...new Set(extensions.map((extension) => extension.lang))];
  return codes.sort((a, b) => {
    if (a === MULTI_LANG || b === MULTI_LANG) return a === MULTI_LANG ? -1 : 1;
    return languageLabel(a).localeCompare(languageLabel(b));
  });
}

/**
 * The selection after clicking a language chip. Toggling the last selected
 * language off leaves the empty set, which is "All" again.
 */
export function toggleLanguage(selected: ReadonlySet<string>, code: string): Set<string> {
  const next = new Set(selected);
  if (!next.delete(code)) next.add(code);
  return next;
}

/**
 * Drop the codes no repository offers any more, so a language that disappears
 * when a repository is removed does not keep filtering the list invisibly.
 */
export function pruneSelection(
  selected: ReadonlySet<string>,
  available: ReadonlyArray<string>,
): Set<string> {
  const offered = new Set(available);
  return new Set([...selected].filter((code) => offered.has(code)));
}

/** The extensions to show; an empty selection shows everything. */
export function filterByLanguage<T extends { lang: string }>(
  extensions: ReadonlyArray<T>,
  selected: ReadonlySet<string>,
): T[] {
  if (selected.size === 0) return [...extensions];
  return extensions.filter((extension) => selected.has(extension.lang));
}
