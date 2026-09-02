import type {
  AnimeBrowserBridgeInstall,
  AnimeBrowserSearchResult,
  AnimeBrowserSource,
  InstalledExtensionView,
} from '../types/anime-browser';

/** Display strings for the anime browser, kept separate from the DOM. */

/** `Nyaa (ja)`, or just the name for a source that serves every language. */
export function sourceOptionLabel(source: AnimeBrowserSource): string {
  return source.lang === 'all' ? source.name : `${source.name} (${source.lang})`;
}

/**
 * The status line after a search.
 *
 * An all-sources search can half-succeed, so the sources that failed are named
 * rather than folded into a count the user cannot act on.
 */
export function summarizeSearch(result: AnimeBrowserSearchResult): string {
  const count = `${result.entries.length} result${result.entries.length === 1 ? '' : 's'}`;
  if (result.failures.length === 0) return count;
  const names = result.failures.map((failure) => failure.sourceName).join(', ');
  return `${count} · ${result.failures.length} unavailable: ${names}`;
}

/** `pkg · 3 sources · en, ja`, trimmed to what the extension actually reported. */
export function describeInstalled(view: InstalledExtensionView): string {
  const parts = [view.pkg];
  if (view.sourceCount > 0) {
    parts.push(`${view.sourceCount} source${view.sourceCount === 1 ? '' : 's'}`);
  }
  if (view.langs.length > 0) parts.push(view.langs.join(', '));
  return parts.join(' · ');
}

/** One line for the Extensions tab: which bridge is running and who updates it. */
export function describeBridgeInstall(install: AnimeBrowserBridgeInstall | null): string {
  if (install === null) return 'The extension bridge has not started yet.';
  const version = install.version ?? 'unknown version';
  if (install.origin === 'system') {
    return `M-Extension-Server ${version} from ${install.dir}, installed outside SubMiner (for example by your package manager), which is where updates come from.`;
  }
  if (install.updateAvailable !== null) {
    return `M-Extension-Server ${version} in ${install.dir}, downloaded by SubMiner. ${install.updateAvailable} is available from the banner above.`;
  }
  return `M-Extension-Server ${version} in ${install.dir}, downloaded by SubMiner and up to date.`;
}
