import { describe, el } from './dom';
import { buildIconIndex, iconMonogram, isSafeIconUrl, repoFaviconUrl } from './extension-icons';
import { describeBridgeInstall, describeInstalled } from './format';
import {
  collectLanguages,
  filterByLanguage,
  languageLabel,
  pruneSelection,
  toggleLanguage,
} from './language-filter';
import { ALL_SOURCES_ID } from '../types/anime-browser';
import type {
  AnimeBrowserAPI,
  AvailableExtension,
  InstalledExtensionView,
} from '../types/anime-browser';
import { hasExtensionUpdate } from '../shared/extension-updates';

/**
 * The Extensions tab: what is installed, which repositories feed it, and what
 * those repositories still offer.
 *
 * Installed extensions get their own section at the top, listed from the
 * extensions directory rather than from a repository catalogue — an APK dropped
 * in by hand, or one whose repository has since been removed, is still
 * installed and has to stay removable.
 */

export interface ExtensionsPanelOptions {
  api: AnimeBrowserAPI;
  setStatus: (message: string, tone?: 'info' | 'ok' | 'error') => void;
  /** Called after an install or removal, so the source picker keeps up. */
  onSourcesChanged: () => Promise<void>;
}

type RowAction =
  | {
      label: string;
      primary?: boolean;
      onClick: () => void | Promise<void>;
    }
  | {
      label: string;
      disabled: true;
      title?: string;
    };

interface RowOptions {
  name: string;
  sub: string;
  /** Repository-published icon for the row, when there is one. */
  iconUrl?: string | null;
  tags?: Array<{ text: string; className: string }>;
  actions?: RowAction[];
  isError?: boolean;
}

/**
 * The row's avatar: the extension's icon, with the name's first letter behind
 * it. Repositories do not always publish an icon for every package, so the
 * monogram shows until the image loads and stays put if it never does.
 */
function extensionIcon(name: string, iconUrl: string | null | undefined): HTMLSpanElement {
  const badge = document.createElement('span');
  badge.className = 'ext-icon';
  badge.setAttribute('aria-hidden', 'true');
  badge.textContent = iconMonogram(name);
  if (!isSafeIconUrl(iconUrl)) return badge;

  const image = document.createElement('img');
  image.className = 'ext-icon-img';
  image.loading = 'lazy';
  image.alt = '';
  image.addEventListener('load', () => badge.classList.add('has-icon'));
  image.addEventListener('error', () => image.remove());
  image.src = iconUrl;
  badge.append(image);
  return badge;
}

function extensionRow(options: RowOptions): HTMLDivElement {
  const row = document.createElement('div');
  row.className = options.isError ? 'ext-row is-error' : 'ext-row';
  row.append(extensionIcon(options.name, options.iconUrl));

  const main = document.createElement('div');
  main.className = 'ext-main';
  const name = document.createElement('div');
  name.className = 'ext-name';
  name.textContent = options.name;
  const sub = document.createElement('div');
  sub.className = 'ext-sub';
  sub.textContent = options.sub;
  main.append(name, sub);
  row.append(main);

  for (const tag of options.tags ?? []) {
    const chip = document.createElement('span');
    chip.className = `ext-tag ${tag.className}`;
    chip.textContent = tag.text;
    row.append(chip);
  }

  for (const action of options.actions ?? []) {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'primary' in action && action.primary ? 'primary-button' : 'ghost-button';
    button.textContent = action.label;
    if ('disabled' in action) {
      button.disabled = true;
      if (action.title) button.title = action.title;
      row.append(button);
      continue;
    }
    button.addEventListener('click', () => {
      button.disabled = true;
      void Promise.resolve(action.onClick()).finally(() => {
        button.disabled = false;
      });
    });
    row.append(button);
  }

  return row;
}

export type ExtensionUpdateState = 'available' | 'current' | 'unknown' | 'unavailable';

export type ExtensionUpdateSummary =
  | { kind: 'available'; count: number }
  | { kind: 'current' }
  | { kind: 'none' };

export function getExtensionUpdateState(
  installed: InstalledExtensionView,
  offered: AvailableExtension | undefined,
): ExtensionUpdateState {
  if (!offered) return 'unavailable';
  if (installed.versionCode === null) return 'unknown';
  return hasExtensionUpdate(installed.versionCode, offered.versionCode) ? 'available' : 'current';
}

export function summarizeExtensionUpdates(
  states: readonly ExtensionUpdateState[],
): ExtensionUpdateSummary {
  const count = states.filter((state) => state === 'available').length;
  if (count > 0) return { kind: 'available', count };
  return states.length > 0 && states.every((state) => state === 'current')
    ? { kind: 'current' }
    : { kind: 'none' };
}

function emptyNote(text: string): HTMLParagraphElement {
  const empty = document.createElement('p');
  empty.className = 'ext-empty';
  empty.textContent = text;
  return empty;
}

export function createExtensionsPanel(options: ExtensionsPanelOptions) {
  const { api, setStatus, onSourcesChanged } = options;

  const extensionsDirLabel = el<HTMLSpanElement>('extensions-dir');
  const bridgeInfo = el<HTMLParagraphElement>('bridge-info');
  const installedList = el<HTMLDivElement>('installed-list');
  const installedCount = el<HTMLSpanElement>('installed-count');
  const updateAllButton = el<HTMLButtonElement>('update-all');
  const availableList = el<HTMLDivElement>('extensions-list');
  const availableCount = el<HTMLSpanElement>('available-count');
  const langFilter = el<HTMLDivElement>('lang-filter');
  const repoInput = el<HTMLInputElement>('repo-input');
  const repoAddButton = el<HTMLButtonElement>('repo-add');
  const repoList = el<HTMLDivElement>('repo-list');

  // What the last refresh found, kept so toggling a language chip re-renders
  // the list without re-fetching every repository index.
  let installable: AvailableExtension[] = [];
  /** Package to icon URL, so installed rows can borrow the catalogue's icon. */
  let iconsByPkg = new Map<string, string>();
  let repoFailures: Array<{ name: string; error: string }> = [];
  let hasRepos = false;
  let updatingAll = false;
  let pendingUpdateCount = 0;
  /** Selected language codes; empty means "All". */
  let selectedLangs = new Set<string>();

  async function afterChange(extensionName: string, verb: string): Promise<void> {
    await refresh();
    await onSourcesChanged();
    setStatus(`${extensionName} ${verb}`, 'ok');
  }

  /**
   * The "default" tag and "Set default" buttons for a row's sources. The tag
   * marks the one source (or All sources) the browser opens on; every other
   * source gets a button, so exactly one can carry the tag at a time.
   */
  function defaultSourceControls(
    sources: Array<{ id: string; name: string }>,
    defaultSourceId: string | null,
  ): { tags: RowOptions['tags']; actions: RowAction[] } {
    const tags: NonNullable<RowOptions['tags']> = [];
    const actions: RowAction[] = [];
    const named = sources.length > 1;
    for (const source of sources) {
      if (source.id === defaultSourceId) {
        tags.push({ text: named ? `default · ${source.name}` : 'default', className: 'default' });
        continue;
      }
      actions.push({
        label: named ? `Set default: ${source.name}` : 'Set default',
        onClick: async () => {
          try {
            await api.setDefaultSource(source.id);
            await refresh();
            setStatus(`The browser now opens on ${source.name}.`, 'ok');
          } catch (error) {
            setStatus(describe(error), 'error');
          }
        },
      });
    }
    return { tags, actions };
  }

  function renderInstalled(
    installed: InstalledExtensionView[],
    offeredByPkg: Map<string, AvailableExtension>,
    extensionsDir: string,
    defaultSourceId: string | null,
  ): void {
    installedCount.textContent = installed.length === 0 ? '' : String(installed.length);
    const updateStates = installed.map((view) =>
      getExtensionUpdateState(view, offeredByPkg.get(view.pkg)),
    );
    const updateSummary = summarizeExtensionUpdates(updateStates);
    const updateCount = updateSummary.kind === 'available' ? updateSummary.count : 0;
    pendingUpdateCount = updateCount;
    updateAllButton.textContent =
      updateSummary.kind === 'available'
        ? `Update all (${updateSummary.count})`
        : updateSummary.kind === 'current'
          ? 'All up to date'
          : 'No updates';
    updateAllButton.disabled = updatingAll || updateCount === 0;

    if (installed.length === 0) {
      installedList.replaceChildren(
        emptyNote(
          `Nothing installed yet. Add a repository below, or drop .apk files in ${extensionsDir}.`,
        ),
      );
      return;
    }

    const rows = installed.map((view) => {
      const defaults = defaultSourceControls(view.sources, defaultSourceId);
      const actions: RowAction[] = [...defaults.actions];
      const updateState = getExtensionUpdateState(view, offeredByPkg.get(view.pkg));
      if (updateState === 'available') {
        actions.push({
          label: 'Update',
          onClick: async () => {
            setStatus(`Updating ${view.name}…`);
            try {
              await api.installExtension(view.pkg);
              await afterChange(view.name, 'updated');
            } catch (error) {
              setStatus(describe(error), 'error');
            }
          },
        });
      } else if (updateState === 'current') {
        actions.push({ label: 'Up to date', disabled: true });
      } else if (updateState === 'unknown') {
        actions.push({
          label: 'Version unknown',
          disabled: true,
          title: 'SubMiner could not read a version code from this APK.',
        });
      }
      actions.push({
        label: 'Remove',
        onClick: async () => {
          setStatus(`Removing ${view.name}…`);
          try {
            await api.removeExtension(view.pkg);
            await afterChange(view.name, 'removed');
          } catch (error) {
            setStatus(describe(error), 'error');
          }
        },
      });

      return extensionRow({
        name: view.name,
        sub: view.error ?? describeInstalled(view),
        iconUrl: iconsByPkg.get(view.pkg) ?? null,
        isError: view.error !== null,
        tags: view.error === null ? defaults.tags : [{ text: 'failed', className: 'nsfw' }],
        actions,
      });
    });

    // "All sources" is a picker entry too, so it can be the default like any
    // source. It only exists with more than one source installed.
    const sourceTotal = installed.reduce((sum, view) => sum + view.sourceCount, 0);
    if (sourceTotal > 1) {
      const defaults = defaultSourceControls(
        [{ id: ALL_SOURCES_ID, name: 'All sources' }],
        defaultSourceId,
      );
      rows.unshift(
        extensionRow({
          name: 'All sources',
          sub: `Search every installed source at once · ${sourceTotal} sources`,
          tags: defaults.tags,
          actions: defaults.actions,
        }),
      );
    }

    installedList.replaceChildren(...rows);
  }

  function renderRepos(repos: string[]): void {
    repoList.replaceChildren(
      ...repos.map((repoUrl) =>
        extensionRow({
          name: repoUrl.replace(/^https:\/\//, '').replace(/\/[^/]*\.json$/, ''),
          sub: repoUrl,
          iconUrl: repoFaviconUrl(repoUrl),
          actions: [
            {
              label: 'Remove',
              onClick: async () => {
                setStatus('Removing repository…');
                try {
                  await api.removeRepo(repoUrl);
                  await refresh();
                  setStatus('Repository removed', 'ok');
                } catch (error) {
                  setStatus(describe(error), 'error');
                }
              },
            },
          ],
        }),
      ),
    );
  }

  function languageChip(label: string, active: boolean, onClick: () => void): HTMLButtonElement {
    const chip = document.createElement('button');
    chip.type = 'button';
    chip.className = active ? 'lang-chip is-active' : 'lang-chip';
    chip.textContent = label;
    chip.setAttribute('aria-pressed', active ? 'true' : 'false');
    chip.addEventListener('click', onClick);
    return chip;
  }

  function renderLanguageFilter(languages: string[]): void {
    // One language, or none: there is nothing to choose between.
    if (languages.length < 2) {
      langFilter.replaceChildren();
      return;
    }

    langFilter.replaceChildren(
      // "All" is the empty selection, so picking a language always replaces it
      // rather than sitting alongside it.
      languageChip('All', selectedLangs.size === 0, () => {
        if (selectedLangs.size === 0) return;
        selectedLangs = new Set();
        renderAvailable();
      }),
      ...languages.map((code) =>
        languageChip(languageLabel(code), selectedLangs.has(code), () => {
          selectedLangs = toggleLanguage(selectedLangs, code);
          renderAvailable();
        }),
      ),
    );
  }

  function renderAvailable(): void {
    const languages = collectLanguages(installable);
    selectedLangs = pruneSelection(selectedLangs, languages);
    renderLanguageFilter(languages);

    const shown = filterByLanguage(installable, selectedLangs);
    availableCount.textContent =
      shown.length === installable.length
        ? installable.length === 0
          ? ''
          : String(installable.length)
        : `${shown.length} of ${installable.length}`;

    const rows: HTMLElement[] = repoFailures.map((failure) =>
      extensionRow({ name: failure.name, sub: failure.error, isError: true }),
    );

    for (const extension of shown) {
      rows.push(
        extensionRow({
          name: extension.name,
          sub: `${languageLabel(extension.lang)} · v${extension.version}`,
          iconUrl: extension.iconUrl,
          tags: extension.nsfw ? [{ text: '18+', className: 'nsfw' }] : [],
          actions: [
            {
              label: 'Install',
              primary: true,
              onClick: async () => {
                setStatus(`Installing ${extension.name}…`);
                try {
                  await api.installExtension(extension.pkg);
                  await afterChange(extension.name, 'installed');
                } catch (error) {
                  setStatus(describe(error), 'error');
                }
              },
            },
          ],
        }),
      );
    }

    if (rows.length === 0) {
      rows.push(
        emptyNote(
          selectedLangs.size > 0
            ? 'No available extension matches the selected languages.'
            : hasRepos
              ? 'Every extension the configured repositories offer is already installed.'
              : 'No repository configured, so there is nothing to install from.',
        ),
      );
    }

    availableList.replaceChildren(...rows);
  }

  async function refresh(): Promise<void> {
    const snapshot = await api.getSnapshot();
    extensionsDirLabel.textContent = snapshot.extensionsDir;
    bridgeInfo.textContent = describeBridgeInstall(snapshot.bridge.install);
    renderRepos(snapshot.repos);

    repoFailures = [];
    let available;
    try {
      available = await api.listAvailableExtensions();
    } catch (error) {
      repoFailures.push({ name: 'Repository error', error: describe(error) });
      available = { extensions: [], failures: [] };
    }
    for (const failure of available.failures) {
      repoFailures.push({ name: failure.repoUrl, error: failure.error });
    }

    const offeredByPkg = new Map(
      available.extensions.map((extension) => [extension.pkg, extension]),
    );
    // The catalogue is the only source of icons, so an installed extension can
    // only show one while a repository still carries its package.
    iconsByPkg = buildIconIndex(available.extensions);
    renderInstalled(
      snapshot.installed,
      offeredByPkg,
      snapshot.extensionsDir,
      snapshot.defaultSourceId,
    );
    // Installed extensions have their own section; leaving them here too would
    // list every one of them twice.
    installable = available.extensions.filter((extension) => !extension.installed);
    hasRepos = snapshot.repos.length > 0;
    renderAvailable();
  }

  repoAddButton.addEventListener('click', () => {
    void (async () => {
      const url = repoInput.value.trim();
      if (url.length === 0) return;
      try {
        await api.addRepo(url);
        repoInput.value = '';
        setStatus('Repository added');
        await refresh();
      } catch (error) {
        setStatus(describe(error), 'error');
      }
    })();
  });

  updateAllButton.addEventListener('click', () => {
    if (updateAllButton.disabled) return;
    updatingAll = true;
    updateAllButton.disabled = true;
    updateAllButton.textContent = 'Updating…';
    setStatus('Updating extensions…');
    void (async () => {
      try {
        const count = await api.updateAllExtensions();
        await refresh();
        await onSourcesChanged();
        setStatus(`${count} ${count === 1 ? 'extension' : 'extensions'} updated`, 'ok');
      } catch (error) {
        await refresh().catch(() => undefined);
        await onSourcesChanged().catch(() => undefined);
        setStatus(describe(error), 'error');
      } finally {
        updatingAll = false;
        updateAllButton.disabled = pendingUpdateCount === 0;
        if (updateAllButton.textContent === 'Updating…') {
          updateAllButton.textContent =
            pendingUpdateCount > 0 ? `Update all (${pendingUpdateCount})` : 'No updates';
        }
      }
    })();
  });

  repoInput.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      repoAddButton.click();
    }
  });

  return { refresh };
}
