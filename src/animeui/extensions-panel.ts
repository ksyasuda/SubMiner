import { describe, el } from './dom';
import { describeInstalled } from './format';
import {
  collectLanguages,
  filterByLanguage,
  languageLabel,
  pruneSelection,
  toggleLanguage,
} from './language-filter';
import type {
  AnimeBrowserAPI,
  AvailableExtension,
  InstalledExtensionView,
} from '../types/anime-browser';

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

interface RowAction {
  label: string;
  primary?: boolean;
  onClick: () => void | Promise<void>;
}

interface RowOptions {
  name: string;
  sub: string;
  tags?: Array<{ text: string; className: string }>;
  actions?: RowAction[];
  isError?: boolean;
}

function extensionRow(options: RowOptions): HTMLDivElement {
  const row = document.createElement('div');
  row.className = options.isError ? 'ext-row is-error' : 'ext-row';

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
    button.className = action.primary ? 'primary-button' : 'ghost-button';
    button.textContent = action.label;
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

function emptyNote(text: string): HTMLParagraphElement {
  const empty = document.createElement('p');
  empty.className = 'ext-empty';
  empty.textContent = text;
  return empty;
}

export function createExtensionsPanel(options: ExtensionsPanelOptions) {
  const { api, setStatus, onSourcesChanged } = options;

  const extensionsDirLabel = el<HTMLSpanElement>('extensions-dir');
  const installedList = el<HTMLDivElement>('installed-list');
  const installedCount = el<HTMLSpanElement>('installed-count');
  const availableList = el<HTMLDivElement>('extensions-list');
  const availableCount = el<HTMLSpanElement>('available-count');
  const langFilter = el<HTMLDivElement>('lang-filter');
  const repoInput = el<HTMLInputElement>('repo-input');
  const repoAddButton = el<HTMLButtonElement>('repo-add');
  const repoList = el<HTMLDivElement>('repo-list');

  // What the last refresh found, kept so toggling a language chip re-renders
  // the list without re-fetching every repository index.
  let installable: AvailableExtension[] = [];
  let repoFailures: Array<{ name: string; error: string }> = [];
  let hasRepos = false;
  /** Selected language codes; empty means "All". */
  let selectedLangs = new Set<string>();

  async function afterChange(extensionName: string, verb: string): Promise<void> {
    await refresh();
    await onSourcesChanged();
    setStatus(`${extensionName} ${verb}`, 'ok');
  }

  function renderInstalled(
    installed: InstalledExtensionView[],
    offeredPkgs: Set<string>,
    extensionsDir: string,
  ): void {
    installedCount.textContent = installed.length === 0 ? '' : String(installed.length);

    if (installed.length === 0) {
      installedList.replaceChildren(
        emptyNote(
          `Nothing installed yet. Add a repository below, or drop .apk files in ${extensionsDir}.`,
        ),
      );
      return;
    }

    installedList.replaceChildren(
      ...installed.map((view) => {
        const actions: RowAction[] = [];
        // Only offer an update for an extension a configured repository still
        // carries; reinstalling overwrites the APK in place.
        if (offeredPkgs.has(view.pkg)) {
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
          isError: view.error !== null,
          tags: view.error === null ? [] : [{ text: 'failed', className: 'nsfw' }],
          actions,
        });
      }),
    );
  }

  function renderRepos(repos: string[]): void {
    repoList.replaceChildren(
      ...repos.map((repoUrl) =>
        extensionRow({
          name: repoUrl.replace(/^https:\/\//, '').replace(/\/[^/]*\.json$/, ''),
          sub: repoUrl,
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

    const offeredPkgs = new Set(available.extensions.map((extension) => extension.pkg));
    renderInstalled(snapshot.installed, offeredPkgs, snapshot.extensionsDir);
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

  repoInput.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      repoAddButton.click();
    }
  });

  return { refresh };
}
