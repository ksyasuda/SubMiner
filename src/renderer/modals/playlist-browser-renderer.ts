import type {
  PlaylistBrowserDirectoryItem,
  PlaylistBrowserQueueItem,
} from '../../types';
import type { RendererContext } from '../context';

type PlaylistBrowserRowRenderActions = {
  appendDirectoryItem: (filePath: string) => void;
  movePlaylistItem: (index: number, direction: 1 | -1) => void;
  playPlaylistItem: (index: number) => void;
  removePlaylistItem: (index: number) => void;
  render: () => void;
};

function createActionButton(label: string, onClick: () => void): HTMLButtonElement {
  const button = document.createElement('button');
  button.type = 'button';
  button.textContent = label;
  button.className = 'playlist-browser-action';
  button.addEventListener('click', (event) => {
    event.stopPropagation();
    onClick();
  });
  button.addEventListener('dblclick', (event) => {
    event.preventDefault?.();
    event.stopPropagation();
  });
  return button;
}

export function renderPlaylistBrowserDirectoryRow(
  ctx: RendererContext,
  item: PlaylistBrowserDirectoryItem,
  index: number,
  actions: PlaylistBrowserRowRenderActions,
): HTMLElement {
  const row = document.createElement('li');
  row.className = 'playlist-browser-row';
  if (item.isCurrentFile) row.classList.add('current');
  if (
    ctx.state.playlistBrowserActivePane === 'directory' &&
    ctx.state.playlistBrowserSelectedDirectoryIndex === index
  ) {
    row.classList.add('active');
  }

  const main = document.createElement('div');
  main.className = 'playlist-browser-row-main';
  const label = document.createElement('div');
  label.className = 'playlist-browser-row-label';
  label.textContent = item.basename;
  const meta = document.createElement('div');
  meta.className = 'playlist-browser-row-meta';
  meta.textContent = item.isCurrentFile
    ? item.episodeLabel
      ? `${item.episodeLabel} · Current file`
      : 'Current file'
    : item.episodeLabel ?? 'Video file';
  main.append(label, meta);

  const trailing = document.createElement('div');
  trailing.className = 'playlist-browser-row-trailing';
  if (item.episodeLabel) {
    const badge = document.createElement('div');
    badge.className = 'playlist-browser-chip';
    badge.textContent = item.episodeLabel;
    trailing.appendChild(badge);
  }
  trailing.appendChild(
    createActionButton('Add', () => {
      void actions.appendDirectoryItem(item.path);
    }),
  );

  row.append(main, trailing);
  row.addEventListener('click', () => {
    ctx.state.playlistBrowserActivePane = 'directory';
    ctx.state.playlistBrowserSelectedDirectoryIndex = index;
    actions.render();
  });
  row.addEventListener('dblclick', () => {
    ctx.state.playlistBrowserSelectedDirectoryIndex = index;
    void actions.appendDirectoryItem(item.path);
  });
  return row;
}

export function renderPlaylistBrowserPlaylistRow(
  ctx: RendererContext,
  item: PlaylistBrowserQueueItem,
  index: number,
  actions: PlaylistBrowserRowRenderActions,
): HTMLElement {
  const row = document.createElement('li');
  row.className = 'playlist-browser-row';
  if (item.current || item.playing) row.classList.add('current');
  if (
    ctx.state.playlistBrowserActivePane === 'playlist' &&
    ctx.state.playlistBrowserSelectedPlaylistIndex === index
  ) {
    row.classList.add('active');
  }

  const main = document.createElement('div');
  main.className = 'playlist-browser-row-main';
  const label = document.createElement('div');
  label.className = 'playlist-browser-row-label';
  label.textContent = `${index + 1}. ${item.displayLabel}`;
  const meta = document.createElement('div');
  meta.className = 'playlist-browser-row-meta';
  meta.textContent = item.current || item.playing ? 'Playing now' : 'Queued';
  const submeta = document.createElement('div');
  submeta.className = 'playlist-browser-row-submeta';
  submeta.textContent = item.filename;
  main.append(label, meta, submeta);

  const trailing = document.createElement('div');
  trailing.className = 'playlist-browser-row-actions';
  trailing.append(
    createActionButton('Play', () => {
      void actions.playPlaylistItem(item.index);
    }),
    createActionButton('Up', () => {
      void actions.movePlaylistItem(item.index, -1);
    }),
    createActionButton('Down', () => {
      void actions.movePlaylistItem(item.index, 1);
    }),
    createActionButton('Remove', () => {
      void actions.removePlaylistItem(item.index);
    }),
  );
  row.append(main, trailing);
  row.addEventListener('click', () => {
    ctx.state.playlistBrowserActivePane = 'playlist';
    ctx.state.playlistBrowserSelectedPlaylistIndex = index;
    actions.render();
  });
  row.addEventListener('dblclick', () => {
    ctx.state.playlistBrowserSelectedPlaylistIndex = index;
    void actions.playPlaylistItem(item.index);
  });
  return row;
}
