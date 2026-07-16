import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fail, log } from '../log.js';
import { commandExists } from '../util.js';
import {
  collectVideos,
  findRofiTheme,
  formatPickerLaunchError,
  showFzfMenu,
  showRofiMenu,
} from '../picker.js';
import {
  findNextEpisode,
  findPreviousEpisode,
  groupHistoryBySeries,
  listSeasonDirs,
  materializeCoverArt,
  queryLocalWatchHistory,
  resolveImmersionDbPath,
  sortVideosByEpisode,
  type HistorySeriesEntry,
} from '../history.js';
import type { Args } from '../types.js';
import type { LauncherCommandContext } from './context.js';

export type HistorySessionAction = 'previous' | 'replay' | 'next' | 'browse' | 'quit';

export interface HistoryPlaybackSelection {
  entry: HistorySeriesEntry;
  videoPath: string;
  themePath?: string | null;
  entryIcon?: string | null;
}

interface HistorySessionMenuAction {
  kind: HistorySessionAction;
  label: string;
}

export function buildHistorySessionActions(
  justPlayedPath: string,
  previousEpisodePath: string | null,
  nextEpisodePath: string | null,
): HistorySessionMenuAction[] {
  const actions: HistorySessionMenuAction[] = [];
  if (previousEpisodePath) {
    actions.push({
      kind: 'previous',
      label: `Previous episode: ${path.basename(previousEpisodePath)}`,
    });
  }
  actions.push({
    kind: 'replay',
    label: `Rewatch episode: ${path.basename(justPlayedPath)}`,
  });
  if (nextEpisodePath) {
    actions.push({
      kind: 'next',
      label: `Play next episode: ${path.basename(nextEpisodePath)}`,
    });
  }
  actions.push(
    { kind: 'browse', label: 'Select / browse episode' },
    { kind: 'quit', label: 'Quit SubMiner' },
  );
  return actions;
}

interface HistoryPlaybackLoopDeps {
  play: (videoPath: string) => Promise<void>;
  pickPostPlaybackAction: (input: {
    entry: HistorySeriesEntry;
    justPlayedPath: string;
    previousEpisodePath: string | null;
    nextEpisodePath: string | null;
  }) => Promise<HistorySessionAction | null>;
  findPreviousEpisode: (videoPath: string) => string | null;
  findNextEpisode: (videoPath: string) => string | null;
  browseEpisodes: (entry: HistorySeriesEntry) => Promise<string | null>;
}

export async function runHistoryPlaybackLoop(
  initial: HistoryPlaybackSelection,
  deps: HistoryPlaybackLoopDeps,
): Promise<void> {
  let videoPath = initial.videoPath;

  while (true) {
    await deps.play(videoPath);
    const previousEpisodePath = deps.findPreviousEpisode(videoPath);
    const nextEpisodePath = deps.findNextEpisode(videoPath);
    const action = await deps.pickPostPlaybackAction({
      entry: initial.entry,
      justPlayedPath: videoPath,
      previousEpisodePath,
      nextEpisodePath,
    });

    switch (action) {
      case 'replay':
        break;
      case 'previous':
        if (!previousEpisodePath) return;
        videoPath = previousEpisodePath;
        break;
      case 'next':
        if (!nextEpisodePath) return;
        videoPath = nextEpisodePath;
        break;
      case 'browse': {
        const browsedPath = await deps.browseEpisodes(initial.entry);
        if (!browsedPath) return;
        videoPath = browsedPath;
        break;
      }
      case 'quit':
      case null:
        return;
    }
  }
}

export async function runHistorySession(
  context: LauncherCommandContext,
  play: (videoPath: string) => Promise<void>,
): Promise<boolean> {
  const initial = await runHistoryCommand(context);
  if (!initial) return false;

  await runHistoryPlaybackLoop(initial, {
    play,
    findPreviousEpisode,
    findNextEpisode,
    browseEpisodes: async (entry) => browseEpisodes(entry, context, initial.themePath ?? null),
    pickPostPlaybackAction: async ({
      entry,
      justPlayedPath,
      previousEpisodePath,
      nextEpisodePath,
    }) => {
      const actions = buildHistorySessionActions(
        justPlayedPath,
        previousEpisodePath,
        nextEpisodePath,
      );
      const actionIdx = pickIndex(
        actions.map((action) => action.label),
        entry.displayName,
        context.args.useRofi,
        initial.themePath ?? null,
        actions.map(() => initial.entryIcon ?? null),
      );
      return actionIdx < 0 ? null : actions[actionIdx]!.kind;
    },
  });
  return true;
}

function checkPickerDependencies(args: Args): void {
  if (args.useRofi) {
    if (!commandExists('rofi')) fail('Missing dependency: rofi');
    return;
  }
  if (!commandExists('fzf')) fail('Missing dependency: fzf');
}

function showRofiIndexMenu(
  labels: string[],
  prompt: string,
  themePath: string | null,
  icons: Array<string | null> = [],
): number {
  const rofiArgs = ['-dmenu', '-i', '-matching', 'fuzzy', '-format', 'i', '-p', prompt];
  const hasIcons = icons.some(Boolean);
  if (hasIcons) rofiArgs.push('-show-icons');
  if (themePath) {
    rofiArgs.push('-theme', themePath);
  } else {
    rofiArgs.push('-theme-str', 'configuration { font: "Noto Sans CJK JP Regular 8";}');
  }
  if (hasIcons) {
    rofiArgs.push('-theme-str', 'configuration { show-icons: true; }');
    rofiArgs.push('-theme-str', 'element-icon { enabled: true; size: 3em; }');
  }
  const lines = labels.map((label, index) =>
    icons[index] ? `${label}\u0000icon\u001f${icons[index]}` : label,
  );
  const result = spawnSync('rofi', rofiArgs, {
    input: `${lines.join('\n')}\n`,
    encoding: 'utf8',
    stdio: ['pipe', 'pipe', 'ignore'],
  });
  if (result.error) {
    fail(formatPickerLaunchError('rofi', result.error as NodeJS.ErrnoException));
  }
  const out = (result.stdout || '').trim();
  if (!out) return -1;
  const idx = Number.parseInt(out, 10);
  return Number.isInteger(idx) && idx >= 0 && idx < labels.length ? idx : -1;
}

function showFzfIndexMenu(labels: string[], prompt: string): number {
  const lines = labels.map((label, index) => `${index}\t${label}`);
  const result = spawnSync(
    'fzf',
    [
      '--ansi',
      '--reverse',
      '--ignore-case',
      `--prompt=${prompt}: `,
      '--delimiter=\t',
      '--with-nth=2..',
    ],
    {
      input: `${lines.join('\n')}\n`,
      encoding: 'utf8',
      stdio: ['pipe', 'pipe', 'inherit'],
    },
  );
  if (result.error) {
    fail(formatPickerLaunchError('fzf', result.error as NodeJS.ErrnoException));
  }
  const picked = (result.stdout || '').trim();
  const tab = picked.indexOf('\t');
  if (tab === -1) return -1;
  const idx = Number.parseInt(picked.slice(0, tab), 10);
  return Number.isInteger(idx) && idx >= 0 && idx < labels.length ? idx : -1;
}

function pickIndex(
  labels: string[],
  prompt: string,
  useRofi: boolean,
  themePath: string | null,
  icons: Array<string | null> = [],
): number {
  if (labels.length === 0) return -1;
  return useRofi
    ? showRofiIndexMenu(labels, prompt, themePath, icons)
    : showFzfIndexMenu(labels, prompt);
}

function formatEpisodeLabel(entry: HistorySeriesEntry): string {
  const { parsedSeason, parsedEpisode } = entry.lastWatched;
  if (parsedEpisode === null) return '';
  return parsedSeason !== null ? `S${parsedSeason}E${parsedEpisode}` : `E${parsedEpisode}`;
}

function formatSeriesLabel(entry: HistorySeriesEntry): string {
  const episodeLabel = formatEpisodeLabel(entry);
  return episodeLabel ? `${entry.displayName}  [last: ${episodeLabel}]` : entry.displayName;
}

function pickEpisodeFromDir(dir: string, context: LauncherCommandContext): string | null {
  const { args, scriptPath } = context;
  const videos = sortVideosByEpisode(collectVideos(dir, false));
  if (videos.length === 0) {
    fail(`No video files found in: ${dir}`);
  }
  const selected = args.useRofi
    ? showRofiMenu(videos, dir, false, scriptPath, args.logLevel)
    : showFzfMenu(videos);
  return selected || null;
}

function browseEpisodes(
  entry: HistorySeriesEntry,
  context: LauncherCommandContext,
  themePath: string | null,
): string | null {
  const { args } = context;
  const seasons = listSeasonDirs(entry.seriesRoot);
  let dir = entry.seriesRoot;

  if (seasons.length > 1) {
    const idx = pickIndex(
      seasons.map((season) => season.name),
      `${entry.displayName}: Season`,
      args.useRofi,
      themePath,
    );
    if (idx < 0) return null;
    dir = seasons[idx]!.path;
  } else if (seasons.length === 1 && collectVideos(dir, false).length === 0) {
    dir = seasons[0]!.path;
  }

  return pickEpisodeFromDir(dir, context);
}

export async function runHistoryCommand(
  context: LauncherCommandContext,
): Promise<HistoryPlaybackSelection | null> {
  const { args, scriptPath } = context;

  checkPickerDependencies(args);
  const themePath = args.useRofi ? findRofiTheme(scriptPath) : null;

  const dbPath = resolveImmersionDbPath();
  if (!fs.existsSync(dbPath)) {
    fail(`Watch history database not found: ${dbPath}`);
  }

  const rows = queryLocalWatchHistory(dbPath);
  const series = groupHistoryBySeries(rows);
  if (series.length === 0) {
    fail('No local watch history found (or watched directories are not accessible).');
  }

  log('info', args.logLevel, `Watch history: ${series.length} series found in ${dbPath}`);

  const coverPaths = args.useRofi
    ? materializeCoverArt(
        dbPath,
        series.map((seriesEntry) => seriesEntry.coverBlobHash),
      )
    : new Map<string, string>();
  const seriesIcons = series.map((seriesEntry) =>
    seriesEntry.coverBlobHash ? (coverPaths.get(seriesEntry.coverBlobHash) ?? null) : null,
  );

  const seriesIdx = pickIndex(
    series.map(formatSeriesLabel),
    'Watch History',
    args.useRofi,
    themePath,
    seriesIcons,
  );
  if (seriesIdx < 0) return null;
  const entry = series[seriesIdx]!;

  const lastPath = path.resolve(entry.lastWatched.sourcePath);
  const lastExists = fs.existsSync(lastPath);
  const nextEpisode = findNextEpisode(lastPath);

  const actions: HistorySessionMenuAction[] = [];
  if (lastExists) {
    actions.push({ kind: 'replay', label: `Replay last watched: ${path.basename(lastPath)}` });
  }
  if (nextEpisode) {
    actions.push({ kind: 'next', label: `Next episode: ${path.basename(nextEpisode)}` });
  }
  actions.push({ kind: 'browse', label: 'Browse episodes' });
  actions.push({ kind: 'quit', label: 'Quit SubMiner' });

  const entryIcon = seriesIcons[seriesIdx] ?? null;
  const actionIdx = pickIndex(
    actions.map((action) => action.label),
    entry.displayName,
    args.useRofi,
    themePath,
    actions.map(() => entryIcon),
  );
  if (actionIdx < 0) return null;

  switch (actions[actionIdx]!.kind) {
    case 'replay':
      return { entry, videoPath: lastPath, themePath, entryIcon };
    case 'next':
      return nextEpisode ? { entry, videoPath: nextEpisode, themePath, entryIcon } : null;
    case 'browse': {
      const videoPath = browseEpisodes(entry, context, themePath);
      return videoPath ? { entry, videoPath, themePath, entryIcon } : null;
    }
    case 'previous':
    case 'quit':
      return null;
  }
}
