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
  groupHistoryBySeries,
  listSeasonDirs,
  queryLocalWatchHistory,
  resolveImmersionDbPath,
  sortVideosByEpisode,
  type HistorySeriesEntry,
} from '../history.js';
import type { Args } from '../types.js';
import type { LauncherCommandContext } from './context.js';

function checkPickerDependencies(args: Args): void {
  if (args.useRofi) {
    if (!commandExists('rofi')) fail('Missing dependency: rofi');
    return;
  }
  if (!commandExists('fzf')) fail('Missing dependency: fzf');
}

function showRofiIndexMenu(labels: string[], prompt: string, themePath: string | null): number {
  const rofiArgs = ['-dmenu', '-i', '-matching', 'fuzzy', '-format', 'i', '-p', prompt];
  if (themePath) {
    rofiArgs.push('-theme', themePath);
  } else {
    rofiArgs.push('-theme-str', 'configuration { font: "Noto Sans CJK JP Regular 8";}');
  }
  const result = spawnSync('rofi', rofiArgs, {
    input: `${labels.join('\n')}\n`,
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
): number {
  if (labels.length === 0) return -1;
  return useRofi ? showRofiIndexMenu(labels, prompt, themePath) : showFzfIndexMenu(labels, prompt);
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
      `${entry.displayName} — Season`,
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

export async function runHistoryCommand(context: LauncherCommandContext): Promise<string | null> {
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

  const seriesIdx = pickIndex(
    series.map(formatSeriesLabel),
    'Watch History',
    args.useRofi,
    themePath,
  );
  if (seriesIdx < 0) return null;
  const entry = series[seriesIdx]!;

  const lastPath = path.resolve(entry.lastWatched.sourcePath);
  const lastExists = fs.existsSync(lastPath);
  const nextEpisode = findNextEpisode(lastPath);

  const actions: Array<{ kind: 'replay' | 'next' | 'browse'; label: string }> = [];
  if (lastExists) {
    actions.push({ kind: 'replay', label: `Replay last watched — ${path.basename(lastPath)}` });
  }
  if (nextEpisode) {
    actions.push({ kind: 'next', label: `Next episode — ${path.basename(nextEpisode)}` });
  }
  actions.push({ kind: 'browse', label: 'Browse episodes' });

  const actionIdx = pickIndex(
    actions.map((action) => action.label),
    entry.displayName,
    args.useRofi,
    themePath,
  );
  if (actionIdx < 0) return null;

  switch (actions[actionIdx]!.kind) {
    case 'replay':
      return lastPath;
    case 'next':
      return nextEpisode;
    case 'browse':
      return browseEpisodes(entry, context, themePath);
  }
}
