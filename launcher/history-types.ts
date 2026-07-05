export interface HistoryVideoRow {
  videoId: number;
  sourcePath: string;
  parsedTitle: string | null;
  parsedSeason: number | null;
  parsedEpisode: number | null;
  animeTitle: string | null;
  lastWatchedMs: number;
}

export interface HistorySeriesEntry {
  seriesRoot: string;
  displayName: string;
  lastWatched: HistoryVideoRow;
}

export interface SeasonDirEntry {
  name: string;
  path: string;
  season: number | null;
}
