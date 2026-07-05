export interface HistoryVideoRow {
  videoId: number;
  sourcePath: string;
  parsedTitle: string | null;
  parsedSeason: number | null;
  parsedEpisode: number | null;
  animeTitle: string | null;
  lastWatchedMs: number;
  coverBlobHash: string | null;
}

export interface HistorySeriesEntry {
  seriesRoot: string;
  displayName: string;
  lastWatched: HistoryVideoRow;
  coverBlobHash: string | null;
}

export interface SeasonDirEntry {
  name: string;
  path: string;
  season: number | null;
}
