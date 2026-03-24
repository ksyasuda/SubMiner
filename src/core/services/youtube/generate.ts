import type { YoutubeTrackOption } from './track-probe';
import { downloadYoutubeSubtitleTrack, downloadYoutubeSubtitleTracks } from './track-download';

export async function acquireYoutubeSubtitleTrack(input: {
  targetUrl: string;
  outputDir: string;
  track: YoutubeTrackOption;
}): Promise<{ path: string }> {
  return await downloadYoutubeSubtitleTrack(input);
}

export async function acquireYoutubeSubtitleTracks(input: {
  targetUrl: string;
  outputDir: string;
  tracks: YoutubeTrackOption[];
}): Promise<Map<string, string>> {
  return await downloadYoutubeSubtitleTracks(input);
}
