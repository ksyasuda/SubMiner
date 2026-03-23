import type { YoutubeFlowMode } from '../../../types';
import type { YoutubeTrackOption } from './track-probe';
import { downloadYoutubeSubtitleTrack, downloadYoutubeSubtitleTracks } from './track-download';

export function isYoutubeGenerationMode(mode: YoutubeFlowMode): boolean {
  return mode === 'generate';
}

export async function acquireYoutubeSubtitleTrack(input: {
  targetUrl: string;
  outputDir: string;
  track: YoutubeTrackOption;
  mode: YoutubeFlowMode;
}): Promise<{ path: string }> {
  return await downloadYoutubeSubtitleTrack(input);
}

export async function acquireYoutubeSubtitleTracks(input: {
  targetUrl: string;
  outputDir: string;
  tracks: YoutubeTrackOption[];
  mode: YoutubeFlowMode;
}): Promise<Map<string, string>> {
  return await downloadYoutubeSubtitleTracks(input);
}
