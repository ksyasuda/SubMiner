import assert from 'node:assert/strict';
import test from 'node:test';
import { buildYoutubeMetadataProbeArgs } from './metadata-probe';
import { buildYoutubePlaybackResolveArgs } from './playback-resolve';
import { buildDownloadArgs } from './track-download';
import { buildYoutubeTrackProbeArgs } from './track-probe';
import { YTDLP_SINGLE_VIDEO_ARG } from './ytdlp-command';

// Regression guard for issue #179: a `list=`/`index=` URL made yt-dlp enumerate the whole
// playlist (e.g. Watch Later) and blow past our 15s timeouts on every single-video call.
const PLAYLIST_URL = 'https://www.youtube.com/watch?v=LKfWC6CgFng&list=WL&index=3';

const cases: Array<{ name: string; args: string[] }> = [
  { name: 'track probe', args: buildYoutubeTrackProbeArgs(PLAYLIST_URL) },
  { name: 'metadata probe', args: buildYoutubeMetadataProbeArgs(PLAYLIST_URL) },
  { name: 'playback resolve', args: buildYoutubePlaybackResolveArgs(PLAYLIST_URL, 'b') },
  {
    name: 'subtitle download',
    args: buildDownloadArgs({
      targetUrl: PLAYLIST_URL,
      outputTemplate: '/tmp/out.%(ext)s',
      sourceLanguages: ['ja'],
      includeAutoSubs: true,
      includeManualSubs: false,
    }),
  },
];

test('YTDLP_SINGLE_VIDEO_ARG is the yt-dlp flag that disables playlist expansion', () => {
  assert.equal(YTDLP_SINGLE_VIDEO_ARG, '--no-playlist');
});

for (const { name, args } of cases) {
  test(`${name} passes --no-playlist for playlist-scoped URLs`, () => {
    assert.ok(args.includes('--no-playlist'), `${name} args: ${args.join(' ')}`);
    assert.equal(args.at(-1), PLAYLIST_URL);
  });
}
