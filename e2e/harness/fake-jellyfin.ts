import { once } from 'node:events';
import fs from 'node:fs';
import http from 'node:http';
import type { AddressInfo } from 'node:net';

export type FakeJellyfinEpisode = { id: string; name: string; mediaPath: string };

export type FakeJellyfinReport = {
  kind: 'playing' | 'progress' | 'stopped';
  itemId: string;
  positionTicks: number;
};

export type FakeJellyfin = {
  url: string;
  userId: string;
  accessToken: string;
  /** Every playback report the app sent, in arrival order. */
  reports: FakeJellyfinReport[];
  stop: () => Promise<void>;
};

const SERIES_ID = 'e2e-series';
const REPORT_KINDS: Record<string, FakeJellyfinReport['kind']> = {
  '/Sessions/Playing': 'playing',
  '/Sessions/Playing/Progress': 'progress',
  '/Sessions/Playing/Stopped': 'stopped',
};

// Serves one file with Range support, since mpv probes and seeks over HTTP.
function serveFile(request: http.IncomingMessage, response: http.ServerResponse, file: string) {
  const size = fs.statSync(file).size;
  const range = /^bytes=(\d+)-(\d*)$/.exec(request.headers.range ?? '');
  const start = range ? Number(range[1]) : 0;
  const end = Math.min(range?.[2] ? Number(range[2]) : size - 1, size - 1);
  response.writeHead(range ? 206 : 200, {
    'content-type': 'video/x-matroska',
    'accept-ranges': 'bytes',
    'content-length': end - start + 1,
    ...(range ? { 'content-range': `bytes ${start}-${end}/${size}` } : {}),
  });
  fs.createReadStream(file, { start, end }).pipe(response);
}

/**
 * A Jellyfin stand-in holding one series of `episodes`, in order. It answers the item, episode
 * list, PlaybackInfo, and stream requests direct play needs, and records playback reports.
 */
export async function startFakeJellyfin(episodes: FakeJellyfinEpisode[]): Promise<FakeJellyfin> {
  const reports: FakeJellyfinReport[] = [];
  const byId = new Map(episodes.map((episode, index) => [episode.id, { episode, index }]));

  const itemJson = (id: string) => {
    const found = byId.get(id);
    if (!found) return null;
    return {
      Id: id,
      Name: found.episode.name,
      Type: 'Episode',
      SeriesName: 'E2E Show',
      SeriesId: SERIES_ID,
      ParentIndexNumber: 1,
      IndexNumber: found.index + 1,
      LocationType: 'FileSystem',
      UserData: { PlaybackPositionTicks: 0 },
      MediaSources: [
        {
          Id: id,
          Container: 'mkv',
          SupportsDirectStream: true,
          SupportsTranscoding: false,
          MediaStreams: [],
        },
      ],
    };
  };

  const server = http.createServer((request, response) => {
    const url = new URL(request.url ?? '/', 'http://fake-jellyfin');
    const json = (status: number, body: unknown) => {
      response.writeHead(status, { 'content-type': 'application/json' });
      response.end(JSON.stringify(body));
    };

    const stream = /^\/Videos\/([^/]+)\/stream$/.exec(url.pathname);
    if (request.method === 'GET' && stream) {
      const found = byId.get(stream[1]!);
      if (!found) return json(404, {});
      return serveFile(request, response, found.episode.mediaPath);
    }

    const chunks: Buffer[] = [];
    request.on('data', (chunk: Buffer) => chunks.push(chunk));
    request.on('end', () => {
      const reportKind = REPORT_KINDS[url.pathname];
      if (request.method === 'POST' && reportKind) {
        const body = JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}');
        reports.push({
          kind: reportKind,
          itemId: String(body.ItemId ?? ''),
          positionTicks: Number(body.PositionTicks ?? 0),
        });
        response.writeHead(204);
        return response.end();
      }
      const item = /^\/Users\/[^/]+\/Items\/([^/]+)$/.exec(url.pathname);
      if (request.method === 'GET' && item) {
        const body = itemJson(item[1]!);
        return body ? json(200, body) : json(404, {});
      }
      const playbackInfo = /^\/Items\/([^/]+)\/PlaybackInfo$/.exec(url.pathname);
      if (request.method === 'POST' && playbackInfo) {
        return json(200, { MediaSources: itemJson(playbackInfo[1]!)?.MediaSources ?? [] });
      }
      if (request.method === 'GET' && url.pathname === `/Shows/${SERIES_ID}/Episodes`) {
        return json(200, { Items: episodes.map((episode) => itemJson(episode.id)) });
      }
      json(404, {});
    });
  });
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  const { port } = server.address() as AddressInfo;

  return {
    url: `http://127.0.0.1:${port}`,
    userId: 'e2e-user',
    accessToken: 'e2e-token',
    reports,
    stop: async () => {
      server.closeAllConnections();
      server.close();
      await once(server, 'close');
    },
  };
}
