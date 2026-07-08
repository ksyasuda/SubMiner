import type { Hono } from 'hono';
import type { ImmersionTrackerService } from './immersion-tracker-service.js';

type StatsCoverImagePayload = {
  contentType: string;
  dataUrl: string;
} | null;

type StatsCoverBatchBody = {
  animeIds?: unknown;
  videoIds?: unknown;
};

const MAX_BACKGROUND_ANIME_COVER_FETCHES = 3;

function parseIntQuery(raw: string | undefined, fallback: number, maxLimit?: number): number {
  if (raw === undefined) return fallback;
  const n = Number(raw);
  if (!Number.isFinite(n) || n < 0) {
    return fallback;
  }
  const parsed = Math.floor(n);
  return maxLimit === undefined ? parsed : Math.min(parsed, maxLimit);
}

function parsePositiveIdList(raw: unknown, maxItems = 100): number[] {
  if (!Array.isArray(raw)) return [];

  const ids = new Set<number>();
  for (const rawId of raw) {
    const id = typeof rawId === 'number' ? rawId : typeof rawId === 'string' ? Number(rawId) : NaN;
    if (Number.isFinite(id) && id > 0) {
      ids.add(Math.floor(id));
      if (ids.size >= maxItems) break;
    }
  }

  return Array.from(ids).sort((a, b) => a - b);
}

function coverImagePayload(
  art: { coverBlob?: Uint8Array | null } | null | undefined,
): StatsCoverImagePayload {
  if (!art?.coverBlob) return null;
  const bytes = new Uint8Array(art.coverBlob);
  const contentType = detectImageContentType(bytes);
  return {
    contentType,
    dataUrl: `data:${contentType};base64,${Buffer.from(bytes).toString('base64')}`,
  };
}

function detectImageContentType(bytes: Uint8Array): string {
  if (
    bytes.length >= 8 &&
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47
  ) {
    return 'image/png';
  }
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) {
    return 'image/jpeg';
  }
  if (
    bytes.length >= 12 &&
    bytes[0] === 0x52 &&
    bytes[1] === 0x49 &&
    bytes[2] === 0x46 &&
    bytes[3] === 0x46 &&
    bytes[8] === 0x57 &&
    bytes[9] === 0x45 &&
    bytes[10] === 0x42 &&
    bytes[11] === 0x50
  ) {
    return 'image/webp';
  }
  return 'application/octet-stream';
}

function createLimitedTaskRunner(maxConcurrentTasks: number): (task: () => Promise<void>) => void {
  const queue: Array<() => Promise<void>> = [];
  let activeTasks = 0;

  const drain = (): void => {
    while (activeTasks < maxConcurrentTasks && queue.length > 0) {
      const task = queue.shift();
      if (!task) return;
      activeTasks += 1;
      void task()
        .catch(() => {})
        .finally(() => {
          activeTasks -= 1;
          drain();
        });
    }
  };

  return (task: () => Promise<void>): void => {
    queue.push(task);
    drain();
  };
}

export function registerStatsCoverRoutes(app: Hono, tracker: ImmersionTrackerService): void {
  const enqueueAnimeCoverBackfill = createLimitedTaskRunner(MAX_BACKGROUND_ANIME_COVER_FETCHES);

  app.post('/api/stats/covers', async (c) => {
    const body = (await c.req.json().catch(() => null)) as StatsCoverBatchBody | null;
    const animeIds = parsePositiveIdList(body?.animeIds);
    const videoIds = parsePositiveIdList(body?.videoIds);
    const anime: Record<number, StatsCoverImagePayload> = {};
    const media: Record<number, StatsCoverImagePayload> = {};

    await Promise.all(
      animeIds.map(async (animeId) => {
        const art = await tracker.getAnimeCoverArt(animeId);
        if (!art?.coverBlob) {
          enqueueAnimeCoverBackfill(async () => {
            await tracker.ensureAnimeCoverArt(animeId);
          });
        }
        anime[animeId] = coverImagePayload(art);
      }),
    );
    await Promise.all(
      videoIds.map(async (videoId) => {
        media[videoId] = coverImagePayload(await tracker.getCoverArt(videoId));
      }),
    );

    return c.json({ anime, media });
  });

  app.get('/api/stats/anime/:animeId/cover', async (c) => {
    const animeId = parseIntQuery(c.req.param('animeId'), 0);
    if (animeId <= 0) return c.body(null, 404);
    let art = await tracker.getAnimeCoverArt(animeId);
    if (!art?.coverBlob) {
      await tracker.ensureAnimeCoverArt(animeId);
      art = await tracker.getAnimeCoverArt(animeId);
    }
    if (!art?.coverBlob) return c.body(null, 404);
    const bytes = new Uint8Array(art.coverBlob);
    return new Response(bytes, {
      headers: {
        'Content-Type': detectImageContentType(bytes),
        'Cache-Control': 'public, max-age=86400',
      },
    });
  });

  app.get('/api/stats/media/:videoId/cover', async (c) => {
    const videoId = parseIntQuery(c.req.param('videoId'), 0);
    if (videoId <= 0) return c.body(null, 404);
    let art = await tracker.getCoverArt(videoId);
    if (!art?.coverBlob) {
      await tracker.ensureCoverArt(videoId);
      art = await tracker.getCoverArt(videoId);
    }
    if (!art?.coverBlob) return c.body(null, 404);
    const bytes = new Uint8Array(art.coverBlob);
    return new Response(bytes, {
      headers: {
        'Content-Type': detectImageContentType(bytes),
        'Cache-Control': 'public, max-age=604800',
      },
    });
  });
}
