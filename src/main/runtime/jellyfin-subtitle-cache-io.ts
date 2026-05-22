import * as path from 'path';

type JellyfinSubtitleCacheTrack = {
  index: number;
  deliveryUrl?: string | null;
};

type JellyfinSubtitleCacheEntry = {
  path: string;
  cleanupDir: string;
};

type FetchResponseLike = {
  ok: boolean;
  status: number;
  arrayBuffer: () => Promise<ArrayBuffer>;
};

type JellyfinSubtitleCacheIoDeps = {
  tmpDir: () => string;
  makeTempDir: (prefix: string) => Promise<string>;
  writeFile: (filePath: string, bytes: Uint8Array) => Promise<void>;
  removeDir: (dir: string, options: { recursive: true; force: true }) => void;
  fetch: (url: string) => Promise<FetchResponseLike>;
};

function getSubtitleExtension(deliveryUrl: string): string {
  const urlPath = (() => {
    try {
      return new URL(deliveryUrl).pathname;
    } catch {
      return deliveryUrl;
    }
  })();
  return path.extname(urlPath).slice(0, 16) || '.srt';
}

export function createJellyfinSubtitleCacheIo(deps: JellyfinSubtitleCacheIoDeps) {
  return {
    async cacheSubtitleTrack(
      track: JellyfinSubtitleCacheTrack,
    ): Promise<JellyfinSubtitleCacheEntry> {
      if (!track.deliveryUrl) {
        throw new Error('Jellyfin subtitle track has no delivery URL');
      }

      const cacheDir = await deps.makeTempDir(
        path.join(deps.tmpDir(), 'subminer-jellyfin-subtitles-'),
      );
      const subtitlePath = path.join(
        cacheDir,
        `track-${track.index}${getSubtitleExtension(track.deliveryUrl)}`,
      );
      try {
        const response = await deps.fetch(track.deliveryUrl);
        if (!response.ok) {
          throw new Error(`Failed to download Jellyfin subtitle (HTTP ${response.status})`);
        }
        const bytes = new Uint8Array(await response.arrayBuffer());
        await deps.writeFile(subtitlePath, bytes);
      } catch (error) {
        deps.removeDir(cacheDir, { recursive: true, force: true });
        throw error;
      }
      return { path: subtitlePath, cleanupDir: cacheDir };
    },
    cleanupCachedSubtitles(dirs: string[]): void {
      for (const dir of dirs) {
        deps.removeDir(dir, { recursive: true, force: true });
      }
    },
  };
}
