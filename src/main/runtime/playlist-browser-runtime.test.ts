import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test, { type TestContext } from 'node:test';

import type { PlaylistBrowserQueueItem } from '../../types';
import {
  appendPlaylistBrowserFileRuntime,
  getPlaylistBrowserSnapshotRuntime,
  movePlaylistBrowserIndexRuntime,
  playPlaylistBrowserIndexRuntime,
  removePlaylistBrowserIndexRuntime,
} from './playlist-browser-runtime';

type FakePlaylistEntry = {
  current?: boolean;
  playing?: boolean;
  filename: string;
  title?: string;
  id?: number;
};

function createTempVideoDir(t: TestContext): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-playlist-browser-'));
  t.after(() => {
    fs.rmSync(dir, { recursive: true, force: true });
  });
  return dir;
}

function createFakeMpvClient(options: {
  currentVideoPath: string;
  playlist: FakePlaylistEntry[];
  connected?: boolean;
}) {
  let playlist = options.playlist.map((item, index) => ({
    id: item.id ?? index + 1,
    current: item.current ?? false,
    playing: item.playing ?? item.current ?? false,
    filename: item.filename,
    title: item.title ?? null,
  }));
  const commands: Array<(string | number)[]> = [];

  const syncFlags = (): void => {
    let playingIndex = playlist.findIndex((item) => item.current || item.playing);
    if (playingIndex < 0 && playlist.length > 0) {
      playingIndex = 0;
    }
    playlist = playlist.map((item, index) => ({
      ...item,
      current: index === playingIndex,
      playing: index === playingIndex,
    }));
  };

  syncFlags();

  return {
    connected: options.connected ?? true,
    currentVideoPath: options.currentVideoPath,
    async requestProperty(name: string): Promise<unknown> {
      if (name === 'playlist') {
        return playlist;
      }
      if (name === 'playlist-playing-pos') {
        return playlist.findIndex((item) => item.current || item.playing);
      }
      if (name === 'path') {
        return this.currentVideoPath;
      }
      throw new Error(`Unexpected property: ${name}`);
    },
    send(payload: { command: unknown[] }): boolean {
      const command = payload.command as (string | number)[];
      commands.push(command);
      const [action, first, second] = command;
      if (action === 'loadfile' && typeof first === 'string' && second === 'append') {
        playlist.push({
          id: playlist.length + 1,
          filename: first,
          title: null,
          current: false,
          playing: false,
        });
        syncFlags();
        return true;
      }
      if (action === 'playlist-play-index' && typeof first === 'number' && playlist[first]) {
        playlist = playlist.map((item, index) => ({
          ...item,
          current: index === first,
          playing: index === first,
        }));
        this.currentVideoPath = playlist[first]!.filename;
        return true;
      }
      if (action === 'playlist-remove' && typeof first === 'number' && playlist[first]) {
        const removingCurrent = playlist[first]!.current || playlist[first]!.playing;
        playlist.splice(first, 1);
        if (removingCurrent) {
          syncFlags();
          this.currentVideoPath =
            playlist.find((item) => item.current || item.playing)?.filename ??
            this.currentVideoPath;
        }
        return true;
      }
      if (
        action === 'playlist-move' &&
        typeof first === 'number' &&
        typeof second === 'number' &&
        playlist[first]
      ) {
        const [moved] = playlist.splice(first, 1);
        playlist.splice(second, 0, moved!);
        syncFlags();
        return true;
      }
      return true;
    },
    getCommands(): Array<(string | number)[]> {
      return commands;
    },
  };
}

function createDeferred<T>(): {
  promise: Promise<T>;
  resolve: (value: T) => void;
} {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((settle) => {
    resolve = settle;
  });
  return { promise, resolve };
}

test('getPlaylistBrowserSnapshotRuntime lists sibling videos in best-effort episode order', async (t) => {
  const dir = createTempVideoDir(t);
  const episode2 = path.join(dir, 'Show - S01E02.mkv');
  const episode1 = path.join(dir, 'Show - S01E01.mkv');
  const special = path.join(dir, 'Show - Special.mp4');
  const ignored = path.join(dir, 'notes.txt');
  fs.writeFileSync(episode2, '');
  fs.writeFileSync(episode1, '');
  fs.writeFileSync(special, '');
  fs.writeFileSync(ignored, '');

  const mpvClient = createFakeMpvClient({
    currentVideoPath: episode2,
    playlist: [
      { filename: episode1, current: false, playing: false, title: 'Episode 1' },
      { filename: episode2, current: true, playing: true, title: 'Episode 2' },
    ],
  });

  const snapshot = await getPlaylistBrowserSnapshotRuntime({
    getMpvClient: () => mpvClient,
  });

  assert.equal(snapshot.directoryAvailable, true);
  assert.equal(snapshot.directoryPath, dir);
  assert.equal(snapshot.currentFilePath, episode2);
  assert.equal(snapshot.playingIndex, 1);
  assert.deepEqual(
    snapshot.directoryItems.map((item) => [item.basename, item.isCurrentFile]),
    [
      ['Show - S01E01.mkv', false],
      ['Show - S01E02.mkv', true],
      ['Show - Special.mp4', false],
    ],
  );
  assert.deepEqual(
    snapshot.playlistItems.map((item) => ({
      index: item.index,
      displayLabel: item.displayLabel,
      current: item.current,
    })),
    [
      { index: 0, displayLabel: 'Episode 1', current: false },
      { index: 1, displayLabel: 'Episode 2', current: true },
    ],
  );
});

test('getPlaylistBrowserSnapshotRuntime clamps stale playing index to the playlist bounds', async (t) => {
  const dir = createTempVideoDir(t);
  const episode1 = path.join(dir, 'Show - S01E01.mkv');
  const episode2 = path.join(dir, 'Show - S01E02.mkv');
  fs.writeFileSync(episode1, '');
  fs.writeFileSync(episode2, '');

  const mpvClient = createFakeMpvClient({
    currentVideoPath: episode1,
    playlist: [
      { filename: episode1, current: true, playing: true, title: 'Episode 1' },
      { filename: episode2, title: 'Episode 2' },
    ],
  });
  const requestProperty = mpvClient.requestProperty.bind(mpvClient);
  mpvClient.requestProperty = async (name: string): Promise<unknown> => {
    if (name === 'playlist-playing-pos') {
      return 99;
    }
    return requestProperty(name);
  };

  const snapshot = await getPlaylistBrowserSnapshotRuntime({
    getMpvClient: () => mpvClient,
  });

  assert.equal(snapshot.playingIndex, 1);
});

test('getPlaylistBrowserSnapshotRuntime degrades directory pane for remote media', async () => {
  const mpvClient = createFakeMpvClient({
    currentVideoPath: 'https://example.com/video.m3u8',
    playlist: [{ filename: 'https://example.com/video.m3u8', current: true }],
  });

  const snapshot = await getPlaylistBrowserSnapshotRuntime({
    getMpvClient: () => mpvClient,
  });

  assert.equal(snapshot.directoryAvailable, false);
  assert.equal(snapshot.directoryItems.length, 0);
  assert.match(snapshot.directoryStatus, /local filesystem/i);
  assert.equal(snapshot.playlistItems.length, 1);
});

test('playlist-browser mutation runtimes mutate queue and return refreshed snapshots', async (t) => {
  const dir = createTempVideoDir(t);
  const episode1 = path.join(dir, 'Show - S01E01.mkv');
  const episode2 = path.join(dir, 'Show - S01E02.mkv');
  const episode3 = path.join(dir, 'Show - S01E03.mkv');
  fs.writeFileSync(episode1, '');
  fs.writeFileSync(episode2, '');
  fs.writeFileSync(episode3, '');

  const mpvClient = createFakeMpvClient({
    currentVideoPath: episode1,
    playlist: [
      { filename: episode1, current: true, title: 'Episode 1' },
      { filename: episode2, title: 'Episode 2' },
    ],
  });

  const scheduled: Array<{ callback: () => void; delayMs: number }> = [];
  const deps = {
    getMpvClient: () => mpvClient,
    schedule: (callback: () => void, delayMs: number) => {
      scheduled.push({ callback, delayMs });
    },
  };

  const appendResult = await appendPlaylistBrowserFileRuntime(deps, episode3);
  assert.equal(appendResult.ok, true);
  assert.deepEqual(mpvClient.getCommands().at(-1), ['loadfile', episode3, 'append']);
  assert.deepEqual(
    appendResult.snapshot?.playlistItems.map((item) => item.path),
    [episode1, episode2, episode3],
  );

  const moveResult = await movePlaylistBrowserIndexRuntime(deps, 2, -1);
  assert.equal(moveResult.ok, true);
  assert.deepEqual(mpvClient.getCommands().at(-1), ['playlist-move', 2, 1]);
  assert.deepEqual(
    moveResult.snapshot?.playlistItems.map((item) => item.path),
    [episode1, episode3, episode2],
  );

  const playResult = await playPlaylistBrowserIndexRuntime(deps, 1);
  assert.equal(playResult.ok, true);
  assert.deepEqual(mpvClient.getCommands().slice(-2), [
    ['set_property', 'sub-auto', 'fuzzy'],
    ['playlist-play-index', 1],
  ]);
  assert.deepEqual(
    scheduled.map((entry) => entry.delayMs),
    [400],
  );
  scheduled[0]?.callback();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.deepEqual(mpvClient.getCommands().slice(-2), [
    ['set_property', 'sid', 'auto'],
    ['set_property', 'secondary-sid', 'auto'],
  ]);
  assert.equal(playResult.snapshot?.playingIndex, 1);

  const removeResult = await removePlaylistBrowserIndexRuntime(deps, 2);
  assert.equal(removeResult.ok, true);
  assert.deepEqual(mpvClient.getCommands().at(-1), ['playlist-remove', 2]);
  assert.deepEqual(
    removeResult.snapshot?.playlistItems.map((item) => item.path),
    [episode1, episode3],
  );
});

test('playlist-browser mutation runtimes report MPV send rejection', async (t) => {
  const dir = createTempVideoDir(t);
  const episode1 = path.join(dir, 'Show - S01E01.mkv');
  const episode2 = path.join(dir, 'Show - S01E02.mkv');
  const episode3 = path.join(dir, 'Show - S01E03.mkv');
  fs.writeFileSync(episode1, '');
  fs.writeFileSync(episode2, '');
  fs.writeFileSync(episode3, '');

  const mpvClient = createFakeMpvClient({
    currentVideoPath: episode1,
    playlist: [
      { filename: episode1, current: true, title: 'Episode 1' },
      { filename: episode2, title: 'Episode 2' },
      { filename: episode3, title: 'Episode 3' },
    ],
  });
  const scheduled: Array<{ callback: () => void; delayMs: number }> = [];
  mpvClient.send = () => false;
  const deps = {
    getMpvClient: () => mpvClient,
    schedule: (callback: () => void, delayMs: number) => {
      scheduled.push({ callback, delayMs });
    },
  };

  const appendResult = await appendPlaylistBrowserFileRuntime(deps, episode3);
  assert.equal(appendResult.ok, false);
  assert.equal(appendResult.snapshot, undefined);

  const playResult = await playPlaylistBrowserIndexRuntime(deps, 1);
  assert.equal(playResult.ok, false);
  assert.equal(playResult.snapshot, undefined);
  assert.deepEqual(scheduled, []);

  const removeResult = await removePlaylistBrowserIndexRuntime(deps, 1);
  assert.equal(removeResult.ok, false);
  assert.equal(removeResult.snapshot, undefined);

  const moveResult = await movePlaylistBrowserIndexRuntime(deps, 1, 1);
  assert.equal(moveResult.ok, false);
  assert.equal(moveResult.snapshot, undefined);
});

test('appendPlaylistBrowserFileRuntime returns an error result when statSync throws', async (t) => {
  const dir = createTempVideoDir(t);
  const episode1 = path.join(dir, 'Show - S01E01.mkv');
  fs.writeFileSync(episode1, '');

  const mutableFs = fs as typeof fs & { statSync: typeof fs.statSync };
  const originalStatSync = mutableFs.statSync;
  mutableFs.statSync = ((targetPath: fs.PathLike) => {
    if (path.resolve(String(targetPath)) === episode1) {
      throw new Error('EACCES');
    }
    return originalStatSync(targetPath);
  }) as typeof fs.statSync;

  try {
    const result = await appendPlaylistBrowserFileRuntime(
      {
        getMpvClient: () =>
          createFakeMpvClient({
            currentVideoPath: episode1,
            playlist: [{ filename: episode1, current: true }],
          }),
      },
      episode1,
    );

    assert.deepEqual(result, {
      ok: false,
      message: 'Playlist browser file is not readable.',
    });
  } finally {
    mutableFs.statSync = originalStatSync;
  }
});

test('movePlaylistBrowserIndexRuntime rejects top and bottom boundary moves', async (t) => {
  const dir = createTempVideoDir(t);
  const episode1 = path.join(dir, 'Show - S01E01.mkv');
  const episode2 = path.join(dir, 'Show - S01E02.mkv');
  fs.writeFileSync(episode1, '');
  fs.writeFileSync(episode2, '');

  const mpvClient = createFakeMpvClient({
    currentVideoPath: episode1,
    playlist: [{ filename: episode1, current: true }, { filename: episode2 }],
  });

  const deps = {
    getMpvClient: () => mpvClient,
  };

  const moveUp = await movePlaylistBrowserIndexRuntime(deps, 0, -1);
  assert.deepEqual(moveUp, {
    ok: false,
    message: 'Playlist item is already at the top.',
  });

  const moveDown = await movePlaylistBrowserIndexRuntime(deps, 1, 1);
  assert.deepEqual(moveDown, {
    ok: false,
    message: 'Playlist item is already at the bottom.',
  });
});

test('getPlaylistBrowserSnapshotRuntime normalizes playlist labels from title then filename', async (t) => {
  const dir = createTempVideoDir(t);
  const episode1 = path.join(dir, 'Show - S01E01.mkv');
  fs.writeFileSync(episode1, '');

  const mpvClient = createFakeMpvClient({
    currentVideoPath: episode1,
    playlist: [{ filename: episode1, current: true, title: '' }],
  });

  const snapshot = await getPlaylistBrowserSnapshotRuntime({
    getMpvClient: () => mpvClient,
  });

  const item = snapshot.playlistItems[0] as PlaylistBrowserQueueItem;
  assert.equal(item.displayLabel, 'Show - S01E01.mkv');
  assert.equal(item.path, episode1);
});

test('playPlaylistBrowserIndexRuntime skips local subtitle reset for remote playlist entries', async () => {
  const scheduled: Array<{ callback: () => void; delayMs: number }> = [];
  const mpvClient = createFakeMpvClient({
    currentVideoPath: 'https://example.com/video-1.m3u8',
    playlist: [
      { filename: 'https://example.com/video-1.m3u8', current: true, title: 'Episode 1' },
      { filename: 'https://example.com/video-2.m3u8', title: 'Episode 2' },
    ],
  });

  const result = await playPlaylistBrowserIndexRuntime(
    {
      getMpvClient: () => mpvClient,
      schedule: (callback, delayMs) => {
        scheduled.push({ callback, delayMs });
      },
    },
    1,
  );

  assert.equal(result.ok, true);
  assert.deepEqual(mpvClient.getCommands().slice(-1), [['playlist-play-index', 1]]);
  assert.equal(scheduled.length, 0);
});

test('playPlaylistBrowserIndexRuntime ignores superseded local subtitle rearm callbacks', async (t) => {
  const dir = createTempVideoDir(t);
  const episode1 = path.join(dir, 'Show - S01E01.mkv');
  const episode2 = path.join(dir, 'Show - S01E02.mkv');
  const episode3 = path.join(dir, 'Show - S01E03.mkv');
  fs.writeFileSync(episode1, '');
  fs.writeFileSync(episode2, '');
  fs.writeFileSync(episode3, '');

  const scheduled: Array<() => void> = [];
  const mpvClient = createFakeMpvClient({
    currentVideoPath: episode1,
    playlist: [
      { filename: episode1, current: true, title: 'Episode 1' },
      { filename: episode2, title: 'Episode 2' },
      { filename: episode3, title: 'Episode 3' },
    ],
  });

  const deps = {
    getMpvClient: () => mpvClient,
    schedule: (callback: () => void) => {
      scheduled.push(callback);
    },
  };

  const firstPlay = await playPlaylistBrowserIndexRuntime(deps, 1);
  const secondPlay = await playPlaylistBrowserIndexRuntime(deps, 2);

  assert.equal(firstPlay.ok, true);
  assert.equal(secondPlay.ok, true);
  assert.equal(scheduled.length, 2);

  scheduled[0]?.();
  scheduled[1]?.();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(mpvClient.getCommands().slice(-6), [
    ['set_property', 'sub-auto', 'fuzzy'],
    ['playlist-play-index', 1],
    ['set_property', 'sub-auto', 'fuzzy'],
    ['playlist-play-index', 2],
    ['set_property', 'sid', 'auto'],
    ['set_property', 'secondary-sid', 'auto'],
  ]);
});

test('playPlaylistBrowserIndexRuntime aborts stale async subtitle rearm work', async (t) => {
  const dir = createTempVideoDir(t);
  const episode1 = path.join(dir, 'Show - S01E01.mkv');
  const episode2 = path.join(dir, 'Show - S01E02.mkv');
  fs.writeFileSync(episode1, '');
  fs.writeFileSync(episode2, '');

  const firstTrackList = createDeferred<unknown>();
  const secondTrackList = createDeferred<unknown>();
  let trackListRequestCount = 0;
  const mpvClient = createFakeMpvClient({
    currentVideoPath: episode1,
    playlist: [
      { filename: episode1, current: true, title: 'Episode 1' },
      { filename: episode2, title: 'Episode 2' },
    ],
  });
  const requestProperty = mpvClient.requestProperty.bind(mpvClient);
  mpvClient.requestProperty = async (name: string): Promise<unknown> => {
    if (name === 'track-list') {
      trackListRequestCount += 1;
      return trackListRequestCount === 1 ? firstTrackList.promise : secondTrackList.promise;
    }
    return requestProperty(name);
  };

  const scheduled: Array<() => void> = [];
  const deps = {
    getMpvClient: () => mpvClient,
    schedule: (callback: () => void) => {
      scheduled.push(callback);
    },
  };

  const firstPlay = await playPlaylistBrowserIndexRuntime(deps, 1);
  assert.equal(firstPlay.ok, true);
  scheduled[0]?.();

  const secondPlay = await playPlaylistBrowserIndexRuntime(deps, 1);
  assert.equal(secondPlay.ok, true);
  scheduled[1]?.();

  secondTrackList.resolve([
    { type: 'sub', id: 21, lang: 'ja', title: 'Japanese', external: false, selected: true },
    { type: 'sub', id: 22, lang: 'en', title: 'English', external: false },
  ]);
  await new Promise((resolve) => setTimeout(resolve, 0));

  firstTrackList.resolve([
    { type: 'sub', id: 11, lang: 'ja', title: 'Japanese', external: false, selected: true },
    { type: 'sub', id: 12, lang: 'en', title: 'English', external: false },
  ]);
  await new Promise((resolve) => setTimeout(resolve, 0));

  const subtitleCommands = mpvClient
    .getCommands()
    .filter(
      (command) =>
        command[0] === 'set_property' && (command[1] === 'sid' || command[1] === 'secondary-sid'),
    );

  assert.deepEqual(subtitleCommands, [
    ['set_property', 'sid', 21],
    ['set_property', 'secondary-sid', 22],
  ]);
});

test('playlist-browser playback reapplies configured preferred subtitle tracks when track metadata is available', async (t) => {
  const dir = createTempVideoDir(t);
  const episode1 = path.join(dir, 'Show - S01E01.mkv');
  const episode2 = path.join(dir, 'Show - S01E02.mkv');
  fs.writeFileSync(episode1, '');
  fs.writeFileSync(episode2, '');

  const mpvClient = createFakeMpvClient({
    currentVideoPath: episode1,
    playlist: [
      { filename: episode1, current: true, title: 'Episode 1' },
      { filename: episode2, title: 'Episode 2' },
    ],
  });
  const requestProperty = mpvClient.requestProperty.bind(mpvClient);
  mpvClient.requestProperty = async (name: string): Promise<unknown> => {
    if (name === 'track-list') {
      return [
        { type: 'sub', id: 1, lang: 'pt', title: '[Infinite]', external: false, selected: true },
        { type: 'sub', id: 3, lang: 'en', title: 'English', external: false },
        { type: 'sub', id: 11, lang: 'en', title: 'en.srt', external: true },
        { type: 'sub', id: 12, lang: 'ja', title: 'ja.srt', external: true },
      ];
    }
    return requestProperty(name);
  };

  const scheduled: Array<() => void> = [];
  const deps = {
    getMpvClient: () => mpvClient,
    getPrimarySubtitleLanguages: () => [],
    getSecondarySubtitleLanguages: () => [],
    schedule: (callback: () => void) => {
      scheduled.push(callback);
    },
  };

  const result = await playPlaylistBrowserIndexRuntime(deps, 1);
  assert.equal(result.ok, true);

  scheduled[0]?.();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(mpvClient.getCommands().slice(-2), [
    ['set_property', 'sid', 12],
    ['set_property', 'secondary-sid', 11],
  ]);
});
