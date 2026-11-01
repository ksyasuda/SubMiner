import type { PluginMediaSnapshot } from './types';

/** Keeps plugin snapshots coherent without asynchronous property reads crossing file changes. */
export function createPluginMediaState(onChange: (media: PluginMediaSnapshot | null) => void) {
  let nextId = 0;
  let current: PluginMediaSnapshot | null = null;
  let waitingForLoad = false;

  function publish(media: PluginMediaSnapshot | null): void {
    current = media && Object.freeze(media);
    onChange(current);
  }

  function start(path: string, title: string | null = null): void {
    publish({ playbackId: ++nextId, path, title });
  }

  return {
    current: () => current,
    reset(path = '', title: string | null = null): void {
      waitingForLoad = false;
      if (path) start(path, title);
      else if (current) publish(null);
    },
    pathChanged(path: string): void {
      if (!path) {
        this.reset();
      } else if (path !== current?.path) {
        waitingForLoad = true;
        start(path);
      }
    },
    titleChanged(title: string | null): void {
      if (current && title !== current.title) publish({ ...current, title });
    },
    fileLoaded(path: string, title: string | null): void {
      if (!path) return;
      if (waitingForLoad && current?.path === path) {
        this.titleChanged(title);
      } else {
        start(path, title);
      }
      waitingForLoad = false;
    },
  };
}
