import assert from 'node:assert/strict';
import test from 'node:test';
import type { BrowserWindow, IpcMainInvokeEvent, WebContents } from 'electron';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { createAppState } from '../state';
import {
  createSubtitleGenerationAssembly,
  type SubtitleGenerationAssemblyDeps,
} from './subtitle-generation';

type Handler = (event: IpcMainInvokeEvent, payload: unknown) => unknown;

function lazy(name: string) {
  return (): never => {
    throw new Error(`${name} called unexpectedly`);
  };
}

function makeDeps(overrides: Partial<SubtitleGenerationAssemblyDeps> = {}) {
  const thrower = <T extends object>(name: string): T =>
    new Proxy({}, { get: (_target, key: string) => lazy(`${name}.${key}`) }) as T;
  const deps: SubtitleGenerationAssemblyDeps = {
    appState: createAppState({ mpvSocketPath: '/tmp/subminer-subgen.sock', texthookerPort: 0 }),
    getConfig: lazy('getConfig'),
    getModelDirectory: lazy('getModelDirectory'),
    overlayManager: thrower('overlayManager'),
    overlayModals: thrower('overlayModals'),
    youtube: {
      whisperSubtitles: thrower('youtube.whisperSubtitles'),
      flow: thrower('youtube.flow'),
      playback: thrower('youtube.playback'),
    },
    jellyfin: thrower('jellyfin'),
    ipcMain: thrower('ipcMain'),
    ...overrides,
  };
  return deps;
}

test('createSubtitleGenerationAssembly does not call its deps while constructing', () => {
  const subtitleGeneration = createSubtitleGenerationAssembly(makeDeps());
  assert.equal(typeof subtitleGeneration.registerIpc, 'function');
});

test('subtitle generation IPC answers the overlay windows and refuses other senders', () => {
  const overlayContents = {} as WebContents;
  const overlayWindow = {
    isDestroyed: () => false,
    webContents: overlayContents,
  } as unknown as BrowserWindow;
  const handlers = new Map<string, Handler>();
  const subtitleGeneration = createSubtitleGenerationAssembly(
    makeDeps({
      overlayManager: { getMainWindow: () => overlayWindow, getModalWindow: () => null },
      ipcMain: {
        handle: (channel: string, handler: Handler) => {
          handlers.set(channel, handler);
        },
      },
    }),
  );
  subtitleGeneration.registerIpc();
  const cancel = handlers.get(IPC_CHANNELS.request.cancelSubtitleGeneration);
  assert.ok(cancel);

  assert.throws(
    () => cancel({ sender: {} as WebContents } as IpcMainInvokeEvent, undefined),
    /only available from the overlay/,
  );
  assert.doesNotThrow(() => cancel({ sender: overlayContents } as IpcMainInvokeEvent, undefined));
});
