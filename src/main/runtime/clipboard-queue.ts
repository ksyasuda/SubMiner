import fs from 'node:fs';
import path from 'node:path';
import { parseClipboardVideoPath } from '../../core/services';

type MpvClientLike = {
  connected: boolean;
};

export type AppendClipboardVideoToQueueRuntimeDeps = {
  getMpvClient: () => MpvClientLike | null;
  readClipboardText: () => string;
  showMpvOsd: (text: string) => void;
  sendMpvCommand: (command: (string | number)[]) => void;
};

export function appendClipboardVideoToQueueRuntime(deps: AppendClipboardVideoToQueueRuntimeDeps): {
  ok: boolean;
  message: string;
} {
  const mpvClient = deps.getMpvClient();
  if (!mpvClient || !mpvClient.connected) {
    return { ok: false, message: 'MPV is not connected.' };
  }

  const clipboardText = deps.readClipboardText();
  const parsedPath = parseClipboardVideoPath(clipboardText);
  if (!parsedPath) {
    deps.showMpvOsd('Clipboard does not contain a supported video path.');
    return { ok: false, message: 'Clipboard does not contain a supported video path.' };
  }

  const resolvedPath = path.resolve(parsedPath);
  if (!fs.existsSync(resolvedPath) || !fs.statSync(resolvedPath).isFile()) {
    deps.showMpvOsd('Clipboard path is not a readable file.');
    return { ok: false, message: 'Clipboard path is not a readable file.' };
  }

  deps.sendMpvCommand(['loadfile', resolvedPath, 'append']);
  deps.showMpvOsd(`Queued from clipboard: ${path.basename(resolvedPath)}`);
  return { ok: true, message: `Queued ${resolvedPath}` };
}
