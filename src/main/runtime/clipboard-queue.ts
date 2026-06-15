import fs from 'node:fs';
import path from 'node:path';
import { parseClipboardVideoPath } from '../../core/services/overlay-drop';
import { i18n } from '../../i18n/index.js';

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
    deps.showMpvOsd(i18n.t('osd.clipboardNoVideo'));
    return { ok: false, message: i18n.t('osd.clipboardNoVideo') };
  }

  const resolvedPath = path.resolve(parsedPath);
  if (!fs.existsSync(resolvedPath) || !fs.statSync(resolvedPath).isFile()) {
    deps.showMpvOsd(i18n.t('osd.clipboardNotFile'));
    return { ok: false, message: i18n.t('osd.clipboardNotFile') };
  }

  deps.sendMpvCommand(['loadfile', resolvedPath, 'append']);
  deps.showMpvOsd(`Queued from clipboard: ${path.basename(resolvedPath)}`);
  return { ok: true, message: `Queued ${resolvedPath}` };
}
