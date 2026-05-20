import { contextBridge, ipcRenderer } from 'electron';

const JELLYFIN_SETUP_SUBMIT_CHANNEL = 'jellyfin:setup-submit';

type JellyfinSetupAction = 'login' | 'logout' | 'done';

type JellyfinSetupSubmission = {
  action?: unknown;
  server?: unknown;
  username?: unknown;
  password?: unknown;
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function normalizeAction(value: unknown): JellyfinSetupAction {
  return value === 'logout' || value === 'done' ? value : 'login';
}

function normalizeString(value: unknown): string {
  return typeof value === 'string' ? value : '';
}

function normalizeSubmission(value: unknown): Required<JellyfinSetupSubmission> {
  const record = isRecord(value) ? value : {};
  return {
    action: normalizeAction(record.action),
    server: normalizeString(record.server),
    username: normalizeString(record.username),
    password: normalizeString(record.password),
  };
}

contextBridge.exposeInMainWorld('subminerJellyfinSetup', {
  submit: (submission: JellyfinSetupSubmission): Promise<unknown> =>
    ipcRenderer.invoke(JELLYFIN_SETUP_SUBMIT_CHANNEL, normalizeSubmission(submission)),
});
