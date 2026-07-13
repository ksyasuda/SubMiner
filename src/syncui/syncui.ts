import type {
  SyncDirection,
  SyncHostEntry,
  SyncUiAPI,
  SyncUiCheckResult,
  SyncUiProgressPayload,
  SyncUiSnapshot,
  SyncUiStartResult,
} from '../types/sync-ui';
import { formatBytes, formatRelativeTime, summarizeMergeCounts } from './syncui-format';

declare global {
  interface Window {
    syncUiAPI: SyncUiAPI;
  }
}

const api = window.syncUiAPI;

function el<T extends HTMLElement>(id: string): T {
  const node = document.getElementById(id);
  if (!node) throw new Error(`Missing element #${id}`);
  return node as T;
}

const runPill = el<HTMLDivElement>('runPill');
const dbPathEl = el<HTMLDivElement>('dbPath');
const hostList = el<HTMLDivElement>('hostList');
const hostListEmpty = el<HTMLDivElement>('hostListEmpty');
const autoIntervalInput = el<HTMLInputElement>('autoIntervalInput');
const newHostInput = el<HTMLInputElement>('newHostInput');
const newLabelInput = el<HTMLInputElement>('newLabelInput');
const testHostButton = el<HTMLButtonElement>('testHostButton');
const addHostButton = el<HTMLButtonElement>('addHostButton');
const checkResultEl = el<HTMLDivElement>('checkResult');
const activityPanel = el<HTMLElement>('activityPanel');
const cancelButton = el<HTMLButtonElement>('cancelButton');
const stageList = el<HTMLDivElement>('stageList');
const remoteOutput = el<HTMLPreElement>('remoteOutput');
const resultCard = el<HTMLDivElement>('resultCard');
const snapshotsDirEl = el<HTMLSpanElement>('snapshotsDir');
const snapshotList = el<HTMLDivElement>('snapshotList');
const snapshotListEmpty = el<HTMLDivElement>('snapshotListEmpty');
const createSnapshotButton = el<HTMLButtonElement>('createSnapshotButton');
const mergeFileButton = el<HTMLButtonElement>('mergeFileButton');

let snapshot: SyncUiSnapshot | null = null;
let activeRunId: number | null = null;
let retryWithForce: (() => void) | null = null;

const DIRECTIONS: Array<{ id: SyncDirection; label: string; title: string }> = [
  { id: 'both', label: '⇄ Two-way', title: 'Merge both databases into each other' },
  { id: 'push', label: '⇡ Push', title: 'Only merge this machine into the device' },
  { id: 'pull', label: '⇣ Pull', title: 'Only merge the device into this machine' },
];

function describeRun(kind: string, host: string | null): string {
  if (kind === 'host-sync') return `Syncing with ${host ?? 'device'}`;
  if (kind === 'merge') return 'Merging snapshot';
  if (kind === 'snapshot') return 'Creating snapshot';
  return 'Working';
}

function setRunPill(): void {
  const running = snapshot?.run.running ?? false;
  runPill.classList.toggle('running', running);
  runPill.classList.toggle('idle', !running);
  runPill.textContent = running
    ? describeRun(snapshot!.run.kind ?? '', snapshot!.run.host)
    : 'Idle';
  cancelButton.classList.toggle('hidden', !running);
}

function clearActivity(): void {
  stageList.replaceChildren();
  remoteOutput.textContent = '';
  remoteOutput.classList.add('hidden');
  resultCard.replaceChildren();
  resultCard.classList.add('hidden');
}

function markStageStates(state: 'done' | 'failed'): void {
  for (const item of stageList.querySelectorAll('.stage-item.active')) {
    item.classList.remove('active');
    item.classList.add(state);
  }
}

function addStage(message: string): void {
  markStageStates('done');
  const item = document.createElement('div');
  item.className = 'stage-item active';
  const icon = document.createElement('span');
  icon.className = 'stage-icon';
  const text = document.createElement('span');
  text.textContent = message;
  item.append(icon, text);
  stageList.appendChild(item);
}

function showResult(ok: boolean, error: string | null): void {
  markStageStates(ok ? 'done' : 'failed');
  resultCard.classList.remove('hidden');
  resultCard.classList.toggle('error', !ok);
  const title = document.createElement('div');
  title.className = 'result-title';
  title.textContent = ok ? 'Sync complete' : 'Sync failed';
  resultCard.prepend(title);
  if (!ok && error) {
    const detail = document.createElement('div');
    detail.className = 'result-error-detail';
    detail.textContent = error;
    resultCard.appendChild(detail);
    if (retryWithForce && /--force/.test(error)) {
      const retry = document.createElement('button');
      retry.className = 'ghost-button';
      retry.style.marginTop = '10px';
      retry.type = 'button';
      retry.textContent = 'Retry with --force (skip the running-app check)';
      retry.addEventListener('click', () => {
        const action = retryWithForce;
        retryWithForce = null;
        action?.();
      });
      resultCard.appendChild(retry);
    }
  }
}

function appendMergeSummary(
  target: 'local' | 'remote',
  lines: ReturnType<typeof summarizeMergeCounts>,
): void {
  resultCard.classList.remove('hidden');
  const heading = document.createElement('div');
  heading.className = 'result-title';
  heading.textContent = target === 'local' ? 'Merged into this machine' : 'Merged into the device';
  const grid = document.createElement('div');
  grid.className = 'count-grid';
  for (const line of lines) {
    const cell = document.createElement('div');
    cell.className = 'count-cell';
    const label = document.createElement('span');
    label.textContent = line.label;
    const value = document.createElement('b');
    value.textContent = String(line.value);
    cell.append(label, value);
    grid.appendChild(cell);
  }
  resultCard.append(heading, grid);
}

function handleProgress(payload: SyncUiProgressPayload): void {
  if (payload.runId !== activeRunId) {
    activeRunId = payload.runId;
    clearActivity();
    activityPanel.classList.remove('hidden');
  }
  const event = payload.event;
  if (event.type === 'stage') {
    addStage(event.message);
  } else if (event.type === 'snapshot-created') {
    addStage(`Snapshot written to ${event.path}`);
  } else if (event.type === 'remote-output') {
    remoteOutput.classList.remove('hidden');
    remoteOutput.textContent += event.text;
  } else if (event.type === 'merge-summary') {
    appendMergeSummary(event.target, summarizeMergeCounts(event.summary));
  } else if (event.type === 'result') {
    // The run's completion handler broadcasts state-changed, which refreshes.
    showResult(event.ok, event.error);
  }
}

async function startRun(action: () => Promise<SyncUiStartResult>): Promise<void> {
  const result = await action();
  if (!result.started) {
    activityPanel.classList.remove('hidden');
    clearActivity();
    activeRunId = null;
    showResult(false, result.reason ?? 'Could not start the sync.');
    return;
  }
  await refresh();
}

function runHostSync(entry: SyncHostEntry): Promise<SyncUiStartResult> {
  retryWithForce = () => {
    void startRun(() => api.runSync({ host: entry.host, direction: entry.direction, force: true }));
  };
  return api.runSync({ host: entry.host, direction: entry.direction });
}

function renderHostCard(entry: SyncHostEntry): HTMLDivElement {
  const running = snapshot?.run.running ?? false;
  const syncingThis = running && snapshot?.run.host === entry.host;
  const card = document.createElement('div');
  card.className = `host-card${syncingThis ? ' syncing' : ''}`;

  const top = document.createElement('div');
  top.className = 'host-top';
  const id = document.createElement('div');
  id.className = 'host-id';
  const name = document.createElement('span');
  name.className = 'host-name';
  name.textContent = entry.host;
  id.appendChild(name);
  if (entry.label) {
    const label = document.createElement('span');
    label.className = 'host-label';
    label.textContent = entry.label;
    id.appendChild(label);
  }
  const status = document.createElement('div');
  status.className = 'host-status';
  const dot = document.createElement('span');
  dot.className = `status-dot${entry.lastSyncStatus ? ` ${entry.lastSyncStatus}` : ''}`;
  const statusText = document.createElement('span');
  statusText.className = 'detail';
  const when = formatRelativeTime(entry.lastSyncAtMs, Date.now());
  statusText.textContent =
    entry.lastSyncStatus === null
      ? 'never synced'
      : `${when}${entry.lastSyncDetail ? ` · ${entry.lastSyncDetail}` : ''}`;
  statusText.title = entry.lastSyncDetail ?? '';
  status.append(dot, statusText);
  top.append(id, status);

  const controls = document.createElement('div');
  controls.className = 'host-controls';

  const toggle = document.createElement('div');
  toggle.className = 'direction-toggle';
  toggle.setAttribute('role', 'group');
  toggle.setAttribute('aria-label', `Sync direction for ${entry.host}`);
  for (const direction of DIRECTIONS) {
    const button = document.createElement('button');
    button.type = 'button';
    button.textContent = direction.label;
    button.title = direction.title;
    button.classList.toggle('active', entry.direction === direction.id);
    button.addEventListener('click', () => {
      // Every mutation triggers a state-changed broadcast, which refreshes.
      void api.saveHost({ host: entry.host, direction: direction.id });
    });
    toggle.appendChild(button);
  }

  const auto = document.createElement('label');
  auto.className = 'auto-toggle';
  const autoBox = document.createElement('input');
  autoBox.type = 'checkbox';
  autoBox.checked = entry.autoSync;
  autoBox.addEventListener('change', () => {
    void api.saveHost({ host: entry.host, autoSync: autoBox.checked });
  });
  const autoText = document.createElement('span');
  autoText.textContent = 'Auto-sync';
  auto.append(autoBox, autoText);

  const actions = document.createElement('div');
  actions.className = 'host-actions';
  const syncButton = document.createElement('button');
  syncButton.type = 'button';
  syncButton.className = 'primary-button';
  syncButton.textContent = syncingThis ? 'Syncing…' : 'Sync now';
  syncButton.disabled = running;
  syncButton.addEventListener('click', () => {
    void startRun(() => runHostSync(entry));
  });
  const testButton = document.createElement('button');
  testButton.type = 'button';
  testButton.className = 'mini-button';
  testButton.textContent = 'Test';
  testButton.addEventListener('click', () => {
    testButton.disabled = true;
    testButton.textContent = 'Testing…';
    void api
      .checkHost(entry.host)
      .then((result) => renderCheckResult(result))
      .finally(() => {
        testButton.disabled = false;
        testButton.textContent = 'Test';
      });
  });
  const removeButton = document.createElement('button');
  removeButton.type = 'button';
  removeButton.className = 'mini-button danger';
  removeButton.textContent = 'Remove';
  removeButton.addEventListener('click', () => {
    if (!window.confirm(`Remove ${entry.host} from saved devices?`)) return;
    void api.removeHost(entry.host);
  });
  actions.append(syncButton, testButton, removeButton);

  controls.append(toggle, auto, actions);
  card.append(top, controls);
  return card;
}

function renderCheckResult(result: SyncUiCheckResult): void {
  checkResultEl.classList.remove('hidden', 'ok', 'fail');
  checkResultEl.classList.add(result.ok ? 'ok' : 'fail');
  checkResultEl.replaceChildren();
  const headline = document.createElement('div');
  headline.textContent = result.ok
    ? `✓ ${result.host} is ready to sync`
    : `✕ ${result.host} is not ready`;
  checkResultEl.appendChild(headline);
  const sub = document.createElement('div');
  sub.className = 'sub';
  if (result.ok) {
    sub.textContent = `SSH ok · remote launcher: ${result.remoteCommand ?? '?'}${
      result.remoteVersion ? ` (${result.remoteVersion})` : ''
    }`;
  } else {
    sub.textContent = result.error ?? (result.sshOk ? 'Remote launcher missing.' : 'SSH failed.');
  }
  checkResultEl.appendChild(sub);
}

function renderHosts(): void {
  const hosts = snapshot?.hosts.hosts ?? [];
  hostList.replaceChildren(...hosts.map(renderHostCard));
  hostListEmpty.classList.toggle('hidden', hosts.length > 0);
}

function renderSnapshots(): void {
  const files = snapshot?.snapshots ?? [];
  snapshotsDirEl.textContent = snapshot?.snapshotsDir ?? '';
  snapshotList.replaceChildren(
    ...files.map((file) => {
      const row = document.createElement('div');
      row.className = 'snapshot-row';
      const name = document.createElement('span');
      name.className = 'snapshot-name';
      name.textContent = file.name;
      name.title = file.path;
      const meta = document.createElement('span');
      meta.className = 'snapshot-meta';
      meta.textContent = `${formatBytes(file.sizeBytes)} · ${new Date(
        file.modifiedAtMs,
      ).toLocaleString()}`;
      const actions = document.createElement('div');
      actions.className = 'snapshot-actions';
      const merge = document.createElement('button');
      merge.type = 'button';
      merge.className = 'mini-button';
      merge.textContent = 'Merge';
      merge.title = 'Merge this snapshot into the local database';
      merge.disabled = snapshot?.run.running ?? false;
      merge.addEventListener('click', () => {
        retryWithForce = () => {
          void startRun(() => api.mergeSnapshotFile(file.path, true));
        };
        void startRun(() => api.mergeSnapshotFile(file.path));
      });
      const reveal = document.createElement('button');
      reveal.type = 'button';
      reveal.className = 'mini-button';
      reveal.textContent = 'Reveal';
      reveal.addEventListener('click', () => {
        void api.revealSnapshot(file.path);
      });
      const remove = document.createElement('button');
      remove.type = 'button';
      remove.className = 'mini-button danger';
      remove.textContent = 'Delete';
      remove.addEventListener('click', () => {
        if (!window.confirm(`Delete snapshot ${file.name}?`)) return;
        void api.deleteSnapshot(file.path);
      });
      actions.append(merge, reveal, remove);
      row.append(name, meta, actions);
      return row;
    }),
  );
  snapshotListEmpty.classList.toggle('hidden', files.length > 0);
}

async function refresh(): Promise<void> {
  snapshot = await api.getSnapshot();
  dbPathEl.textContent = snapshot.dbPath;
  dbPathEl.title = snapshot.dbPath;
  if (document.activeElement !== autoIntervalInput) {
    autoIntervalInput.value = String(snapshot.hosts.autoSyncIntervalMinutes);
  }
  setRunPill();
  renderHosts();
  renderSnapshots();
  createSnapshotButton.disabled = snapshot.run.running;
  mergeFileButton.disabled = snapshot.run.running;
}

function refreshFromState(): void {
  void refresh();
}

addHostButton.addEventListener('click', () => {
  const host = newHostInput.value.trim();
  if (!host) {
    newHostInput.focus();
    return;
  }
  const label = newLabelInput.value.trim();
  void api
    .saveHost({ host, label: label || null })
    .then(() => {
      newHostInput.value = '';
      newLabelInput.value = '';
      checkResultEl.classList.add('hidden');
    })
    .catch((error: unknown) => {
      renderCheckResult({
        host,
        sshOk: false,
        remoteCommand: null,
        remoteVersion: null,
        ok: false,
        error: error instanceof Error ? error.message : String(error),
      });
    });
});

newHostInput.addEventListener('keydown', (event) => {
  if (event.key === 'Enter') addHostButton.click();
});

testHostButton.addEventListener('click', () => {
  const host = newHostInput.value.trim();
  if (!host) {
    newHostInput.focus();
    return;
  }
  testHostButton.disabled = true;
  testHostButton.textContent = 'Testing…';
  void api
    .checkHost(host)
    .then((result) => renderCheckResult(result))
    .finally(() => {
      testHostButton.disabled = false;
      testHostButton.textContent = 'Test connection';
    });
});

autoIntervalInput.addEventListener('change', () => {
  const minutes = Number(autoIntervalInput.value);
  if (!Number.isFinite(minutes) || minutes < 1) return;
  void api.setAutoSyncInterval(Math.floor(minutes));
});

cancelButton.addEventListener('click', () => {
  void api.cancelRun();
});

createSnapshotButton.addEventListener('click', () => {
  retryWithForce = null;
  void startRun(() => api.createSnapshot());
});

mergeFileButton.addEventListener('click', () => {
  void api.pickSnapshotFile().then((path) => {
    if (!path) return;
    retryWithForce = () => {
      void startRun(() => api.mergeSnapshotFile(path, true));
    };
    void startRun(() => api.mergeSnapshotFile(path));
  });
});

api.onProgress(handleProgress);
api.onStateChanged(refreshFromState);
void refresh();
