import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { isValidSyncHost } from '../../shared/sync/sync-hosts-store';
import type { SyncProgressEvent } from '../../shared/sync/sync-events';
import type {
  SyncUiCheckResult,
  SyncUiRunKind,
  SyncUiRunState,
  SyncUiStartResult,
} from '../../types/sync-ui';
import {
  runSyncLauncher,
  type SyncLauncherRunHandle,
  type SyncLauncherRunResult,
} from './sync-launcher-client';

interface ActiveRun {
  id: number;
  kind: SyncUiRunKind;
  host: string | null;
  handle: SyncLauncherRunHandle;
  resultSeen: boolean;
  completion?: Promise<void>;
}

interface SyncUiRunLifecycleDeps {
  runLauncher: typeof runSyncLauncher;
  resolveLauncherCommand: () => string[];
  log?: (message: string) => void;
  notify?: (payload: { title: string; body: string; variant: 'success' | 'error' }) => void;
  sendToWindow: (channel: string, payload?: unknown) => void;
  broadcastStateChanged: () => void;
}

export function createSyncUiRunLifecycle(deps: SyncUiRunLifecycleDeps) {
  let runCounter = 0;
  let currentRun: ActiveRun | null = null;

  function launchRun(
    kind: SyncUiRunKind,
    host: string | null,
    args: string[],
    options: {
      notify?: boolean;
      timeoutMs?: number;
      emitProgress?: boolean;
      onEvent?: (event: SyncProgressEvent) => void;
    } = {},
  ): { start: SyncUiStartResult; done: Promise<SyncLauncherRunResult> | null } {
    if (currentRun) {
      return {
        start: { started: false, runId: null, reason: 'A sync operation is already running.' },
        done: null,
      };
    }
    const emitProgress = options.emitProgress ?? true;
    runCounter += 1;
    const runId = runCounter;
    const run: ActiveRun = {
      id: runId,
      kind,
      host,
      resultSeen: false,
      handle: deps.runLauncher({
        command: deps.resolveLauncherCommand(),
        args,
        timeoutMs: options.timeoutMs,
        onEvent: (event: SyncProgressEvent) => {
          if (event.type === 'result') run.resultSeen = true;
          options.onEvent?.(event);
          if (emitProgress) {
            deps.sendToWindow(IPC_CHANNELS.event.syncUiProgress, { runId, kind, host, event });
          }
        },
        onStderr: (text) => deps.log?.(`[sync-ui] ${text.trimEnd()}`),
      }),
    };
    currentRun = run;
    run.completion = run.handle.done
      .then((result: SyncLauncherRunResult) => {
        if (currentRun?.id === runId) currentRun = null;
        if (emitProgress && !run.resultSeen) {
          deps.sendToWindow(IPC_CHANNELS.event.syncUiProgress, {
            runId,
            kind,
            host,
            event: { type: 'result', ok: result.ok, error: result.error },
          });
        }
        deps.broadcastStateChanged();
        if (options.notify && deps.notify) {
          const target = host ?? 'local database';
          deps.notify(
            result.ok
              ? { title: 'Sync complete', body: `Synced with ${target}`, variant: 'success' }
              : {
                  title: 'Sync failed',
                  body: `${target}: ${result.error ?? 'unknown error'}`,
                  variant: 'error',
                },
          );
        }
      })
      .catch((error) => {
        deps.log?.(
          `[sync-ui] Post-run cleanup failed: ${error instanceof Error ? error.message : String(error)}`,
        );
      });
    return { start: { started: true, runId, reason: null }, done: run.handle.done };
  }

  function startRun(
    kind: SyncUiRunKind,
    host: string | null,
    args: string[],
    options: { notify?: boolean } = {},
  ): SyncUiStartResult {
    return launchRun(kind, host, args, options).start;
  }

  function checkHost(host: string): Promise<SyncUiCheckResult> {
    const trimmed = host.trim();
    const failed = (error: string): SyncUiCheckResult => ({
      host: trimmed,
      sshOk: false,
      remoteCommand: null,
      remoteVersion: null,
      ok: false,
      error,
    });
    if (!isValidSyncHost(trimmed)) {
      return Promise.resolve(failed(`Invalid sync host: ${host}`));
    }
    let checkResult: SyncUiCheckResult | null = null;
    const { start, done } = launchRun('check', trimmed, ['sync', trimmed, '--check', '--json'], {
      timeoutMs: 30_000,
      emitProgress: false,
      onEvent: (event) => {
        if (event.type === 'check-result') {
          checkResult = {
            host: event.host,
            sshOk: event.sshOk,
            remoteCommand: event.remoteCommand,
            remoteVersion: event.remoteVersion,
            ok: event.ok,
            error: event.error,
          };
        }
      },
    });
    if (!start.started || !done) {
      return Promise.resolve(failed(start.reason ?? 'A sync operation is already running.'));
    }
    return done.then((result) => checkResult ?? failed(result.error ?? 'Connection check failed.'));
  }

  function cancelRun(): boolean {
    if (!currentRun) return false;
    currentRun.handle.cancel();
    return true;
  }

  async function shutdown(): Promise<void> {
    const run = currentRun;
    if (!run) return;
    run.handle.cancel();
    await (run.completion ?? run.handle.done.then(() => undefined));
  }

  function runState(): SyncUiRunState {
    return currentRun
      ? { running: true, runId: currentRun.id, kind: currentRun.kind, host: currentRun.host }
      : { running: false, runId: null, kind: null, host: null };
  }

  return {
    startRun,
    checkHost,
    cancelRun,
    shutdown,
    runState,
    isRunning: () => currentRun !== null,
  };
}
