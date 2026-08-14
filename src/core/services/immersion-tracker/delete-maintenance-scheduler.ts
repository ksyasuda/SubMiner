import type { DeleteMaintenanceOperation, DeleteMaintenanceTask } from './delete-maintenance';

type ResolveDeleteMaintenanceOperation = () =>
  | DeleteMaintenanceOperation
  | null
  | Promise<DeleteMaintenanceOperation | null>;

interface PendingDeleteMaintenanceRequest {
  resolveTask: ResolveDeleteMaintenanceOperation;
  resolve: () => void;
  reject: (error: unknown) => void;
}

interface DeleteMaintenanceSchedulerOptions {
  batchWindowMs: number;
  runTask: (task: DeleteMaintenanceTask) => Promise<void>;
  onBusy: () => void;
  onIdle: () => void;
}

export class DeleteMaintenanceScheduler {
  private readonly pendingRequests: PendingDeleteMaintenanceRequest[] = [];
  private running = false;
  private drainTimer: ReturnType<typeof setTimeout> | null = null;
  private pendingTaskCount = 0;
  private destroyed = false;

  constructor(private readonly options: DeleteMaintenanceSchedulerOptions) {}

  enqueue(resolveTask: ResolveDeleteMaintenanceOperation): Promise<void> {
    if (this.destroyed) {
      return Promise.reject(new Error('Immersion tracker is shutting down'));
    }

    if (this.pendingTaskCount === 0) this.options.onBusy();
    this.pendingTaskCount += 1;

    const result = new Promise<void>((resolve, reject) => {
      this.pendingRequests.push({ resolveTask, resolve, reject });
      this.scheduleDrain();
    });

    return result.finally(() => {
      this.pendingTaskCount -= 1;
      if (this.pendingTaskCount === 0) this.options.onIdle();
    });
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    if (this.drainTimer) {
      clearTimeout(this.drainTimer);
      this.drainTimer = null;
    }
    const error = new Error('Immersion tracker is shutting down');
    for (const request of this.pendingRequests.splice(0)) request.reject(error);
  }

  private scheduleDrain(): void {
    if (this.destroyed || this.running || this.drainTimer || this.pendingRequests.length === 0) {
      return;
    }
    this.drainTimer = setTimeout(() => {
      this.drainTimer = null;
      void this.drain();
    }, this.options.batchWindowMs);
  }

  private async drain(): Promise<void> {
    if (this.running || this.pendingRequests.length === 0) return;
    this.running = true;
    const requests = this.pendingRequests.splice(0);
    const runnable: Array<{
      request: PendingDeleteMaintenanceRequest;
      task: DeleteMaintenanceOperation;
    }> = [];

    for (const request of requests) {
      try {
        const task = await request.resolveTask();
        if (task) runnable.push({ request, task });
        else request.resolve();
      } catch (error) {
        request.reject(error);
      }
    }

    if (runnable.length > 0) {
      const task: DeleteMaintenanceTask =
        runnable.length === 1
          ? runnable[0]!.task
          : { kind: 'batch', tasks: runnable.map((entry) => entry.task) };
      try {
        await this.options.runTask(task);
        for (const { request } of runnable) request.resolve();
      } catch (error) {
        for (const { request } of runnable) request.reject(error);
      }
    }

    this.running = false;
    this.scheduleDrain();
  }
}
