import { NotificationOptions } from '../types';

export interface UiFeedbackState {
  progressDepth: number;
  progressTimer: ReturnType<typeof setInterval> | null;
  progressMessage: string;
  progressFrame: number;
}

export interface UiFeedbackResult {
  success: boolean;
  message: string;
}

export interface UiFeedbackNotificationContext {
  getNotificationType: () => string | undefined;
  showOsd: (text: string) => void;
  showSystemNotification: (title: string, options: NotificationOptions) => void;
}

export interface UiFeedbackOptions {
  setUpdateInProgress: (value: boolean) => void;
  showOsdNotification: (text: string) => void;
}

export function createUiFeedbackState(): UiFeedbackState {
  return {
    progressDepth: 0,
    progressTimer: null,
    progressMessage: '',
    progressFrame: 0,
  };
}

export function showStatusNotification(
  message: string,
  context: UiFeedbackNotificationContext,
): void {
  const type = context.getNotificationType() || 'osd';

  if (type === 'osd' || type === 'both') {
    context.showOsd(message);
  }

  if (type === 'system' || type === 'both') {
    context.showSystemNotification('SubMiner', { body: message });
  }
}

export function beginUpdateProgress(
  state: UiFeedbackState,
  initialMessage: string,
  showProgressTick: (text: string) => void,
): void {
  state.progressDepth += 1;
  if (state.progressDepth > 1) return;

  state.progressMessage = initialMessage;
  state.progressFrame = 0;
  showProgressTick(`${state.progressMessage}`);
  state.progressTimer = setInterval(() => {
    showProgressTick(`${state.progressMessage} ${['|', '/', '-', '\\'][state.progressFrame % 4]}`);
    state.progressFrame += 1;
  }, 180);
}

export function endUpdateProgress(
  state: UiFeedbackState,
  clearProgressTimer: (timer: ReturnType<typeof setInterval>) => void,
): void {
  state.progressDepth = Math.max(0, state.progressDepth - 1);
  if (state.progressDepth > 0) return;

  clearUpdateProgress(state, clearProgressTimer);
}

export function clearUpdateProgress(
  state: UiFeedbackState,
  clearProgressTimer: (timer: ReturnType<typeof setInterval>) => void,
): void {
  state.progressDepth = 0;

  if (state.progressTimer) {
    clearProgressTimer(state.progressTimer);
    state.progressTimer = null;
  }
  state.progressMessage = '';
  state.progressFrame = 0;
}

export function showProgressTick(
  state: UiFeedbackState,
  showOsdNotification: (text: string) => void,
): void {
  if (!state.progressMessage) return;
  const frames = ['|', '/', '-', '\\'];
  const frame = frames[state.progressFrame % frames.length];
  state.progressFrame += 1;
  showOsdNotification(`${state.progressMessage} ${frame}`);
}

export function showUpdateResult(
  state: UiFeedbackState,
  options: {
    clearProgressTimer: (timer: ReturnType<typeof setInterval>) => void;
    showOsdNotification: (text: string) => void;
  },
  result: UiFeedbackResult,
): void {
  clearUpdateProgress(state, options.clearProgressTimer);
  const prefix = result.success ? '✓' : 'x';
  options.showOsdNotification(`${prefix} ${result.message}`);
}

export async function withUpdateProgress<T>(
  state: UiFeedbackState,
  options: UiFeedbackOptions,
  initialMessage: string,
  action: () => Promise<T>,
): Promise<T> {
  beginUpdateProgress(state, initialMessage, () =>
    showProgressTick(state, options.showOsdNotification),
  );
  options.setUpdateInProgress(true);
  try {
    return await action();
  } finally {
    options.setUpdateInProgress(false);
    endUpdateProgress(state, clearInterval);
  }
}
