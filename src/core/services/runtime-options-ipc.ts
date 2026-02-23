import { RuntimeOptionApplyResult, RuntimeOptionId, RuntimeOptionValue } from '../../types';

export interface RuntimeOptionsManagerLike {
  setOptionValue: (id: RuntimeOptionId, value: RuntimeOptionValue) => RuntimeOptionApplyResult;
  cycleOption: (id: RuntimeOptionId, direction: 1 | -1) => RuntimeOptionApplyResult;
}

export function applyRuntimeOptionResultRuntime(
  result: RuntimeOptionApplyResult,
  showMpvOsd: (text: string) => void,
): RuntimeOptionApplyResult {
  if (result.ok && result.osdMessage) {
    showMpvOsd(result.osdMessage);
  }
  return result;
}

export function setRuntimeOptionFromIpcRuntime(
  manager: RuntimeOptionsManagerLike | null,
  id: RuntimeOptionId,
  value: RuntimeOptionValue,
  showMpvOsd: (text: string) => void,
): RuntimeOptionApplyResult {
  if (!manager) {
    return { ok: false, error: 'Runtime options manager unavailable' };
  }
  const result = applyRuntimeOptionResultRuntime(manager.setOptionValue(id, value), showMpvOsd);
  if (!result.ok && result.error) {
    showMpvOsd(result.error);
  }
  return result;
}

export function cycleRuntimeOptionFromIpcRuntime(
  manager: RuntimeOptionsManagerLike | null,
  id: RuntimeOptionId,
  direction: 1 | -1,
  showMpvOsd: (text: string) => void,
): RuntimeOptionApplyResult {
  if (!manager) {
    return { ok: false, error: 'Runtime options manager unavailable' };
  }
  const result = applyRuntimeOptionResultRuntime(manager.cycleOption(id, direction), showMpvOsd);
  if (!result.ok && result.error) {
    showMpvOsd(result.error);
  }
  return result;
}
