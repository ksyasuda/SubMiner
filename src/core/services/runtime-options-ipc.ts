import {
  RuntimeOptionApplyResult,
  RuntimeOptionId,
  RuntimeOptionValue,
} from '../../types/runtime-options';

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
