import type { ControllerDeviceInfo } from '../types';

type ControllerSnapshot = {
  connectedGamepads: ControllerDeviceInfo[];
  activeGamepadId: string | null;
};

type ControllerStatusIndicatorOptions = {
  durationMs?: number;
  setTimeout?: (callback: () => void, delay: number) => ReturnType<typeof setTimeout>;
  clearTimeout?: (timer: ReturnType<typeof setTimeout> | number) => void;
};

function getDeviceLabel(device: ControllerDeviceInfo | undefined): string {
  if (!device) return 'Controller';
  return device.id || `Gamepad ${device.index}`;
}

export function createControllerStatusIndicator(
  dom: {
    controllerStatusToast: {
      textContent: string;
      classList: { add: (...entries: string[]) => void; remove: (...entries: string[]) => void };
    };
  },
  options: ControllerStatusIndicatorOptions = {},
) {
  const durationMs = options.durationMs ?? 2200;
  const scheduleTimeout = options.setTimeout ?? globalThis.setTimeout;
  const cancelTimeout =
    options.clearTimeout ??
    ((timer: ReturnType<typeof setTimeout> | number) =>
      globalThis.clearTimeout(timer as ReturnType<typeof setTimeout>));
  let hideTimeout: ReturnType<typeof setTimeout> | number | null = null;
  let previousConnectedIds = new Set<string>();

  function show(message: string): void {
    if (hideTimeout !== null) {
      cancelTimeout(hideTimeout);
      hideTimeout = null;
    }

    dom.controllerStatusToast.textContent = message;
    dom.controllerStatusToast.classList.remove('hidden');
    hideTimeout = scheduleTimeout(() => {
      dom.controllerStatusToast.classList.add('hidden');
      dom.controllerStatusToast.textContent = '';
      hideTimeout = null;
    }, durationMs);
  }

  function update(snapshot: ControllerSnapshot): void {
    const newDevices = snapshot.connectedGamepads.filter(
      (device) => !previousConnectedIds.has(device.id),
    );
    if (newDevices.length > 0) {
      const activeDevice = snapshot.connectedGamepads.find(
        (device) => device.id === snapshot.activeGamepadId,
      );
      const announcedDevice =
        newDevices.find((device) => device.id === snapshot.activeGamepadId) ??
        newDevices[0] ??
        activeDevice;
      show(`Controller detected: ${getDeviceLabel(announcedDevice)}`);
    }

    previousConnectedIds = new Set(snapshot.connectedGamepads.map((device) => device.id));
  }

  return { update };
}
