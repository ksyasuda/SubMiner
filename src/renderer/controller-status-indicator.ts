import type { ControllerDeviceInfo } from '../types';

type ControllerSnapshot = {
  connectedGamepads: ControllerDeviceInfo[];
  activeGamepadId: string | null;
};

function getDeviceLabel(device: ControllerDeviceInfo | undefined): string {
  if (!device) return 'Controller';
  return device.id || `Gamepad ${device.index}`;
}

/** Watches controller snapshots and calls `notify` when a new device appears. */
export function createControllerStatusIndicator(notify: (message: string) => void) {
  let previousConnectedIds = new Set<string>();

  function update(snapshot: ControllerSnapshot): void {
    const newDevices = snapshot.connectedGamepads.filter(
      (device) => !previousConnectedIds.has(device.id),
    );
    if (newDevices.length > 0) {
      const announcedDevice =
        newDevices.find((device) => device.id === snapshot.activeGamepadId) ?? newDevices[0];
      notify(`Controller connected: ${getDeviceLabel(announcedDevice)}`);
    }

    previousConnectedIds = new Set(snapshot.connectedGamepads.map((device) => device.id));
  }

  return { update };
}
