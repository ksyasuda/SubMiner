import type { ControllerConfigUpdate, RawConfig } from '../types';

type RawControllerConfig = NonNullable<RawConfig['controller']>;
type RawControllerBindings = NonNullable<RawControllerConfig['bindings']>;

export function applyControllerConfigUpdate(
  currentController: RawConfig['controller'] | undefined,
  update: ControllerConfigUpdate,
): RawControllerConfig {
  const nextController: RawControllerConfig = {
    ...(currentController ?? {}),
    ...update,
  };

  if (currentController?.buttonIndices || update.buttonIndices) {
    nextController.buttonIndices = {
      ...(currentController?.buttonIndices ?? {}),
      ...(update.buttonIndices ?? {}),
    };
  }

  if (currentController?.bindings || update.bindings) {
    const nextBindings: RawControllerBindings = {
      ...(currentController?.bindings ?? {}),
    };

    for (const [key, value] of Object.entries(update.bindings ?? {}) as Array<
      [keyof RawControllerBindings, RawControllerBindings[keyof RawControllerBindings] | undefined]
    >) {
      if (value === undefined) continue;
      (nextBindings as Record<string, unknown>)[key] = JSON.parse(JSON.stringify(value));
    }

    nextController.bindings = nextBindings;
  }

  return nextController;
}
