import type { ControllerConfigUpdate, RawConfig } from '../types';

type RawControllerConfig = NonNullable<RawConfig['controller']>;
type RawControllerBindings = NonNullable<RawControllerConfig['bindings']>;
type RawControllerScreenBindings = Record<string, unknown>;
/** Per-dialog controller layouts stored as per-action patches. */
const SCREEN_BINDING_KEYS = ['reviewBindings', 'fieldGroupingBindings'] as const;
type RawControllerButtonIndices = NonNullable<RawControllerConfig['buttonIndices']>;
type RawControllerProfiles = NonNullable<RawControllerConfig['profiles']>;
type RawControllerProfile = NonNullable<RawControllerProfiles[string]>;

const RESERVED_CONTROLLER_PROFILE_IDS = new Set(['__proto__', 'constructor', 'prototype']);

function mergeBindingPatch(
  currentBindings: RawControllerBindings | undefined,
  updateBindings: RawControllerBindings | undefined,
): RawControllerBindings | undefined {
  if (!currentBindings && !updateBindings) return undefined;
  const nextBindings: RawControllerBindings = {
    ...(currentBindings ?? {}),
  };

  for (const [key, value] of Object.entries(updateBindings ?? {}) as Array<
    [keyof RawControllerBindings, RawControllerBindings[keyof RawControllerBindings] | undefined]
  >) {
    if (value === undefined) continue;
    (nextBindings as Record<string, unknown>)[key] = structuredClone(value);
  }

  return nextBindings;
}

function mergeScreenBindingPatch(
  currentBindings: RawControllerScreenBindings | undefined,
  updateBindings: RawControllerScreenBindings | undefined,
): RawControllerScreenBindings | undefined {
  if (!currentBindings && !updateBindings) return undefined;
  const nextBindings: RawControllerScreenBindings = { ...(currentBindings ?? {}) };
  for (const [action, binding] of Object.entries(updateBindings ?? {})) {
    if (binding === undefined) continue;
    nextBindings[action] = structuredClone(binding);
  }
  return nextBindings;
}

/** Merges every per-dialog layout patch from `update` onto `current` into `target`. */
function mergeScreenBindingPatches(
  current: Partial<Record<(typeof SCREEN_BINDING_KEYS)[number], unknown>> | undefined,
  update: Partial<Record<(typeof SCREEN_BINDING_KEYS)[number], unknown>>,
  target: Partial<Record<(typeof SCREEN_BINDING_KEYS)[number], unknown>>,
): void {
  for (const key of SCREEN_BINDING_KEYS) {
    const merged = mergeScreenBindingPatch(
      current?.[key] as RawControllerScreenBindings | undefined,
      update[key] as RawControllerScreenBindings | undefined,
    );
    if (merged) target[key] = merged;
  }
}

function mergeButtonIndexPatch(
  currentButtonIndices: RawControllerButtonIndices | undefined,
  updateButtonIndices: RawControllerButtonIndices | undefined,
): RawControllerButtonIndices | undefined {
  if (!currentButtonIndices && !updateButtonIndices) return undefined;
  return {
    ...(currentButtonIndices ?? {}),
    ...(updateButtonIndices ?? {}),
  };
}

function mergeControllerProfile(
  currentProfile: RawControllerProfile | undefined,
  updateProfile: RawControllerProfile,
): RawControllerProfile {
  const nextProfile: RawControllerProfile = {
    ...(currentProfile ?? {}),
    ...updateProfile,
  };

  const buttonIndices = mergeButtonIndexPatch(
    currentProfile?.buttonIndices,
    updateProfile.buttonIndices,
  );
  if (buttonIndices) {
    nextProfile.buttonIndices = buttonIndices;
  }

  const bindings = mergeBindingPatch(currentProfile?.bindings, updateProfile.bindings);
  if (bindings) {
    nextProfile.bindings = bindings;
  }

  mergeScreenBindingPatches(currentProfile, updateProfile, nextProfile);

  return nextProfile;
}

export function applyControllerConfigUpdate(
  currentController: RawConfig['controller'] | undefined,
  update: ControllerConfigUpdate,
): RawControllerConfig {
  const nextController: RawControllerConfig = {
    ...(currentController ?? {}),
    ...update,
  };

  const buttonIndices = mergeButtonIndexPatch(
    currentController?.buttonIndices,
    update.buttonIndices,
  );
  if (buttonIndices) {
    nextController.buttonIndices = buttonIndices;
  }

  const bindings = mergeBindingPatch(currentController?.bindings, update.bindings);
  if (bindings) {
    nextController.bindings = bindings;
  }

  mergeScreenBindingPatches(currentController, update, nextController);

  if (currentController?.profiles || update.profiles) {
    const nextProfiles: RawControllerProfiles = {};
    for (const [profileId, profile] of Object.entries(currentController?.profiles ?? {}) as Array<
      [string, RawControllerProfile]
    >) {
      if (RESERVED_CONTROLLER_PROFILE_IDS.has(profileId)) continue;
      nextProfiles[profileId] = profile;
    }
    for (const [profileId, profileUpdate] of Object.entries(update.profiles ?? {}) as Array<
      [string, RawControllerProfile | undefined]
    >) {
      if (RESERVED_CONTROLLER_PROFILE_IDS.has(profileId)) continue;
      if (profileUpdate === undefined) continue;
      nextProfiles[profileId] = mergeControllerProfile(
        currentController?.profiles?.[profileId],
        profileUpdate,
      );
    }
    nextController.profiles = nextProfiles;
  }

  return nextController;
}
