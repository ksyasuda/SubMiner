import type { ControllerConfigUpdate, RawConfig } from '../types';

type RawControllerConfig = NonNullable<RawConfig['controller']>;
type RawControllerBindings = NonNullable<RawControllerConfig['bindings']>;
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
