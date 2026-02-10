interface RuntimeAutoUpdateOptionManagerLike {
  getOptionValue: (id: "anki.autoUpdateNewCards") => unknown;
}

interface RuntimeConfigLike {
  auto_start_overlay?: boolean;
  bind_visible_overlay_to_mpv_sub_visibility: boolean;
  invisibleOverlay: {
    startupVisibility: "visible" | "hidden" | "platform-default";
  };
  ankiConnect?: {
    behavior?: {
      autoUpdateNewCards?: boolean;
    };
  };
}

export function getInitialInvisibleOverlayVisibilityService(
  config: RuntimeConfigLike,
  platform: NodeJS.Platform,
): boolean {
  const visibility = config.invisibleOverlay.startupVisibility;
  if (visibility === "visible") return true;
  if (visibility === "hidden") return false;
  if (platform === "linux") return false;
  return true;
}

export function shouldAutoInitializeOverlayRuntimeFromConfigService(
  config: RuntimeConfigLike,
): boolean {
  if (config.auto_start_overlay === true) return true;
  if (config.invisibleOverlay.startupVisibility === "visible") return true;
  return false;
}

export function shouldBindVisibleOverlayToMpvSubVisibilityService(
  config: RuntimeConfigLike,
): boolean {
  return config.bind_visible_overlay_to_mpv_sub_visibility;
}

export function isAutoUpdateEnabledRuntimeService(
  config: RuntimeConfigLike,
  runtimeOptionsManager: RuntimeAutoUpdateOptionManagerLike | null,
): boolean {
  const value = runtimeOptionsManager?.getOptionValue("anki.autoUpdateNewCards");
  if (typeof value === "boolean") return value;
  return config.ankiConnect?.behavior?.autoUpdateNewCards !== false;
}
