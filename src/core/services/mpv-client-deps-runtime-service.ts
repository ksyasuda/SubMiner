import {
  MpvIpcClientDeps,
} from "./mpv-service";
import { Config, MpvSubtitleRenderMetrics, SubtitleData } from "../../types";

interface SubtitleTimingTrackerLike {
  recordSubtitle: (text: string, start: number, end: number) => void;
}

export interface MpvClientDepsRuntimeOptions {
  getResolvedConfig: () => Config;
  autoStartOverlay: boolean;
  setOverlayVisible: (visible: boolean) => void;
  shouldBindVisibleOverlayToMpvSubVisibility: () => boolean;
  isVisibleOverlayVisible: () => boolean;
  getReconnectTimer: () => ReturnType<typeof setTimeout> | null;
  setReconnectTimer: (timer: ReturnType<typeof setTimeout> | null) => void;
  getCurrentSubText: () => string;
  setCurrentSubText: (text: string) => void;
  setCurrentSubAssText: (text: string) => void;
  getSubtitleTimingTracker: () => SubtitleTimingTrackerLike | null;
  subtitleWsBroadcast: (text: string) => void;
  getOverlayWindowsCount: () => number;
  tokenizeSubtitle: (text: string) => Promise<SubtitleData>;
  broadcastToOverlayWindows: (channel: string, ...args: unknown[]) => void;
  updateCurrentMediaPath: (mediaPath: unknown) => void;
  updateMpvSubtitleRenderMetrics: (
    patch: Partial<MpvSubtitleRenderMetrics>,
  ) => void;
  getMpvSubtitleRenderMetrics: () => MpvSubtitleRenderMetrics;
  setPreviousSecondarySubVisibility: (value: boolean | null) => void;
  showMpvOsd: (text: string) => void;
}

export function createMpvIpcClientDepsRuntimeService(
  options: MpvClientDepsRuntimeOptions,
): MpvIpcClientDeps {
  return {
    getResolvedConfig: options.getResolvedConfig,
    autoStartOverlay: options.autoStartOverlay,
    setOverlayVisible: options.setOverlayVisible,
    shouldBindVisibleOverlayToMpvSubVisibility:
      options.shouldBindVisibleOverlayToMpvSubVisibility,
    isVisibleOverlayVisible: options.isVisibleOverlayVisible,
    getReconnectTimer: options.getReconnectTimer,
    setReconnectTimer: options.setReconnectTimer,
    getCurrentSubText: options.getCurrentSubText,
    setCurrentSubText: options.setCurrentSubText,
    setCurrentSubAssText: options.setCurrentSubAssText,
    getSubtitleTimingTracker: options.getSubtitleTimingTracker,
    subtitleWsBroadcast: options.subtitleWsBroadcast,
    getOverlayWindowsCount: options.getOverlayWindowsCount,
    tokenizeSubtitle: options.tokenizeSubtitle,
    broadcastToOverlayWindows: options.broadcastToOverlayWindows,
    updateCurrentMediaPath: options.updateCurrentMediaPath,
    updateMpvSubtitleRenderMetrics: options.updateMpvSubtitleRenderMetrics,
    getMpvSubtitleRenderMetrics: options.getMpvSubtitleRenderMetrics,
    setPreviousSecondarySubVisibility: options.setPreviousSecondarySubVisibility,
    showMpvOsd: options.showMpvOsd,
  };
}
