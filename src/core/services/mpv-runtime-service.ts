export interface MpvRuntimeClientLike {
  connected: boolean;
  send: (payload: { command: (string | number)[] }) => void;
  replayCurrentSubtitle?: () => void;
  playNextSubtitle?: () => void;
  setSubVisibility?: (visible: boolean) => void;
}

export function showMpvOsdRuntimeService(
  mpvClient: MpvRuntimeClientLike | null,
  text: string,
  fallbackLog: (text: string) => void = console.log,
): void {
  if (mpvClient && mpvClient.connected) {
    mpvClient.send({ command: ["show-text", text, "3000"] });
    return;
  }
  fallbackLog(`OSD (MPV not connected): ${text}`);
}

export function replayCurrentSubtitleRuntimeService(
  mpvClient: MpvRuntimeClientLike | null,
): void {
  if (!mpvClient?.replayCurrentSubtitle) return;
  mpvClient.replayCurrentSubtitle();
}

export function playNextSubtitleRuntimeService(
  mpvClient: MpvRuntimeClientLike | null,
): void {
  if (!mpvClient?.playNextSubtitle) return;
  mpvClient.playNextSubtitle();
}

export function sendMpvCommandRuntimeService(
  mpvClient: MpvRuntimeClientLike | null,
  command: (string | number)[],
): void {
  if (!mpvClient) return;
  mpvClient.send({ command });
}

export function setMpvSubVisibilityRuntimeService(
  mpvClient: MpvRuntimeClientLike | null,
  visible: boolean,
): void {
  if (!mpvClient?.setSubVisibility) return;
  mpvClient.setSubVisibility(visible);
}
