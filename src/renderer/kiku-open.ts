export function prepareForKikuFieldGroupingOpen(options: {
  closeLookupWindow: () => boolean;
  pausePlayback: () => void;
}): void {
  options.closeLookupWindow();
  options.pausePlayback();
}
