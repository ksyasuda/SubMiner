export type AppAction =
  | { type: "overlay.toggleVisible" }
  | { type: "overlay.toggleInvisible" }
  | { type: "overlay.setVisible"; visible: boolean }
  | { type: "overlay.setInvisibleVisible"; visible: boolean }
  | { type: "overlay.openSettings" }
  | { type: "subtitle.copyCurrent" }
  | { type: "subtitle.copyMultiplePrompt"; timeoutMs: number }
  | { type: "anki.mineSentence" }
  | { type: "anki.mineSentenceMultiplePrompt"; timeoutMs: number }
  | { type: "anki.updateLastCardFromClipboard" }
  | { type: "anki.markAudioCard" }
  | { type: "kiku.triggerFieldGrouping" }
  | { type: "subsync.triggerFromConfig" }
  | { type: "secondarySub.toggleMode" }
  | { type: "runtimeOptions.openPalette" };
