type ParserWindowLike = {
  isDestroyed?: () => boolean;
  destroy?: () => void;
} | null;

export interface YomitanParserRuntimeStateDeps {
  getYomitanParserWindow: () => ParserWindowLike;
  setYomitanParserWindow: (window: null) => void;
  setYomitanParserReadyPromise: (promise: Promise<void> | null) => void;
  setYomitanParserInitPromise: (promise: Promise<boolean> | null) => void;
}

export interface YomitanExtensionRuntimeStateDeps extends YomitanParserRuntimeStateDeps {
  setYomitanExtension: (extension: null) => void;
  setYomitanSession: (session: null) => void;
}

export function clearYomitanParserRuntimeState(deps: YomitanParserRuntimeStateDeps): void {
  const parserWindow = deps.getYomitanParserWindow();
  if (parserWindow && !parserWindow.isDestroyed?.()) {
    parserWindow.destroy?.();
  }
  deps.setYomitanParserWindow(null);
  deps.setYomitanParserReadyPromise(null);
  deps.setYomitanParserInitPromise(null);
}

export function clearYomitanExtensionRuntimeState(deps: YomitanExtensionRuntimeStateDeps): void {
  clearYomitanParserRuntimeState(deps);
  deps.setYomitanExtension(null);
  deps.setYomitanSession(null);
}
