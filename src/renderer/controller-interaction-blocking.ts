type ControllerInteractionModalState = {
  controllerSelectModalOpen: boolean;
  controllerDebugModalOpen: boolean;
  jimakuModalOpen: boolean;
  kikuModalOpen: boolean;
  runtimeOptionsModalOpen: boolean;
  subsyncModalOpen: boolean;
  subtitleGenerationModalOpen?: boolean;
  youtubePickerModalOpen: boolean;
  sessionHelpModalOpen: boolean;
  subtitleSidebarModalOpen: boolean;
};

export function isControllerInteractionBlocked(state: ControllerInteractionModalState): boolean {
  return (
    state.controllerSelectModalOpen ||
    state.controllerDebugModalOpen ||
    state.jimakuModalOpen ||
    state.kikuModalOpen ||
    state.runtimeOptionsModalOpen ||
    state.subsyncModalOpen ||
    Boolean(state.subtitleGenerationModalOpen) ||
    state.youtubePickerModalOpen ||
    state.sessionHelpModalOpen
  );
}
