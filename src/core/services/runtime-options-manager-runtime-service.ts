import { RuntimeOptionsManager } from "../../runtime-options";
import { AnkiConnectConfig, RuntimeOptionState } from "../../types";

export interface RuntimeOptionsManagerRuntimeDeps {
  getAnkiConfig: () => AnkiConnectConfig;
  applyAnkiPatch: (patch: Partial<AnkiConnectConfig>) => void;
  onOptionsChanged: (options: RuntimeOptionState[]) => void;
}

export function createRuntimeOptionsManagerRuntimeService(
  deps: RuntimeOptionsManagerRuntimeDeps,
): RuntimeOptionsManager {
  return new RuntimeOptionsManager(deps.getAnkiConfig, {
    applyAnkiPatch: deps.applyAnkiPatch,
    onOptionsChanged: deps.onOptionsChanged,
  });
}
