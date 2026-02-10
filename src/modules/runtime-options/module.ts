import { AppContext } from "../../core/app-context";
import { SubminerModule } from "../../core/module";
import { RuntimeOptionsManager } from "../../runtime-options";
import {
  AnkiConnectConfig,
  RuntimeOptionApplyResult,
  RuntimeOptionId,
  RuntimeOptionState,
  RuntimeOptionValue,
} from "../../types";

export class RuntimeOptionsModule implements SubminerModule<AppContext> {
  readonly id = "runtime-options";
  private manager: RuntimeOptionsManager | null = null;

  init(context: AppContext): void {
    if (!context.runtimeOptions) {
      throw new Error("Runtime options context is missing");
    }

    this.manager = new RuntimeOptionsManager(
      context.runtimeOptions.getAnkiConfig,
      {
        applyAnkiPatch: context.runtimeOptions.applyAnkiPatch,
        onOptionsChanged: context.runtimeOptions.onOptionsChanged,
      },
    );
  }

  listOptions(): RuntimeOptionState[] {
    return this.manager ? this.manager.listOptions() : [];
  }

  getOptionValue(id: RuntimeOptionId): RuntimeOptionValue | undefined {
    return this.manager?.getOptionValue(id);
  }

  setOptionValue(
    id: RuntimeOptionId,
    value: RuntimeOptionValue,
  ): RuntimeOptionApplyResult {
    if (!this.manager) {
      return { ok: false, error: "Runtime options manager unavailable" };
    }
    return this.manager.setOptionValue(id, value);
  }

  cycleOption(id: RuntimeOptionId, direction: 1 | -1): RuntimeOptionApplyResult {
    if (!this.manager) {
      return { ok: false, error: "Runtime options manager unavailable" };
    }
    return this.manager.cycleOption(id, direction);
  }

  getEffectiveAnkiConnectConfig(baseConfig?: AnkiConnectConfig): AnkiConnectConfig {
    if (!this.manager) {
      return baseConfig ? JSON.parse(JSON.stringify(baseConfig)) : {};
    }
    return this.manager.getEffectiveAnkiConnectConfig(baseConfig);
  }
}
