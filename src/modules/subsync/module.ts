import { AppContext } from "../../core/app-context";
import { SubminerModule } from "../../core/module";
import { SubsyncManualRunRequest, SubsyncResult } from "../../types";

export class SubsyncModule implements SubminerModule<AppContext> {
  readonly id = "subsync";
  private inProgress = false;
  private context: AppContext["subsync"] | undefined;

  init(context: AppContext): void {
    if (!context.subsync) {
      throw new Error("Subsync context is missing");
    }
    this.context = context.subsync;
  }

  isInProgress(): boolean {
    return this.inProgress;
  }

  async triggerFromConfig(): Promise<void> {
    if (!this.context) {
      throw new Error("Subsync module not initialized");
    }

    if (this.inProgress) {
      this.context.showOsd("Subsync already running");
      return;
    }

    try {
      if (this.context.getDefaultMode() === "manual") {
        await this.context.openManualPicker();
        this.context.showOsd("Subsync: choose engine and source");
        return;
      }

      this.inProgress = true;
      const result = await this.context.runWithSpinner(
        () => this.context!.runAuto(),
        "Subsync: syncing",
      );
      this.context.showOsd(result.message);
    } catch (error) {
      this.context.showOsd(`Subsync failed: ${(error as Error).message}`);
    } finally {
      this.inProgress = false;
    }
  }

  async runManual(request: SubsyncManualRunRequest): Promise<SubsyncResult> {
    if (!this.context) {
      return { ok: false, message: "Subsync module not initialized" };
    }

    if (this.inProgress) {
      const busy = "Subsync already running";
      this.context.showOsd(busy);
      return { ok: false, message: busy };
    }

    try {
      this.inProgress = true;
      const result = await this.context.runWithSpinner(
        () => this.context!.runManual(request),
        "Subsync: syncing",
      );
      this.context.showOsd(result.message);
      return result;
    } catch (error) {
      const message = `Subsync failed: ${(error as Error).message}`;
      this.context.showOsd(message);
      return { ok: false, message };
    } finally {
      this.inProgress = false;
    }
  }
}
