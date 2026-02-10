import { AppContext } from "../../core/app-context";
import { SubminerModule } from "../../core/module";
import {
  JimakuApiResponse,
  JimakuDownloadQuery,
  JimakuDownloadResult,
  JimakuEntry,
  JimakuFileEntry,
  JimakuFilesQuery,
  JimakuMediaInfo,
  JimakuSearchQuery,
} from "../../types";

export class JimakuModule implements SubminerModule<AppContext> {
  readonly id = "jimaku";
  private context: AppContext["jimaku"] | undefined;

  init(context: AppContext): void {
    if (!context.jimaku) {
      throw new Error("Jimaku context is missing");
    }
    this.context = context.jimaku;
  }

  getMediaInfo(): JimakuMediaInfo {
    if (!this.context) {
      return {
        title: "",
        season: null,
        episode: null,
        confidence: "low",
        filename: "",
        rawTitle: "",
      };
    }
    return this.context.getMediaInfo();
  }

  searchEntries(
    query: JimakuSearchQuery,
  ): Promise<JimakuApiResponse<JimakuEntry[]>> {
    if (!this.context) {
      return Promise.resolve({
        ok: false,
        error: { error: "Jimaku module not initialized" },
      });
    }
    return this.context.searchEntries(query);
  }

  listFiles(
    query: JimakuFilesQuery,
  ): Promise<JimakuApiResponse<JimakuFileEntry[]>> {
    if (!this.context) {
      return Promise.resolve({
        ok: false,
        error: { error: "Jimaku module not initialized" },
      });
    }
    return this.context.listFiles(query);
  }

  downloadFile(query: JimakuDownloadQuery): Promise<JimakuDownloadResult> {
    if (!this.context) {
      return Promise.resolve({
        ok: false,
        error: { error: "Jimaku module not initialized" },
      });
    }
    return this.context.downloadFile(query);
  }
}
