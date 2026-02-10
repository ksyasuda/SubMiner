/*
  subminer - Yomitan integration for mpv
  Copyright (C) 2024 sudacode

  This program is free software: you can redistribute it and/or modify
  it under the terms of the GNU General Public License as published by
  the Free Software Foundation, either version 3 of the License, or
  (at your option) any later version.

  This program is distributed in the hope that it will be useful,
  but WITHOUT ANY WARRANTY; without even the implied warranty of
  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
  GNU General Public License for more details.

  You should have received a copy of the GNU General Public License
  along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

import { execFile } from "child_process";
import { BaseWindowTracker } from "./base-tracker";

export class MacOSWindowTracker extends BaseWindowTracker {
  private pollInterval: ReturnType<typeof setInterval> | null = null;
  private pollInFlight = false;

  start(): void {
    this.pollInterval = setInterval(() => this.pollGeometry(), 250);
    this.pollGeometry();
  }

  stop(): void {
    if (this.pollInterval) {
      clearInterval(this.pollInterval);
      this.pollInterval = null;
    }
  }

  private pollGeometry(): void {
    if (this.pollInFlight) {
      return;
    }

    this.pollInFlight = true;

    const script = `
      set processNames to {"mpv", "MPV", "org.mpv.mpv"}
      tell application "System Events"
        repeat with procName in processNames
          set procList to (every process whose name is procName)
          repeat with p in procList
            try
              if (count of windows of p) > 0 then
                set targetWindow to window 1 of p
                set windowPos to position of targetWindow
                set windowSize to size of targetWindow
                return (item 1 of windowPos) & "," & (item 2 of windowPos) & "," & (item 1 of windowSize) & "," & (item 2 of windowSize)
              end if
            end try
          end repeat
        end repeat
      end tell
      return "not-found"
    `;

    execFile(
      "osascript",
      ["-e", script],
      {
        encoding: "utf-8",
        timeout: 1000,
        maxBuffer: 1024 * 1024,
      },
      (err, stdout) => {
        if (err) {
          this.updateGeometry(null);
          this.pollInFlight = false;
          return;
        }

        const result = (stdout || "").trim();
        if (result && result !== "not-found") {
          const parts = result.split(",");
          if (parts.length === 4) {
            const x = parseInt(parts[0], 10);
            const y = parseInt(parts[1], 10);
            const width = parseInt(parts[2], 10);
            const height = parseInt(parts[3], 10);

            if (
              Number.isFinite(x) &&
              Number.isFinite(y) &&
              Number.isFinite(width) &&
              Number.isFinite(height) &&
              width > 0 &&
              height > 0
            ) {
              this.updateGeometry({
                x,
                y,
                width,
                height,
              });
              this.pollInFlight = false;
              return;
            }
          }
        }

        this.updateGeometry(null);
        this.pollInFlight = false;
      },
    );
  }
}
