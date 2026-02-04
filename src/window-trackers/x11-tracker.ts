/*
  SubMiner - All-in-one sentence mining overlay
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

import { execSync } from "child_process";
import { BaseWindowTracker } from "./base-tracker";

export class X11WindowTracker extends BaseWindowTracker {
  private pollInterval: ReturnType<typeof setInterval> | null = null;

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
    try {
      const windowIds = execSync("xdotool search --class mpv", {
        encoding: "utf-8",
      }).trim();

      if (!windowIds) {
        this.updateGeometry(null);
        return;
      }

      const windowId = windowIds.split("\n")[0];

      const winInfo = execSync(`xwininfo -id ${windowId}`, {
        encoding: "utf-8",
      });

      const xMatch = winInfo.match(/Absolute upper-left X:\s*(\d+)/);
      const yMatch = winInfo.match(/Absolute upper-left Y:\s*(\d+)/);
      const widthMatch = winInfo.match(/Width:\s*(\d+)/);
      const heightMatch = winInfo.match(/Height:\s*(\d+)/);

      if (xMatch && yMatch && widthMatch && heightMatch) {
        this.updateGeometry({
          x: parseInt(xMatch[1], 10),
          y: parseInt(yMatch[1], 10),
          width: parseInt(widthMatch[1], 10),
          height: parseInt(heightMatch[1], 10),
        });
      } else {
        this.updateGeometry(null);
      }
    } catch (err) {
      this.updateGeometry(null);
    }
  }
}
