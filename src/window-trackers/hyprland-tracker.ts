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

import * as net from "net";
import { execSync } from "child_process";
import { BaseWindowTracker } from "./base-tracker";

interface HyprlandClient {
  class: string;
  at: [number, number];
  size: [number, number];
}

export class HyprlandWindowTracker extends BaseWindowTracker {
  private pollInterval: ReturnType<typeof setInterval> | null = null;
  private eventSocket: net.Socket | null = null;

  start(): void {
    this.pollInterval = setInterval(() => this.pollGeometry(), 250);
    this.pollGeometry();
    this.connectEventSocket();
  }

  stop(): void {
    if (this.pollInterval) {
      clearInterval(this.pollInterval);
      this.pollInterval = null;
    }
    if (this.eventSocket) {
      this.eventSocket.destroy();
      this.eventSocket = null;
    }
  }

  private connectEventSocket(): void {
    const hyprlandSig = process.env.HYPRLAND_INSTANCE_SIGNATURE;
    if (!hyprlandSig) {
      console.log("HYPRLAND_INSTANCE_SIGNATURE not set, skipping event socket");
      return;
    }

    const xdgRuntime = process.env.XDG_RUNTIME_DIR || "/tmp";
    const socketPath = `${xdgRuntime}/hypr/${hyprlandSig}/.socket2.sock`;
    this.eventSocket = new net.Socket();

    this.eventSocket.on("connect", () => {
      console.log("Connected to Hyprland event socket");
    });

    this.eventSocket.on("data", (data: Buffer) => {
      const events = data.toString().split("\n");
      for (const event of events) {
        if (
          event.includes("movewindow") ||
          event.includes("windowtitle") ||
          event.includes("openwindow") ||
          event.includes("closewindow") ||
          event.includes("fullscreen")
        ) {
          this.pollGeometry();
        }
      }
    });

    this.eventSocket.on("error", (err: Error) => {
      console.error("Hyprland event socket error:", err.message);
    });

    this.eventSocket.on("close", () => {
      console.log("Hyprland event socket closed");
    });

    this.eventSocket.connect(socketPath);
  }

  private pollGeometry(): void {
    try {
      const output = execSync("hyprctl clients -j", { encoding: "utf-8" });
      const clients: HyprlandClient[] = JSON.parse(output);
      const mpvWindow = clients.find((c) => c.class === "mpv");

      if (mpvWindow) {
        this.updateGeometry({
          x: mpvWindow.at[0],
          y: mpvWindow.at[1],
          width: mpvWindow.size[0],
          height: mpvWindow.size[1],
        });
      } else {
        this.updateGeometry(null);
      }
    } catch (err) {
      // hyprctl not available or failed - silent fail
    }
  }
}
