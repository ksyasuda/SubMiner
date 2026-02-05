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

import { BaseWindowTracker } from "./base-tracker";
import { HyprlandWindowTracker } from "./hyprland-tracker";
import { SwayWindowTracker } from "./sway-tracker";
import { X11WindowTracker } from "./x11-tracker";
import { MacOSWindowTracker } from "./macos-tracker";

export type Compositor = "hyprland" | "sway" | "x11" | "macos" | null;
export type Backend = "auto" | Exclude<Compositor, null>;

export function detectCompositor(): Compositor {
  if (process.platform === "darwin") return "macos";
  if (process.env.HYPRLAND_INSTANCE_SIGNATURE) return "hyprland";
  if (process.env.SWAYSOCK) return "sway";
  if (process.env.XDG_SESSION_TYPE === "x11") return "x11";
  return null;
}

function normalizeCompositor(value: string): Compositor | null {
  const normalized = value.trim().toLowerCase();
  if (normalized === "hyprland") return "hyprland";
  if (normalized === "sway") return "sway";
  if (normalized === "x11") return "x11";
  if (normalized === "macos") return "macos";
  return null;
}

export function createWindowTracker(
  override?: string | null,
): BaseWindowTracker | null {
  let compositor = detectCompositor();

  if (override && override !== "auto") {
    const normalized = normalizeCompositor(override);
    if (normalized) {
      compositor = normalized;
    } else {
      console.warn(
        `Unsupported backend override "${override}", falling back to auto.`,
      );
    }
  }
  console.log(`Detected compositor: ${compositor || "none"}`);

  switch (compositor) {
    case "hyprland":
      return new HyprlandWindowTracker();
    case "sway":
      return new SwayWindowTracker();
    case "x11":
      return new X11WindowTracker();
    case "macos":
      return new MacOSWindowTracker();
    default:
      console.warn(
        "No supported compositor detected. Window tracking disabled.",
      );
      return null;
  }
}

export {
  BaseWindowTracker,
  HyprlandWindowTracker,
  SwayWindowTracker,
  X11WindowTracker,
  MacOSWindowTracker,
};
