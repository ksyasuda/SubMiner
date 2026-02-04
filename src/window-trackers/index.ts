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

export type Compositor = "hyprland" | "sway" | "x11" | null;
export type Backend = "auto" | Exclude<Compositor, null>;

export function detectCompositor(): Compositor {
  if (process.env.HYPRLAND_INSTANCE_SIGNATURE) return "hyprland";
  if (process.env.SWAYSOCK) return "sway";
  if (process.env.XDG_SESSION_TYPE === "x11") return "x11";
  return null;
}

export function createWindowTracker(backend: Backend = "auto"): BaseWindowTracker | null {
  const compositor = backend === "auto" ? detectCompositor() : backend;
  console.log(`Detected compositor: ${compositor || "none"}`);

  switch (compositor) {
    case "hyprland":
      return new HyprlandWindowTracker();
    case "sway":
      return new SwayWindowTracker();
    case "x11":
      return new X11WindowTracker();
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
};
