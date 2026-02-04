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

import { WindowGeometry } from "../types";

export type GeometryChangeCallback = (geometry: WindowGeometry) => void;
export type WindowFoundCallback = (geometry: WindowGeometry) => void;
export type WindowLostCallback = () => void;

export abstract class BaseWindowTracker {
  protected currentGeometry: WindowGeometry | null = null;
  protected windowFound: boolean = false;
  public onGeometryChange: GeometryChangeCallback | null = null;
  public onWindowFound: WindowFoundCallback | null = null;
  public onWindowLost: WindowLostCallback | null = null;

  abstract start(): void;
  abstract stop(): void;

  getGeometry(): WindowGeometry | null {
    return this.currentGeometry;
  }

  isTracking(): boolean {
    return this.windowFound;
  }

  protected updateGeometry(newGeometry: WindowGeometry | null): void {
    if (newGeometry) {
      if (!this.windowFound) {
        this.windowFound = true;
        if (this.onWindowFound) this.onWindowFound(newGeometry);
      }

      if (
        !this.currentGeometry ||
        this.currentGeometry.x !== newGeometry.x ||
        this.currentGeometry.y !== newGeometry.y ||
        this.currentGeometry.width !== newGeometry.width ||
        this.currentGeometry.height !== newGeometry.height
      ) {
        this.currentGeometry = newGeometry;
        if (this.onGeometryChange) this.onGeometryChange(newGeometry);
      }
    } else {
      if (this.windowFound) {
        this.windowFound = false;
        this.currentGeometry = null;
        if (this.onWindowLost) this.onWindowLost();
      }
    }
  }
}
