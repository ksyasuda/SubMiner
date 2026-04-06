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

import { BaseWindowTracker } from './base-tracker';
import { HyprlandWindowTracker } from './hyprland-tracker';
import { KWinWindowTracker } from './kwin-tracker';
import { SwayWindowTracker } from './sway-tracker';
import { X11WindowTracker } from './x11-tracker';
import { MacOSWindowTracker } from './macos-tracker';
import { WindowsWindowTracker } from './windows-tracker';
import { detectSessionBackend, type SessionBackend } from '../shared/backend-detection';
import { createLogger } from '../logger';

const log = createLogger('tracker');

export type Compositor = SessionBackend;
export type Backend = 'auto' | Exclude<Compositor, null>;

export function detectCompositor(): Compositor {
  return detectSessionBackend();
}

function normalizeCompositor(value: string): Compositor | null {
  const normalized = value.trim().toLowerCase();
  if (normalized === 'hyprland') return 'hyprland';
  if (normalized === 'kwin') return 'kwin';
  if (normalized === 'sway') return 'sway';
  if (normalized === 'x11') return 'x11';
  if (normalized === 'macos') return 'macos';
  if (normalized === 'windows') return 'windows';
  return null;
}

export function createWindowTracker(
  override?: string | null,
  targetMpvSocketPath?: string | null,
): BaseWindowTracker | null {
  let compositor = detectCompositor();

  if (override && override !== 'auto') {
    const normalized = normalizeCompositor(override);
    if (normalized) {
      compositor = normalized;
    } else {
      log.warn(`Unsupported backend override "${override}", falling back to auto.`);
    }
  }
  log.info(`Detected compositor: ${compositor || 'none'}`);

  switch (compositor) {
    case 'hyprland':
      return new HyprlandWindowTracker(targetMpvSocketPath?.trim() || undefined);
    case 'kwin':
      return new KWinWindowTracker(targetMpvSocketPath?.trim() || undefined);
    case 'sway':
      return new SwayWindowTracker(targetMpvSocketPath?.trim() || undefined);
    case 'x11':
      return new X11WindowTracker(targetMpvSocketPath?.trim() || undefined);
    case 'macos':
      return new MacOSWindowTracker(targetMpvSocketPath?.trim() || undefined);
    case 'windows':
      return new WindowsWindowTracker(targetMpvSocketPath?.trim() || undefined);
    default:
      log.warn('No supported compositor detected. Window tracking disabled.');
      return null;
  }
}

export {
  BaseWindowTracker,
  HyprlandWindowTracker,
  KWinWindowTracker,
  SwayWindowTracker,
  X11WindowTracker,
  MacOSWindowTracker,
  WindowsWindowTracker,
};
