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

import type { MpvPollResult } from './win32';

function loadWin32(): typeof import('./win32') {
  return require('./win32') as typeof import('./win32');
}

export function findWindowsMpvTargetWindowHandle(result?: MpvPollResult): number | null {
  const poll = result ?? loadWin32().findMpvWindows();
  const focused = poll.matches.find((match) => match.isForeground);
  const best =
    focused ?? [...poll.matches].sort((a, b) => b.area - a.area || b.bounds.width - a.bounds.width)[0];
  return best?.hwnd ?? null;
}

export function setWindowsOverlayOwner(overlayHwnd: number, mpvHwnd: number): boolean {
  try {
    loadWin32().setOverlayOwner(overlayHwnd, mpvHwnd);
    return true;
  } catch {
    return false;
  }
}

export function ensureWindowsOverlayTransparency(overlayHwnd: number): boolean {
  try {
    loadWin32().ensureOverlayTransparency(overlayHwnd);
    return true;
  } catch {
    return false;
  }
}

export function bindWindowsOverlayAboveMpv(overlayHwnd: number, mpvHwnd: number): boolean {
  try {
    const win32 = loadWin32();
    win32.bindOverlayAboveMpv(overlayHwnd, mpvHwnd);
    win32.ensureOverlayTransparency(overlayHwnd);
    return true;
  } catch {
    return false;
  }
}

export function clearWindowsOverlayOwner(overlayHwnd: number): boolean {
  try {
    loadWin32().clearOverlayOwner(overlayHwnd);
    return true;
  } catch {
    return false;
  }
}

export function getWindowsForegroundProcessName(): string | null {
  try {
    return loadWin32().getForegroundProcessName();
  } catch {
    return null;
  }
}
