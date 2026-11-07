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
import { createLogger } from '../logger';

const log = createLogger('overlay:windows');

type OverlayWin32 = Pick<
  typeof import('./win32'),
  'setOverlayOwner' | 'ensureOverlayTransparency' | 'bindOverlayAboveMpv' | 'clearOverlayOwner'
>;

// These run on every visibility update, so each operation warns at most every 30 seconds
// while it keeps failing, and once more when it recovers.
export function createWindowsOverlayOperations(deps: {
  loadWin32: () => OverlayWin32;
  warn: (message: string, details: Record<string, unknown>) => void;
  now: () => number;
}) {
  const failures = new Map<string, number>();

  const run = (
    operation: string,
    handles: { overlayHwnd: number; mpvHwnd?: number },
    action: (win32: OverlayWin32) => void,
  ): boolean => {
    try {
      action(deps.loadWin32());
      if (failures.delete(operation)) {
        deps.warn('Windows native overlay operation recovered', { operation, ...handles });
      }
      return true;
    } catch (error) {
      const lastWarning = failures.get(operation);
      if (lastWarning === undefined || deps.now() - lastWarning >= 30_000) {
        deps.warn('Windows native overlay operation failed', {
          operation,
          ...handles,
          error: error instanceof Error ? error.message : String(error),
        });
        failures.set(operation, deps.now());
      }
      return false;
    }
  };

  return {
    setOwner: (overlayHwnd: number, mpvHwnd: number): boolean =>
      run('set-owner', { overlayHwnd, mpvHwnd }, (win32) => {
        win32.setOverlayOwner(overlayHwnd, mpvHwnd);
      }),
    ensureTransparency: (overlayHwnd: number): boolean =>
      run('transparency', { overlayHwnd }, (win32) => {
        win32.ensureOverlayTransparency(overlayHwnd);
      }),
    bindAboveMpv: (overlayHwnd: number, mpvHwnd: number): boolean =>
      run('bind-above-mpv', { overlayHwnd, mpvHwnd }, (win32) => {
        win32.bindOverlayAboveMpv(overlayHwnd, mpvHwnd);
        win32.ensureOverlayTransparency(overlayHwnd);
      }),
    clearOwner: (overlayHwnd: number): boolean =>
      run('clear-owner', { overlayHwnd }, (win32) => {
        win32.clearOverlayOwner(overlayHwnd);
      }),
  };
}

function loadWin32(): typeof import('./win32') {
  return require('./win32') as typeof import('./win32');
}

const overlayOperations = createWindowsOverlayOperations({
  loadWin32,
  warn: (message, details) => log.warn(message, details),
  now: Date.now,
});

export const setWindowsOverlayOwner = overlayOperations.setOwner;
export const ensureWindowsOverlayTransparency = overlayOperations.ensureTransparency;
export const bindWindowsOverlayAboveMpv = overlayOperations.bindAboveMpv;
export const clearWindowsOverlayOwner = overlayOperations.clearOwner;

export function findWindowsMpvTargetWindowHandle(result?: MpvPollResult): number | null {
  const poll = result ?? loadWin32().findMpvWindows();
  const focused = poll.matches.find((match) => match.isForeground);
  const best =
    focused ??
    [...poll.matches].sort((a, b) => b.area - a.area || b.bounds.width - a.bounds.width)[0];
  return best?.hwnd ?? null;
}

export function getWindowsForegroundProcessName(): string | null {
  try {
    return loadWin32().getForegroundProcessName();
  } catch {
    return null;
  }
}
