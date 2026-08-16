import { execFile } from 'node:child_process';
import koffi from 'koffi';
import { matchesMpvSocketPathInCommandLine } from './mpv-socket-match';

const user32 = koffi.load('user32.dll');
const dwmapi = koffi.load('dwmapi.dll');
const kernel32 = koffi.load('kernel32.dll');

const RECT = koffi.struct('RECT', {
  Left: 'int',
  Top: 'int',
  Right: 'int',
  Bottom: 'int',
});

const MARGINS = koffi.struct('MARGINS', {
  cxLeftWidth: 'int',
  cxRightWidth: 'int',
  cyTopHeight: 'int',
  cyBottomHeight: 'int',
});

const WNDENUMPROC = koffi.proto('bool __stdcall WNDENUMPROC(intptr hwnd, intptr lParam)');

const EnumWindows = user32.func('bool __stdcall EnumWindows(WNDENUMPROC *cb, intptr lParam)');
const IsWindowVisible = user32.func('bool __stdcall IsWindowVisible(intptr hwnd)');
const IsIconic = user32.func('bool __stdcall IsIconic(intptr hwnd)');
const GetForegroundWindow = user32.func('intptr __stdcall GetForegroundWindow()');
const SetWindowPos = user32.func(
  'bool __stdcall SetWindowPos(intptr hwnd, intptr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags)',
);
const GetWindowThreadProcessId = user32.func(
  'uint __stdcall GetWindowThreadProcessId(intptr hwnd, _Out_ uint *lpdwProcessId)',
);
const GetWindowLongW = user32.func('int __stdcall GetWindowLongW(intptr hwnd, int nIndex)');
const SetWindowLongPtrW = user32.func(
  'intptr __stdcall SetWindowLongPtrW(intptr hwnd, int nIndex, intptr dwNewLong)',
);
const GetWindowFn = user32.func('intptr __stdcall GetWindow(intptr hwnd, uint uCmd)');
const GetWindowRect = user32.func('bool __stdcall GetWindowRect(intptr hwnd, _Out_ RECT *lpRect)');

const DwmGetWindowAttribute = dwmapi.func(
  'int __stdcall DwmGetWindowAttribute(intptr hwnd, uint dwAttribute, _Out_ RECT *pvAttribute, uint cbAttribute)',
);
const DwmExtendFrameIntoClientArea = dwmapi.func(
  'int __stdcall DwmExtendFrameIntoClientArea(intptr hwnd, MARGINS *pMarInset)',
);

const OpenProcess = kernel32.func(
  'intptr __stdcall OpenProcess(uint dwDesiredAccess, bool bInheritHandle, uint dwProcessId)',
);
const CloseHandle = kernel32.func('bool __stdcall CloseHandle(intptr hObject)');
const GetLastError = kernel32.func('uint __stdcall GetLastError()');
const SetLastError = kernel32.func('void __stdcall SetLastError(uint dwErrCode)');
const QueryFullProcessImageNameW = kernel32.func(
  'bool __stdcall QueryFullProcessImageNameW(intptr hProcess, uint dwFlags, _Out_ uint16 *lpExeName, _Inout_ uint *lpdwSize)',
);

const GWL_EXSTYLE = -20;
const WS_EX_TOPMOST = 0x00000008;
const GWLP_HWNDPARENT = -8;
const GW_HWNDPREV = 3;
const DWMWA_EXTENDED_FRAME_BOUNDS = 9;
const PROCESS_QUERY_LIMITED_INFORMATION = 0x1000;
const SWP_NOSIZE = 0x0001;
const SWP_NOMOVE = 0x0002;
const SWP_NOACTIVATE = 0x0010;
const SWP_NOOWNERZORDER = 0x0200;
const SWP_FLAGS = SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOOWNERZORDER;
const HWND_TOP = 0;
const HWND_BOTTOM = 1;
const HWND_TOPMOST = -1;
const HWND_NOTOPMOST = -2;

function extendOverlayFrameIntoClientArea(overlayHwnd: number): void {
  const hr = DwmExtendFrameIntoClientArea(overlayHwnd, {
    cxLeftWidth: -1,
    cxRightWidth: -1,
    cyTopHeight: -1,
    cyBottomHeight: -1,
  });
  if (hr !== 0) {
    throw new Error(`DwmExtendFrameIntoClientArea failed (${hr})`);
  }
}

function resetLastError(): void {
  SetLastError(0);
}

function assertSetWindowLongPtrSucceeded(operation: string, result: number): void {
  if (result !== 0) {
    return;
  }

  if (GetLastError() === 0) {
    return;
  }

  throw new Error(`${operation} failed (${GetLastError()})`);
}

function assertSetWindowPosSucceeded(operation: string, result: boolean): void {
  if (result) {
    return;
  }

  throw new Error(`${operation} failed (${GetLastError()})`);
}

function assertGetWindowLongSucceeded(operation: string, result: number): number {
  if (result !== 0 || GetLastError() === 0) {
    return result;
  }

  throw new Error(`${operation} failed (${GetLastError()})`);
}

export interface WindowBounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface MpvWindowMatch {
  hwnd: number;
  bounds: WindowBounds;
  area: number;
  isForeground: boolean;
  commandLine?: string | null;
}

export interface MpvPollResult {
  matches: MpvWindowMatch[];
  focusState: boolean;
  windowState: 'visible' | 'minimized' | 'not-found';
}

function getWindowBounds(hwnd: number): WindowBounds | null {
  const rect = { Left: 0, Top: 0, Right: 0, Bottom: 0 };
  const hr = DwmGetWindowAttribute(hwnd, DWMWA_EXTENDED_FRAME_BOUNDS, rect, koffi.sizeof(RECT));
  if (hr !== 0) {
    if (!GetWindowRect(hwnd, rect)) {
      return null;
    }
  }

  const width = rect.Right - rect.Left;
  const height = rect.Bottom - rect.Top;
  if (width <= 0 || height <= 0) return null;

  return { x: rect.Left, y: rect.Top, width, height };
}

function getProcessNameByPid(pid: number): string | null {
  const hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid);
  if (!hProcess) return null;

  try {
    const buffer = new Uint16Array(260);
    const size = new Uint32Array([260]);

    if (!QueryFullProcessImageNameW(hProcess, 0, buffer, size)) {
      return null;
    }

    const fullPath = String.fromCharCode(...buffer.slice(0, size[0]));
    const fileName = fullPath.split('\\').pop() || '';
    return fileName.replace(/\.exe$/i, '');
  } finally {
    CloseHandle(hProcess);
  }
}

// Short-lived cache so the 250ms poll doesn't re-query every top-level window's process
// on each pass. The TTL bounds staleness from PID reuse.
const PROCESS_NAME_CACHE_TTL_MS = 5_000;
const PROCESS_NAME_CACHE_PRUNE_THRESHOLD = 512;
const processNameCache = new Map<number, { name: string | null; expiresAtMs: number }>();

function getCachedProcessNameByPid(pid: number): string | null {
  const nowMs = Date.now();
  const cached = processNameCache.get(pid);
  if (cached && cached.expiresAtMs > nowMs) {
    return cached.name;
  }
  const name = getProcessNameByPid(pid);
  processNameCache.set(pid, { name, expiresAtMs: nowMs + PROCESS_NAME_CACHE_TTL_MS });
  return name;
}

function pruneExpiredProcessNames(nowMs: number): void {
  if (processNameCache.size <= PROCESS_NAME_CACHE_PRUNE_THRESHOLD) return;
  for (const [pid, entry] of processNameCache) {
    if (entry.expiresAtMs <= nowMs) {
      processNameCache.delete(pid);
    }
  }
}

type ProcessCommandLineCacheEntry =
  | { state: 'resolved'; commandLine: string; expiresAtMs: number; refreshInFlight: boolean }
  | { state: 'pending' }
  | { state: 'failed'; retryAtMs: number; backoffMs: number };

const COMMAND_LINE_RETRY_INITIAL_BACKOFF_MS = 2_000;
const COMMAND_LINE_RETRY_MAX_BACKOFF_MS = 30_000;
// A process command line never changes, so a resolved entry only has to expire to survive
// Windows PID reuse (a dead mpv's PID handed to a new instance, whose stale socket path would
// otherwise match the wrong window forever). Longer than the process-name TTL because each
// refresh costs a PowerShell spawn, and the cached value keeps being served while the refresh
// runs, so expiry never interrupts window matching.
const COMMAND_LINE_CACHE_TTL_MS = 60_000;
const processCommandLineCache = new Map<number, ProcessCommandLineCacheEntry>();

function queryProcessCommandLine(
  pid: number,
  onResult: (commandLine: string | null) => void,
): void {
  execFile(
    'powershell.exe',
    [
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy',
      'Bypass',
      '-Command',
      `$process = Get-CimInstance Win32_Process -Filter "ProcessId = ${pid}"; if ($process -and $process.CommandLine) { [Console]::Out.Write($process.CommandLine) }`,
    ],
    {
      encoding: 'utf8',
      windowsHide: true,
      timeout: 1500,
    },
    (error, stdout) => {
      const output = error ? '' : stdout.trim();
      onResult(output.length > 0 ? output : null);
    },
  );
}

// Resolves a process command line via a background PowerShell lookup. Returns null until the
// first lookup completes; the caller's next poll picks up the cached result. Failures are
// negative-cached with exponential backoff: the synchronous version of this lookup could
// block the main thread for its full 1.5s timeout on every 250ms poll, which (combined with
// the forward:true mouse hook) stalled mouse input system-wide.
function getProcessCommandLineByPid(pid: number): string | null {
  const entry = processCommandLineCache.get(pid);
  const nowMs = Date.now();

  if (entry?.state === 'resolved') {
    if (nowMs >= entry.expiresAtMs && !entry.refreshInFlight) {
      entry.refreshInFlight = true;
      queryProcessCommandLine(pid, (commandLine) => {
        processCommandLineCache.set(pid, {
          state: 'resolved',
          // A failed refresh is usually a transient query error rather than a dead process
          // (a gone process owns no window, so it is never looked up again). Keep the last
          // known command line and re-check after the next TTL.
          commandLine: commandLine ?? entry.commandLine,
          expiresAtMs: Date.now() + COMMAND_LINE_CACHE_TTL_MS,
          refreshInFlight: false,
        });
      });
    }
    return entry.commandLine;
  }

  if (entry?.state === 'pending') return null;
  if (entry?.state === 'failed' && nowMs < entry.retryAtMs) return null;

  const nextBackoffMs =
    entry?.state === 'failed'
      ? Math.min(entry.backoffMs * 2, COMMAND_LINE_RETRY_MAX_BACKOFF_MS)
      : COMMAND_LINE_RETRY_INITIAL_BACKOFF_MS;
  processCommandLineCache.set(pid, { state: 'pending' });
  queryProcessCommandLine(pid, (commandLine) => {
    if (commandLine !== null) {
      processCommandLineCache.set(pid, {
        state: 'resolved',
        commandLine,
        expiresAtMs: Date.now() + COMMAND_LINE_CACHE_TTL_MS,
        refreshInFlight: false,
      });
    } else {
      processCommandLineCache.set(pid, {
        state: 'failed',
        retryAtMs: Date.now() + nextBackoffMs,
        backoffMs: nextBackoffMs,
      });
    }
  });
  return null;
}

export function findMpvWindows(targetSocketPath?: string | null): MpvPollResult {
  const foregroundHwnd = GetForegroundWindow();
  const matches: MpvWindowMatch[] = [];
  let hasMinimized = false;
  let hasFocused = false;
  pruneExpiredProcessNames(Date.now());

  const cb = koffi.register((hwnd: number, _lParam: number) => {
    if (!IsWindowVisible(hwnd)) return true;

    const pid = new Uint32Array(1);
    GetWindowThreadProcessId(hwnd, pid);
    const pidValue = pid[0]!;
    if (pidValue === 0) return true;

    const processName = getCachedProcessNameByPid(pidValue);
    if (!processName || processName.toLowerCase() !== 'mpv') return true;

    let commandLine: string | null = null;
    if (targetSocketPath) {
      commandLine = getProcessCommandLineByPid(pidValue);
      if (!commandLine || !matchesMpvSocketPathInCommandLine(commandLine, targetSocketPath)) {
        return true;
      }
    }

    if (IsIconic(hwnd)) {
      hasMinimized = true;
      return true;
    }

    const bounds = getWindowBounds(hwnd);
    if (!bounds) return true;

    const isForeground = foregroundHwnd !== 0 && hwnd === foregroundHwnd;
    if (isForeground) hasFocused = true;

    matches.push({
      hwnd,
      bounds,
      area: bounds.width * bounds.height,
      isForeground,
      commandLine,
    });

    return true;
  }, koffi.pointer(WNDENUMPROC));

  try {
    EnumWindows(cb, 0);
  } finally {
    koffi.unregister(cb);
  }

  return {
    matches,
    focusState: hasFocused,
    windowState: matches.length > 0 ? 'visible' : hasMinimized ? 'minimized' : 'not-found',
  };
}

export function getForegroundProcessName(): string | null {
  const foregroundHwnd = GetForegroundWindow();
  if (!foregroundHwnd) return null;

  const pid = new Uint32Array(1);
  GetWindowThreadProcessId(foregroundHwnd, pid);
  const pidValue = pid[0]!;
  if (pidValue === 0) return null;

  return getProcessNameByPid(pidValue);
}

export function setOverlayOwner(overlayHwnd: number, mpvHwnd: number): void {
  resetLastError();
  const result = SetWindowLongPtrW(overlayHwnd, GWLP_HWNDPARENT, mpvHwnd);
  assertSetWindowLongPtrSucceeded('setOverlayOwner', result);
  extendOverlayFrameIntoClientArea(overlayHwnd);
}

export function ensureOverlayTransparency(overlayHwnd: number): void {
  extendOverlayFrameIntoClientArea(overlayHwnd);
}

export function clearOverlayOwner(overlayHwnd: number): void {
  resetLastError();
  const result = SetWindowLongPtrW(overlayHwnd, GWLP_HWNDPARENT, 0);
  assertSetWindowLongPtrSucceeded('clearOverlayOwner', result);
}

export function bindOverlayAboveMpv(overlayHwnd: number, mpvHwnd: number): void {
  resetLastError();
  const ownerResult = SetWindowLongPtrW(overlayHwnd, GWLP_HWNDPARENT, mpvHwnd);
  assertSetWindowLongPtrSucceeded('bindOverlayAboveMpv owner assignment', ownerResult);

  resetLastError();
  const mpvExStyle = assertGetWindowLongSucceeded(
    'bindOverlayAboveMpv target window style',
    GetWindowLongW(mpvHwnd, GWL_EXSTYLE),
  );
  const mpvIsTopmost = (mpvExStyle & WS_EX_TOPMOST) !== 0;

  resetLastError();
  const overlayExStyle = assertGetWindowLongSucceeded(
    'bindOverlayAboveMpv overlay window style',
    GetWindowLongW(overlayHwnd, GWL_EXSTYLE),
  );
  const overlayIsTopmost = (overlayExStyle & WS_EX_TOPMOST) !== 0;

  if (mpvIsTopmost && !overlayIsTopmost) {
    resetLastError();
    const topmostResult = SetWindowPos(overlayHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_FLAGS);
    assertSetWindowPosSucceeded('bindOverlayAboveMpv topmost adjustment', topmostResult);
  } else if (!mpvIsTopmost && overlayIsTopmost) {
    resetLastError();
    const notTopmostResult = SetWindowPos(overlayHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_FLAGS);
    assertSetWindowPosSucceeded('bindOverlayAboveMpv notopmost adjustment', notTopmostResult);
  }

  const windowAboveMpv = GetWindowFn(mpvHwnd, GW_HWNDPREV);
  if (windowAboveMpv !== 0 && windowAboveMpv === overlayHwnd) return;

  let insertAfter = HWND_TOP;
  if (windowAboveMpv !== 0) {
    try {
      resetLastError();
      const aboveExStyle = assertGetWindowLongSucceeded(
        'bindOverlayAboveMpv window above style',
        GetWindowLongW(windowAboveMpv, GWL_EXSTYLE),
      );
      const aboveIsTopmost = (aboveExStyle & WS_EX_TOPMOST) !== 0;
      if (aboveIsTopmost === mpvIsTopmost) {
        insertAfter = windowAboveMpv;
      }
    } catch {
      insertAfter = HWND_TOP;
    }
  }

  resetLastError();
  const positionResult = SetWindowPos(overlayHwnd, insertAfter, 0, 0, 0, 0, SWP_FLAGS);
  assertSetWindowPosSucceeded('bindOverlayAboveMpv z-order adjustment', positionResult);
}

export function lowerOverlay(overlayHwnd: number): void {
  resetLastError();
  const notTopmostResult = SetWindowPos(overlayHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_FLAGS);
  assertSetWindowPosSucceeded('lowerOverlay notopmost adjustment', notTopmostResult);

  resetLastError();
  const bottomResult = SetWindowPos(overlayHwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_FLAGS);
  assertSetWindowPosSucceeded('lowerOverlay bottom adjustment', bottomResult);
}
