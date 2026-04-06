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

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import * as dbus from 'dbus-next';
import { BaseWindowTracker } from './base-tracker';
import { createLogger } from '../logger';

const log = createLogger('tracker').child('kwin');

const KWIN_SERVICE_NAME = 'org.kde.KWin';
const KWIN_SCRIPTING_PATH = '/Scripting';
const KWIN_SCRIPTING_INTERFACE = 'org.kde.kwin.Scripting';
const KWIN_SCRIPT_INTERFACE = 'org.kde.kwin.Script';
const BRIDGE_INTERFACE_NAME = 'io.github.subminer.kwinbridge.Interface';
const BRIDGE_OBJECT_PATH = '/io/github/subminer/kwinbridge';
const COMMAND_LINE_CACHE_TTL_MS = 1000;

type MessageBus = dbus.MessageBus;
type KWinLoadedScript = {
  scriptId: number;
  unloadKey: string;
};

export interface KWinWindow {
  active?: boolean;
  caption?: string;
  minimized?: boolean;
  normalWindow?: boolean;
  pid?: number;
  resourceClass?: string;
  resourceName?: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
}

interface KWinUpdatePayload {
  windows?: KWinWindow[];
}

interface SelectKWinMpvWindowOptions {
  targetMpvSocketPath: string | null;
  getWindowCommandLine: (pid: number) => string | null;
}

function createKWinTrackerInstanceToken(
  pid: number = process.pid,
  now: number = Date.now(),
  randomSuffix: string = Math.random().toString(36).slice(2, 10) || 'tracker',
): string {
  return `p${pid}_${now.toString(36)}_${randomSuffix}`;
}

export function buildKWinTrackerServiceName(instanceToken: string): string {
  return `io.github.subminer.kwinbridge.${instanceToken}`;
}

export function buildKWinTrackerPluginName(instanceToken: string): string {
  return `subminerKWinTracker_${instanceToken}`;
}

function shouldRetryUnnamedLoadScript(error: unknown): boolean {
  const message = error instanceof Error ? error.message : String(error ?? '');
  return (
    message.includes("Expected 1 body elements for signature 's'") ||
    message.includes('UnknownMethod')
  );
}

function normalizeWindowText(value: string | undefined): string {
  return value?.trim().toLowerCase() ?? '';
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function matchesTargetSocket(commandLine: string, targetMpvSocketPath: string): boolean {
  const escapedTarget = escapeRegExp(targetMpvSocketPath);
  const pattern = new RegExp(
    `(?:^|\\s)--input-ipc-server(?:=|\\s+)(?:"${escapedTarget}"|'${escapedTarget}'|${escapedTarget})(?=\\s|$)`,
  );
  return pattern.test(commandLine);
}

function preferActiveKWinWindow(windows: KWinWindow[]): KWinWindow | null {
  return windows.find((window) => window.active === true) ?? windows[0] ?? null;
}

function isMpvWindow(window: KWinWindow): boolean {
  return [window.resourceClass, window.resourceName, window.caption].some((value) =>
    normalizeWindowText(value).includes('mpv'),
  );
}

function hasValidGeometry(window: KWinWindow): boolean {
  return (
    Number.isFinite(window.x) &&
    Number.isFinite(window.y) &&
    Number.isFinite(window.width) &&
    Number.isFinite(window.height) &&
    (window.width ?? 0) > 0 &&
    (window.height ?? 0) > 0
  );
}

export function selectKWinMpvWindow(
  windows: KWinWindow[],
  options: SelectKWinMpvWindowOptions,
): KWinWindow | null {
  const visibleMpvWindows = windows.filter(
    (window) =>
      window.normalWindow !== false &&
      window.minimized !== true &&
      isMpvWindow(window) &&
      hasValidGeometry(window),
  );

  if (!options.targetMpvSocketPath) {
    return preferActiveKWinWindow(visibleMpvWindows);
  }

  const targetMpvSocketPath = options.targetMpvSocketPath;
  const matchingWindows = visibleMpvWindows.filter((window) => {
    if (!Number.isInteger(window.pid) || (window.pid ?? 0) <= 0) {
      return false;
    }

    const commandLine = options.getWindowCommandLine(window.pid!);
    if (!commandLine) {
      return false;
    }

    return matchesTargetSocket(commandLine, targetMpvSocketPath);
  });

  return preferActiveKWinWindow(matchingWindows);
}

function isKWinWindowCandidate(candidate: unknown): candidate is KWinWindow {
  return candidate !== null && typeof candidate === 'object';
}

class KWinTrackerBridgeInterface extends dbus.interface.Interface {
  private readonly onUpdatePayload: (payload: string) => void;

  constructor(onUpdatePayload: (payload: string) => void) {
    super(BRIDGE_INTERFACE_NAME);
    this.onUpdatePayload = onUpdatePayload;
  }

  Update(payload: string): void {
    this.onUpdatePayload(payload);
  }
}

KWinTrackerBridgeInterface.configureMembers({
  methods: {
    Update: {
      inSignature: 's',
      outSignature: '',
    },
  },
});

export function buildKWinBridgeScript(serviceName: string): string {
  return `
const SERVICE_NAME = ${JSON.stringify(serviceName)};
const OBJECT_PATH = ${JSON.stringify(BRIDGE_OBJECT_PATH)};
const INTERFACE_NAME = ${JSON.stringify(BRIDGE_INTERFACE_NAME)};
const OVERLAY_OWNER_PID = ${JSON.stringify(process.pid)};
const OVERLAY_WINDOW_CAPTION = "SubMiner";
const trackedWindows = new WeakSet();
const overlayKeepAboveState = [];
const overlayHiddenByScript = [];
let syncingPairState = false;
let pendingPairSync = false;
let pendingSyncTriggerWindow = null;
let pendingSyncTriggerEvent = "";

function isWatchableWindow(window) {
  try {
    if (!window || typeof window !== "object") {
      return false;
    }
    if (window.managed === false || window.deleted === true) {
      return false;
    }
    if (window.normalWindow !== true) {
      return false;
    }
    if (window.specialWindow === true || window.transient === true) {
      return false;
    }
    if (window.popupWindow === true || window.outline === true) {
      return false;
    }
  } catch (_error) {
    return false;
  }

  return true;
}

function isTrackableOverlayWindow(window) {
  if (!isWatchableWindow(window)) {
    return false;
  }

  if (Number(window.pid || 0) !== OVERLAY_OWNER_PID) {
    return false;
  }

  return String(window.caption || "") === OVERLAY_WINDOW_CAPTION;
}

function isMpvWindow(window) {
  if (!isWatchableWindow(window)) {
    return false;
  }

  const values = [window.resourceClass, window.resourceName, window.caption];
  for (const value of values) {
    if (String(value || "").toLowerCase().includes("mpv")) {
      return true;
    }
  }

  return false;
}

function isTrackableWindow(window) {
  return isMpvWindow(window) || isTrackableOverlayWindow(window);
}

function getWindowGeometry(window) {
  return window.clientGeometry || window.frameGeometry || {};
}

function serializeWindow(window) {
  const geometry = getWindowGeometry(window);
  return {
    active: window.active === true,
    caption: String(window.caption || ""),
    minimized: window.minimized === true,
    normalWindow: window.normalWindow === true,
    pid: Number(window.pid || 0),
    resourceClass: String(window.resourceClass || ""),
    resourceName: String(window.resourceName || ""),
    x: Number(geometry.x || 0),
    y: Number(geometry.y || 0),
    width: Number(geometry.width || 0),
    height: Number(geometry.height || 0),
  };
}

function windowRefIndex(entries, window) {
  for (let i = 0; i < entries.length; i += 1) {
    if (entries[i].window === window) {
      return i;
    }
  }
  return -1;
}

function ensureOverlayStateEntry(entries, window, value) {
  const index = windowRefIndex(entries, window);
  if (index >= 0) {
    return entries[index];
  }

  const entry = {
    window: window,
    value: value,
  };
  entries.push(entry);
  return entry;
}

function removeOverlayStateEntry(entries, window) {
  const index = windowRefIndex(entries, window);
  if (index >= 0) {
    entries.splice(index, 1);
  }
}

function readOverlayStateEntry(entries, window) {
  const index = windowRefIndex(entries, window);
  if (index >= 0) {
    return entries[index];
  }
  return null;
}

function pruneOverlayStateEntries(entries, windows) {
  for (let i = entries.length - 1; i >= 0; i -= 1) {
    const entry = entries[i];
    let found = false;
    for (let j = 0; j < windows.length; j += 1) {
      if (windows[j] === entry.window) {
        found = true;
        break;
      }
    }
    if (!found) {
      entries.splice(i, 1);
    }
  }
}

function hasUsableGeometry(window) {
  const geometry = getWindowGeometry(window);
  return (
    Number(geometry.width || 0) > 0 &&
    Number(geometry.height || 0) > 0
  );
}

function isWindowVisible(window) {
  if (!window) {
    return false;
  }
  if (window.minimized === true) {
    return false;
  }
  if (window.visible === false) {
    return false;
  }
  if (window.hidden === true) {
    return false;
  }
  return true;
}

function isWindowActive(window) {
  return window && (window.active === true || workspace.activeWindow === window);
}

function selectTargetMpvWindow() {
  const visibleMpvWindows = [];
  const hiddenMpvWindows = [];

  for (const window of workspace.windowList()) {
    if (!isMpvWindow(window) || !hasUsableGeometry(window)) {
      continue;
    }
    if (isWindowVisible(window)) {
      visibleMpvWindows.push(window);
    } else {
      hiddenMpvWindows.push(window);
    }
  }

  for (const window of visibleMpvWindows) {
    if (isWindowActive(window)) {
      return window;
    }
  }

  if (visibleMpvWindows.length > 0) {
    return visibleMpvWindows[0];
  }

  for (const window of hiddenMpvWindows) {
    if (isWindowActive(window)) {
      return window;
    }
  }

  return hiddenMpvWindows[0] || null;
}

function getOverlayWindows() {
  const windows = [];
  for (const window of workspace.windowList()) {
    if (isTrackableOverlayWindow(window)) {
      windows.push(window);
    }
  }
  windows.sort(function (left, right) {
    if (left.modal === true && right.modal !== true) {
      return 1;
    }
    if (left.modal !== true && right.modal === true) {
      return -1;
    }
    return 0;
  });
  return windows;
}

function rememberOverlayKeepAbove(window) {
  ensureOverlayStateEntry(overlayKeepAboveState, window, window.keepAbove === true);
}

function restoreOverlayKeepAbove(window) {
  const entry = readOverlayStateEntry(overlayKeepAboveState, window);
  if (!entry) {
    return;
  }
  try {
    window.keepAbove = entry.value === true;
  } catch (_error) {
    // ignore
  }
  removeOverlayStateEntry(overlayKeepAboveState, window);
}

function shouldRestoreOverlay(window) {
  const entry = readOverlayStateEntry(overlayHiddenByScript, window);
  return entry !== null && entry.value === true;
}

function markOverlayHiddenByScript(window, hidden) {
  if (hidden === true) {
    ensureOverlayStateEntry(overlayHiddenByScript, window, true);
    return;
  }
  removeOverlayStateEntry(overlayHiddenByScript, window);
}

function applyOverlayGeometry(overlayWindow, mpvWindow) {
  const targetGeometry = getWindowGeometry(mpvWindow);
  const currentGeometry = overlayWindow.frameGeometry || {};
  if (
    Number(currentGeometry.x || 0) === Number(targetGeometry.x || 0) &&
    Number(currentGeometry.y || 0) === Number(targetGeometry.y || 0) &&
    Number(currentGeometry.width || 0) === Number(targetGeometry.width || 0) &&
    Number(currentGeometry.height || 0) === Number(targetGeometry.height || 0)
  ) {
    return;
  }

  try {
    overlayWindow.frameGeometry = targetGeometry;
  } catch (_error) {
    // ignore
  }
}

function setOverlayKeepAbove(window, enabled) {
  if (enabled) {
    rememberOverlayKeepAbove(window);
    try {
      window.keepAbove = true;
    } catch (_error) {
      // ignore
    }
    return;
  }

  if (readOverlayStateEntry(overlayKeepAboveState, window)) {
    restoreOverlayKeepAbove(window);
  }
}

function restoreVisibleOverlayWindow(window) {
  if (!shouldRestoreOverlay(window)) {
    return;
  }

  try {
    window.minimized = false;
  } catch (_error) {
    // ignore
  }

  markOverlayHiddenByScript(window, false);
}

function hideVisibleOverlayWindow(window) {
  if (!isWindowVisible(window)) {
    return;
  }

  try {
    window.minimized = true;
    markOverlayHiddenByScript(window, true);
  } catch (_error) {
    // ignore
  }
}

function raiseWindow(window) {
  if (!window) {
    return;
  }

  try {
    workspace.raiseWindow(window);
  } catch (_error) {
    // ignore
  }
}

function activateWindow(window) {
  if (!window) {
    return;
  }

  try {
    workspace.activeWindow = window;
  } catch (_error) {
    // ignore
  }
}

function raiseWindowPair(mpvWindow, overlayWindows) {
  raiseWindow(mpvWindow);
  activateWindow(mpvWindow);

  let preferredOverlayWindow = null;
  for (const overlayWindow of overlayWindows) {
    if (!isWindowVisible(overlayWindow)) {
      continue;
    }
    raiseWindow(overlayWindow);
    preferredOverlayWindow = overlayWindow;
  }

  if (preferredOverlayWindow) {
    activateWindow(preferredOverlayWindow);
  }
}

function releaseOverlayPair(overlays) {
  for (const overlayWindow of overlays) {
    setOverlayKeepAbove(overlayWindow, false);
  }
}

function syncOverlayPairState(triggerWindow, triggerEvent) {
  const currentWindows = workspace.windowList();
  pruneOverlayStateEntries(overlayKeepAboveState, currentWindows);
  pruneOverlayStateEntries(overlayHiddenByScript, currentWindows);

  const mpvWindow = selectTargetMpvWindow();
  const overlayWindows = getOverlayWindows();

  if (!mpvWindow) {
    for (const overlayWindow of overlayWindows) {
      hideVisibleOverlayWindow(overlayWindow);
      setOverlayKeepAbove(overlayWindow, false);
    }
    return;
  }

  if (!isWindowVisible(mpvWindow)) {
    const overlayRequestedRestore =
      isTrackableOverlayWindow(triggerWindow) &&
      (
        triggerEvent === "windowShown" ||
        triggerEvent === "activeChanged" ||
        triggerEvent === "workspace-windowActivated" ||
        triggerEvent === "windowAdded"
      ) &&
      (isWindowVisible(triggerWindow) || isWindowActive(triggerWindow));

    if (overlayRequestedRestore) {
      try {
        mpvWindow.minimized = false;
      } catch (_error) {
        // ignore
      }
      raiseWindow(mpvWindow);
      activateWindow(mpvWindow);
    }
  }

  if (!isWindowVisible(mpvWindow)) {
    for (const overlayWindow of overlayWindows) {
      hideVisibleOverlayWindow(overlayWindow);
      setOverlayKeepAbove(overlayWindow, false);
    }
    return;
  }

  let pairActive = isWindowActive(mpvWindow);
  for (const overlayWindow of overlayWindows) {
    restoreVisibleOverlayWindow(overlayWindow);
    applyOverlayGeometry(overlayWindow, mpvWindow);
    if (isWindowActive(overlayWindow)) {
      pairActive = true;
    }
  }

  for (const overlayWindow of overlayWindows) {
    setOverlayKeepAbove(overlayWindow, pairActive && isWindowVisible(overlayWindow));
  }

  if (pairActive) {
    raiseWindowPair(mpvWindow, overlayWindows);
  }
}

function runStateSync(triggerWindow, triggerEvent) {
  if (syncingPairState) {
    pendingPairSync = true;
    pendingSyncTriggerWindow = triggerWindow || pendingSyncTriggerWindow;
    pendingSyncTriggerEvent = triggerEvent || pendingSyncTriggerEvent;
    return;
  }

  const nextTriggerWindow = triggerWindow || pendingSyncTriggerWindow;
  const nextTriggerEvent = triggerEvent || pendingSyncTriggerEvent;
  pendingSyncTriggerWindow = null;
  pendingSyncTriggerEvent = "";
  syncingPairState = true;
  try {
    syncOverlayPairState(nextTriggerWindow, nextTriggerEvent);
    emitState();
  } finally {
    syncingPairState = false;
  }

  if (pendingPairSync) {
    pendingPairSync = false;
    runStateSync(pendingSyncTriggerWindow, pendingSyncTriggerEvent);
  }
}

function emitState() {
  const windows = [];
  for (const window of workspace.windowList()) {
    if (isMpvWindow(window)) {
      windows.push(serializeWindow(window));
    }
  }
  callDBus(
    SERVICE_NAME,
    OBJECT_PATH,
    INTERFACE_NAME,
    "Update",
    JSON.stringify({ windows })
  );
}

function watchWindow(window) {
  if (!isTrackableWindow(window) || trackedWindows.has(window)) {
    return;
  }

  trackedWindows.add(window);
  if (window.closed) {
    window.closed.connect(function () {
      runStateSync(window, "closed");
    });
  }
  if (window.frameGeometryChanged) {
    window.frameGeometryChanged.connect(function () {
      runStateSync(window, "frameGeometryChanged");
    });
  }
  if (window.clientGeometryChanged) {
    window.clientGeometryChanged.connect(function () {
      runStateSync(window, "clientGeometryChanged");
    });
  }
  if (window.outputChanged) {
    window.outputChanged.connect(function () {
      runStateSync(window, "outputChanged");
    });
  }
  if (window.windowClassChanged) {
    window.windowClassChanged.connect(function () {
      runStateSync(window, "windowClassChanged");
    });
  }
  if (window.windowShown) {
    window.windowShown.connect(function () {
      runStateSync(window, "windowShown");
    });
  }
  if (window.windowHidden) {
    window.windowHidden.connect(function () {
      runStateSync(window, "windowHidden");
    });
  }
  if (window.activeChanged) {
    window.activeChanged.connect(function () {
      runStateSync(window, "activeChanged");
    });
  }
}

function refresh() {
  for (const window of workspace.windowList()) {
    watchWindow(window);
  }
  runStateSync();
}

workspace.windowAdded.connect(function (window) {
  watchWindow(window);
  runStateSync(window, "windowAdded");
});
workspace.windowRemoved.connect(function () {
  runStateSync(null, "windowRemoved");
});
workspace.windowActivated.connect(function (window) {
  watchWindow(window);
  runStateSync(window, "workspace-windowActivated");
});
workspace.screensChanged.connect(function () {
  runStateSync(null, "screensChanged");
});

refresh();
`;
}

export class KWinWindowTracker extends BaseWindowTracker {
  private readonly targetMpvSocketPath: string | null;
  private readonly serviceName: string;
  private readonly pluginName: string;
  private readonly bridgeInterface: KWinTrackerBridgeInterface;
  private tempDir: string | null = null;
  private scriptPath: string | null = null;
  private bus: MessageBus | null = null;
  private scriptId: number | null = null;
  private unloadScriptKey: string | null = null;
  private readonly commandLineCache = new Map<number, { expiresAt: number; value: string | null }>();
  private stopped = false;

  constructor(targetMpvSocketPath?: string) {
    super();
    const instanceToken = createKWinTrackerInstanceToken();
    this.targetMpvSocketPath = targetMpvSocketPath?.trim() || null;
    this.serviceName = buildKWinTrackerServiceName(instanceToken);
    this.pluginName = buildKWinTrackerPluginName(instanceToken);
    this.bridgeInterface = new KWinTrackerBridgeInterface((payload) => this.handleUpdate(payload));
  }

  start(): void {
    this.stopped = false;
    void this.startAsync();
  }

  stop(): void {
    this.stopped = true;
    void this.stopAsync();
  }

  private async startAsync(): Promise<void> {
    try {
      const scriptPath = this.ensureScriptWorkspace();
      fs.writeFileSync(scriptPath, buildKWinBridgeScript(this.serviceName), 'utf-8');
      const bus = dbus.sessionBus();
      bus.on('error', (error) => {
        log.error('KWin session bus error:', (error as Error).message);
      });
      this.bus = bus;
      await bus.requestName(this.serviceName, 0);
      bus.export(BRIDGE_OBJECT_PATH, this.bridgeInterface);

      const loadedScript = await this.loadScript(bus, scriptPath, this.pluginName);
      this.scriptId = loadedScript.scriptId;
      this.unloadScriptKey = loadedScript.unloadKey;
      await this.runScript(bus, loadedScript.scriptId);

      if (this.stopped) {
        await this.stopAsync();
      }
    } catch (error) {
      log.error('Failed to start KWin window tracker:', (error as Error).message);
      this.updateGeometry(null);
      await this.stopAsync();
    }
  }

  private async stopAsync(): Promise<void> {
    const bus = this.bus;
    const scriptId = this.scriptId;
    const unloadScriptKey = this.unloadScriptKey;
    const tempDir = this.tempDir;
    this.scriptId = null;
    this.bus = null;
    this.unloadScriptKey = null;
    this.tempDir = null;
    this.scriptPath = null;
    this.commandLineCache.clear();

    if (bus && scriptId !== null) {
      try {
        await this.stopScript(bus, scriptId);
      } catch {
        // ignore
      }
    }

    if (bus && unloadScriptKey) {
      try {
        await this.unloadScript(bus, unloadScriptKey);
      } catch {
        // ignore
      }
    }

    if (bus) {
      try {
        bus.unexport(BRIDGE_OBJECT_PATH, this.bridgeInterface);
      } catch {
        // ignore
      }

      try {
        await bus.releaseName(this.serviceName);
      } catch {
        // ignore
      }

      bus.disconnect();
    }

    try {
      if (tempDir) {
        fs.rmSync(tempDir, { recursive: true, force: true });
      }
    } catch {
      // ignore
    }
  }

  private handleUpdate(payload: string): void {
    const parsed = this.parsePayload(payload);
    if (!parsed) {
      this.updateGeometry(null);
      return;
    }

    const windows = Array.isArray(parsed.windows) ? parsed.windows.filter(isKWinWindowCandidate) : [];
    const targetWindow = selectKWinMpvWindow(windows, {
      targetMpvSocketPath: this.targetMpvSocketPath,
      getWindowCommandLine: (pid) => this.getWindowCommandLine(pid),
    });

    if (!targetWindow) {
      this.updateGeometry(null);
      return;
    }

    this.updateGeometry({
      x: targetWindow.x ?? 0,
      y: targetWindow.y ?? 0,
      width: targetWindow.width ?? 0,
      height: targetWindow.height ?? 0,
    });
    this.updateFocus(targetWindow.active === true);
  }

  private parsePayload(payload: string): KWinUpdatePayload | null {
    try {
      const parsed = JSON.parse(payload) as unknown;
      if (!parsed || typeof parsed !== 'object') {
        return null;
      }
      return parsed as KWinUpdatePayload;
    } catch {
      return null;
    }
  }

  private getWindowCommandLine(pid: number): string | null {
    const cached = this.commandLineCache.get(pid);
    const now = Date.now();
    if (cached && cached.expiresAt > now) {
      return cached.value;
    }

    const commandLine = this.readProcessCommandLine(pid);
    this.commandLineCache.set(pid, {
      expiresAt: now + COMMAND_LINE_CACHE_TTL_MS,
      value: commandLine,
    });
    return commandLine;
  }

  private readProcessCommandLine(pid: number): string | null {
    if (!Number.isInteger(pid) || pid <= 0) {
      return null;
    }

    const safePid = String(pid);
    if (process.platform === 'linux') {
      try {
        const commandLine = fs
          .readFileSync(`/proc/${safePid}/cmdline`, 'utf-8')
          .replace(/\0/g, ' ')
          .trim();
        return commandLine || null;
      } catch {
        // fall through to ps for environments without /proc access
      }
    }

    try {
      const commandLine = execFileSync('ps', ['-p', safePid, '-o', 'args='], {
        encoding: 'utf-8',
      }).trim();
      return commandLine || null;
    } catch {
      return null;
    }
  }

  private ensureScriptWorkspace(): string {
    if (!this.tempDir || !this.scriptPath) {
      this.tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-kwin-'));
      this.scriptPath = path.join(this.tempDir, 'main.js');
    }

    return this.scriptPath;
  }

  private async loadScript(
    bus: MessageBus,
    filePath: string,
    pluginName: string,
  ): Promise<KWinLoadedScript> {
    try {
      const scriptId = await this.callMethod<number>(bus, {
        path: KWIN_SCRIPTING_PATH,
        interfaceName: KWIN_SCRIPTING_INTERFACE,
        member: 'loadScript',
        signature: 'ss',
        body: [filePath, pluginName],
      });
      return { scriptId, unloadKey: pluginName };
    } catch (error) {
      if (!shouldRetryUnnamedLoadScript(error)) {
        throw error;
      }

      log.warn('KWin named loadScript overload failed; retrying unnamed loadScript call.');
      const scriptId = await this.callMethod<number>(bus, {
        path: KWIN_SCRIPTING_PATH,
        interfaceName: KWIN_SCRIPTING_INTERFACE,
        member: 'loadScript',
        signature: 's',
        body: [filePath],
      });
      return { scriptId, unloadKey: filePath };
    }
  }

  private async unloadScript(bus: MessageBus, pluginName: string): Promise<boolean> {
    return this.callMethod<boolean>(bus, {
      path: KWIN_SCRIPTING_PATH,
      interfaceName: KWIN_SCRIPTING_INTERFACE,
      member: 'unloadScript',
      signature: 's',
      body: [pluginName],
    });
  }

  private async runScript(bus: MessageBus, scriptId: number): Promise<void> {
    await this.callMethod<void>(bus, {
      path: `${KWIN_SCRIPTING_PATH}/Script${scriptId}`,
      interfaceName: KWIN_SCRIPT_INTERFACE,
      member: 'run',
    });
  }

  private async stopScript(bus: MessageBus, scriptId: number): Promise<void> {
    await this.callMethod<void>(bus, {
      path: `${KWIN_SCRIPTING_PATH}/Script${scriptId}`,
      interfaceName: KWIN_SCRIPT_INTERFACE,
      member: 'stop',
    });
  }

  private async callMethod<T>(
    bus: MessageBus,
    options: {
      path: string;
      interfaceName: string;
      member: string;
      signature?: string;
      body?: unknown[];
    },
  ): Promise<T> {
    const reply = await bus.call(
      new dbus.Message({
        destination: KWIN_SERVICE_NAME,
        path: options.path,
        interface: options.interfaceName,
        member: options.member,
        signature: options.signature ?? '',
        body: options.body ?? [],
      }),
    );

    if (!reply) {
      throw new Error(`No reply received from ${options.interfaceName}.${options.member}`);
    }

    const values = (reply.body ?? []) as T[];
    if (values.length === 0 && options.signature) {
      throw new Error(`Empty reply body from ${options.interfaceName}.${options.member}`);
    }
    return values[0] as T;
  }
}
