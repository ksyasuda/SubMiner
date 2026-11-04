import { spawn as spawnProcess } from 'node:child_process';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as readline from 'node:readline';
import { isCompiledMacOSHelperCurrent } from '../../window-trackers/macos-tracker';

// macOS only delivers GameController input to the frontmost app, so the overlay's Gamepad API
// sees nothing while mpv is in front. `scripts/macos-controller-watch.swift` opts in to
// background events and reports each button press; SubMiner then focuses the overlay (after
// which Chromium reads the pad itself) and replays that first press there.

const HELPER_NAME = 'macos-controller-watch';

export type MacOSControllerWatchCommand = { command: string; args: string[] };

/**
 * Finds the helper: the packaged binary in Resources, a compiled build in dist/scripts that is
 * newer than its source, or the Swift source run through `swift` in a dev checkout.
 */
export function resolveMacOSControllerWatchCommand(
  options: {
    resourcesPath: string | undefined;
    /** Directory holding the built main process (dist/ in a checkout). */
    mainDirname: string;
  },
  helperFs: Pick<typeof fs, 'existsSync' | 'statSync'> = fs,
): MacOSControllerWatchCommand | null {
  const sourcePath = path.join(options.mainDirname, '..', 'scripts', `${HELPER_NAME}.swift`);
  if (options.resourcesPath) {
    const packagedPath = path.join(options.resourcesPath, 'scripts', HELPER_NAME);
    if (helperFs.existsSync(packagedPath)) return { command: packagedPath, args: [] };
  }
  const builtPath = path.join(options.mainDirname, 'scripts', HELPER_NAME);
  if (isCompiledMacOSHelperCurrent(builtPath, sourcePath, helperFs)) {
    return { command: builtPath, args: [] };
  }
  if (helperFs.existsSync(sourcePath)) return { command: 'swift', args: [sourcePath] };
  return null;
}

export type ControllerWatchPress = { buttonIndex: number; controllerCount: number };

/** Parses a helper line ("press <standard-mapping button index> <connected pads>"). */
export function parseControllerWatchLine(line: string): ControllerWatchPress | null {
  const match = /^press (\d+) (\d+)$/.exec(line.trim());
  return match ? { buttonIndex: Number(match[1]), controllerCount: Number(match[2]) } : null;
}

type ControllerWatchLogger = {
  info: (message: string, ...args: unknown[]) => void;
  warn: (message: string, ...args: unknown[]) => void;
};

/** Runs the helper until `stop` is called; the helper also exits when SubMiner does. */
export function startMacOSControllerWatch(options: {
  command: MacOSControllerWatchCommand;
  onPress: (press: ControllerWatchPress) => void;
  logger: ControllerWatchLogger;
  spawn?: typeof spawnProcess;
}): { stop: () => void } {
  const spawn = options.spawn ?? spawnProcess;
  const child = spawn(options.command.command, options.command.args, {
    stdio: ['pipe', 'pipe', 'pipe'],
  });
  let stopped = false;

  readline.createInterface({ input: child.stdout }).on('line', (line) => {
    const press = parseControllerWatchLine(line);
    if (press) options.onPress(press);
  });
  readline.createInterface({ input: child.stderr }).on('line', (line) => {
    options.logger.warn(`[${HELPER_NAME}] ${line}`);
  });
  child.on('error', (error) => {
    options.logger.warn(`Failed to start ${HELPER_NAME}`, error);
  });
  child.on('exit', (code, signal) => {
    if (!stopped) options.logger.warn(`${HELPER_NAME} exited`, { code, signal });
  });
  options.logger.info(`Started ${HELPER_NAME}: ${options.command.command}`);

  return {
    stop: () => {
      stopped = true;
      child.stdin?.end();
      child.kill();
    },
  };
}

type FocusableWindow = {
  isDestroyed: () => boolean;
  isVisible: () => boolean;
};

/**
 * A press the overlay missed because mpv was in front: focus the overlay window that owns the
 * controller (a visible modal, else the main overlay) and replay the press in it. Presses made
 * while another app is in front are left alone, so a pad used elsewhere is not hijacked. With
 * several pads connected the press only focuses: it might not be from the pad the overlay reads.
 */
export function handleBackgroundControllerPress<TWindow extends FocusableWindow>(
  press: ControllerWatchPress,
  deps: {
    isControllerEnabled: () => boolean;
    isMpvFocused: () => boolean;
    getModalWindow: () => TWindow | null;
    getMainWindow: () => TWindow | null;
    focusWindow: (window: TWindow) => void;
    replayPress: (window: TWindow, buttonIndex: number) => void;
  },
): void {
  if (!deps.isControllerEnabled() || !deps.isMpvFocused()) return;
  const target = [deps.getModalWindow(), deps.getMainWindow()].find(
    (window): window is TWindow => window !== null && !window.isDestroyed() && window.isVisible(),
  );
  if (!target) return;
  deps.focusWindow(target);
  if (press.controllerCount === 1) deps.replayPress(target, press.buttonIndex);
}
