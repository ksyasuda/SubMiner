import { execFile as nodeExecFile } from 'node:child_process';

const FOCUS_MPV_PROCESS_SCRIPT =
  'tell application "System Events" to set frontmost of the first process whose name is "mpv" to true';

type ExecFileForMacOSFocus = (
  command: string,
  args: string[],
  options: { timeout: number },
  callback: (error: Error | null) => void,
) => void;

export type MacOSMpvFocusDeps = {
  execFile?: ExecFileForMacOSFocus;
};

export async function focusMacOSMpvProcess(deps: MacOSMpvFocusDeps = {}): Promise<void> {
  const execFile: ExecFileForMacOSFocus =
    deps.execFile ??
    ((command, args, options, callback) => {
      nodeExecFile(command, args, options, (error) => {
        callback(error);
      });
    });

  await new Promise<void>((resolve) => {
    execFile('/usr/bin/osascript', ['-e', FOCUS_MPV_PROCESS_SCRIPT], { timeout: 2000 }, () => {
      resolve();
    });
  });
}
