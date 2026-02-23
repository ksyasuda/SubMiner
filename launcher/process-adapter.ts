export interface ProcessAdapter {
  platform(): NodeJS.Platform;
  onSignal(signal: NodeJS.Signals, handler: () => void): void;
  writeStdout(text: string): void;
  exit(code: number): never;
  setExitCode(code: number): void;
}

export const nodeProcessAdapter: ProcessAdapter = {
  platform: () => process.platform,
  onSignal: (signal, handler) => {
    process.on(signal, handler);
  },
  writeStdout: (text) => {
    process.stdout.write(text);
  },
  exit: (code) => process.exit(code),
  setExitCode: (code) => {
    process.exitCode = code;
  },
};
