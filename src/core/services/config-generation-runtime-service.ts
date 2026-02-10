import { CliArgs } from "../../cli/args";

export interface ConfigGenerationRuntimeDeps {
  shouldStartApp: (args: CliArgs) => boolean;
  generateConfig: (args: CliArgs) => Promise<number>;
  onSuccess: (exitCode: number) => void;
  onError: (error: Error) => void;
}

export function runGenerateConfigFlowRuntimeService(
  args: CliArgs,
  deps: ConfigGenerationRuntimeDeps,
): boolean {
  if (!args.generateConfig || deps.shouldStartApp(args)) {
    return false;
  }

  deps.generateConfig(args)
    .then((exitCode) => {
      deps.onSuccess(exitCode);
    })
    .catch((error: Error) => {
      deps.onError(error);
    });
  return true;
}
