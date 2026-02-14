import { handleCliCommandService, createCliCommandDepsRuntimeService } from "../core/services";
import type { CliArgs, CliCommandSource } from "../cli/args";
import { createCliCommandRuntimeServiceDeps, CliCommandRuntimeServiceDepsParams } from "./dependencies";

export function handleCliCommandRuntimeService(
  args: CliArgs,
  source: CliCommandSource,
  params: CliCommandRuntimeServiceDepsParams,
): void {
  const deps = createCliCommandDepsRuntimeService(
    createCliCommandRuntimeServiceDeps(params),
  );
  handleCliCommandService(args, source, deps);
}

