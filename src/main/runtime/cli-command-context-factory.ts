import { createCliCommandContext } from './cli-command-context';
import { createBuildCliCommandContextDepsHandler } from './cli-command-context-deps';
import { createBuildCliCommandContextMainDepsHandler } from './cli-command-context-main-deps';

type CliCommandContextMainDeps = Parameters<typeof createBuildCliCommandContextMainDepsHandler>[0];

export function createCliCommandContextFactory(deps: CliCommandContextMainDeps) {
  const buildCliCommandContextMainDepsHandler = createBuildCliCommandContextMainDepsHandler(deps);
  const cliCommandContextMainDeps = buildCliCommandContextMainDepsHandler();
  const buildCliCommandContextDepsHandler =
    createBuildCliCommandContextDepsHandler(cliCommandContextMainDeps);

  return () => createCliCommandContext(buildCliCommandContextDepsHandler());
}
