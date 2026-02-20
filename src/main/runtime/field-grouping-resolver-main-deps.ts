import type {
  createGetFieldGroupingResolverHandler,
  createSetFieldGroupingResolverHandler,
} from './field-grouping-resolver';

type GetFieldGroupingResolverMainDeps = Parameters<typeof createGetFieldGroupingResolverHandler>[0];
type SetFieldGroupingResolverMainDeps = Parameters<typeof createSetFieldGroupingResolverHandler>[0];

export function createBuildGetFieldGroupingResolverMainDepsHandler(
  deps: GetFieldGroupingResolverMainDeps,
) {
  return (): GetFieldGroupingResolverMainDeps => ({
    getResolver: () => deps.getResolver(),
  });
}

export function createBuildSetFieldGroupingResolverMainDepsHandler(
  deps: SetFieldGroupingResolverMainDeps,
) {
  return (): SetFieldGroupingResolverMainDeps => ({
    setResolver: (resolver) => deps.setResolver(resolver),
    nextSequence: () => deps.nextSequence(),
    getSequence: () => deps.getSequence(),
  });
}
