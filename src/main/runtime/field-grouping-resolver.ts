import type { KikuFieldGroupingChoice } from '../../types';

type FieldGroupingResolver = ((choice: KikuFieldGroupingChoice) => void) | null;

export function createGetFieldGroupingResolverHandler(deps: {
  getResolver: () => FieldGroupingResolver;
}) {
  return (): FieldGroupingResolver => deps.getResolver();
}

export function createSetFieldGroupingResolverHandler(deps: {
  setResolver: (resolver: FieldGroupingResolver) => void;
  nextSequence: () => number;
  getSequence: () => number;
}) {
  return (resolver: FieldGroupingResolver): void => {
    if (!resolver) {
      deps.setResolver(null);
      return;
    }

    const sequence = deps.nextSequence();
    const wrappedResolver = (choice: KikuFieldGroupingChoice): void => {
      if (sequence !== deps.getSequence()) return;
      resolver(choice);
    };
    deps.setResolver(wrappedResolver);
  };
}
