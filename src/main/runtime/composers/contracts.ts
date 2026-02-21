type ComposerShape = Record<string, unknown>;

export type ComposerInputs<T extends ComposerShape> = Readonly<Required<T>>;

export type ComposerOutputs<T extends ComposerShape> = Readonly<T>;

export type BuiltMainDeps<TFactory> = TFactory extends (
  ...args: infer _TFactoryArgs
) => infer TBuilder
  ? TBuilder extends (...args: infer _TBuilderArgs) => infer TDeps
    ? TDeps
    : never
  : never;
