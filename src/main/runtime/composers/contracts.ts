type ComposerShape = Record<string, unknown>;

export type ComposerInputs<T extends ComposerShape> = Readonly<Required<T>>;

export type ComposerOutputs<T extends ComposerShape> = Readonly<T>;
