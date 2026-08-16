export type ModalDescriptor<TId extends string> = {
  id: TId;
  isOpen: () => boolean;
  close: () => void;
  suppressesSubtitles: boolean;
};

export function createModalRegistry<TId extends string>(
  descriptors: readonly ModalDescriptor<TId>[],
) {
  return {
    isAnyOpen: (): boolean => descriptors.some((descriptor) => descriptor.isOpen()),
    isAnySuppressingSubtitlesOpen: (): boolean =>
      descriptors.some((descriptor) => descriptor.suppressesSubtitles && descriptor.isOpen()),
    getActive: (): TId | null => descriptors.find((descriptor) => descriptor.isOpen())?.id ?? null,
    dismissOpen: (): void => {
      for (const descriptor of descriptors) {
        if (descriptor.isOpen()) {
          descriptor.close();
        }
      }
    },
    dismissOpenExcept: (id: TId): void => {
      for (const descriptor of descriptors) {
        if (descriptor.id !== id && descriptor.isOpen()) {
          descriptor.close();
        }
      }
    },
  };
}
