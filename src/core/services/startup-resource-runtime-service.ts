interface MecabTokenizerLike {
  checkAvailability: () => Promise<unknown>;
}

interface SubtitleTimingTrackerLike {}

export async function createMecabTokenizerAndCheckRuntimeService<
  T extends MecabTokenizerLike,
>(options: {
  createMecabTokenizer: () => T;
  setMecabTokenizer: (tokenizer: T) => void;
}): Promise<void> {
  const tokenizer = options.createMecabTokenizer();
  options.setMecabTokenizer(tokenizer);
  await tokenizer.checkAvailability();
}

export function createSubtitleTimingTrackerRuntimeService<
  T extends SubtitleTimingTrackerLike,
>(options: {
  createSubtitleTimingTracker: () => T;
  setSubtitleTimingTracker: (tracker: T) => void;
}): void {
  const tracker = options.createSubtitleTimingTracker();
  options.setSubtitleTimingTracker(tracker);
}
