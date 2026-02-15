import test from "node:test";
import assert from "node:assert/strict";
import {
  copyCurrentSubtitleService,
  handleMineSentenceDigitService,
  handleMultiCopyDigitService,
  mineSentenceCardService,
} from "./mining-service";

test("copyCurrentSubtitleService reports tracker and subtitle guards", () => {
  const osd: string[] = [];
  const copied: string[] = [];

  copyCurrentSubtitleService({
    subtitleTimingTracker: null,
    writeClipboardText: (text) => copied.push(text),
    showMpvOsd: (text) => osd.push(text),
  });
  assert.equal(osd.at(-1), "Subtitle tracker not available");

  copyCurrentSubtitleService({
    subtitleTimingTracker: {
      getRecentBlocks: () => [],
      getCurrentSubtitle: () => null,
      findTiming: () => null,
    },
    writeClipboardText: (text) => copied.push(text),
    showMpvOsd: (text) => osd.push(text),
  });
  assert.equal(osd.at(-1), "No current subtitle");
  assert.deepEqual(copied, []);
});

test("copyCurrentSubtitleService copies current subtitle text", () => {
  const osd: string[] = [];
  const copied: string[] = [];

  copyCurrentSubtitleService({
    subtitleTimingTracker: {
      getRecentBlocks: () => [],
      getCurrentSubtitle: () => "hello world",
      findTiming: () => null,
    },
    writeClipboardText: (text) => copied.push(text),
    showMpvOsd: (text) => osd.push(text),
  });

  assert.deepEqual(copied, ["hello world"]);
  assert.equal(osd.at(-1), "Copied subtitle");
});

test("mineSentenceCardService handles missing integration and disconnected mpv", async () => {
  const osd: string[] = [];

  await mineSentenceCardService({
    ankiIntegration: null,
    mpvClient: null,
    showMpvOsd: (text) => osd.push(text),
  });
  assert.equal(osd.at(-1), "AnkiConnect integration not enabled");

  await mineSentenceCardService({
    ankiIntegration: {
      updateLastAddedFromClipboard: async () => {},
      triggerFieldGroupingForLastAddedCard: async () => {},
      markLastCardAsAudioCard: async () => {},
      createSentenceCard: async () => {},
    },
    mpvClient: {
      connected: false,
      currentSubText: "line",
      currentSubStart: 1,
      currentSubEnd: 2,
    },
    showMpvOsd: (text) => osd.push(text),
  });

  assert.equal(osd.at(-1), "MPV not connected");
});

test("mineSentenceCardService creates sentence card from mpv subtitle state", async () => {
  const created: Array<{
    sentence: string;
    startTime: number;
    endTime: number;
    secondarySub?: string;
  }> = [];

  await mineSentenceCardService({
    ankiIntegration: {
      updateLastAddedFromClipboard: async () => {},
      triggerFieldGroupingForLastAddedCard: async () => {},
      markLastCardAsAudioCard: async () => {},
      createSentenceCard: async (sentence, startTime, endTime, secondarySub) => {
        created.push({ sentence, startTime, endTime, secondarySub });
      },
    },
    mpvClient: {
      connected: true,
      currentSubText: "subtitle line",
      currentSubStart: 10,
      currentSubEnd: 12,
      currentSecondarySubText: "secondary line",
    },
    showMpvOsd: () => {},
  });

  assert.deepEqual(created, [
    {
      sentence: "subtitle line",
      startTime: 10,
      endTime: 12,
      secondarySub: "secondary line",
    },
  ]);
});

test("handleMultiCopyDigitService copies available history and reports truncation", () => {
  const osd: string[] = [];
  const copied: string[] = [];

  handleMultiCopyDigitService(5, {
    subtitleTimingTracker: {
      getRecentBlocks: (count) => ["a", "b"].slice(0, count),
      getCurrentSubtitle: () => null,
      findTiming: () => null,
    },
    writeClipboardText: (text) => copied.push(text),
    showMpvOsd: (text) => osd.push(text),
  });

  assert.deepEqual(copied, ["a\n\nb"]);
  assert.equal(osd.at(-1), "Only 2 lines available, copied 2");
});

test("handleMineSentenceDigitService reports async create failures", async () => {
  const osd: string[] = [];
  const logs: Array<{ message: string; err: unknown }> = [];

  handleMineSentenceDigitService(2, {
    subtitleTimingTracker: {
      getRecentBlocks: () => ["one", "two"],
      getCurrentSubtitle: () => null,
      findTiming: (text) =>
        text === "one"
          ? { startTime: 1, endTime: 3 }
          : { startTime: 4, endTime: 7 },
    },
    ankiIntegration: {
      updateLastAddedFromClipboard: async () => {},
      triggerFieldGroupingForLastAddedCard: async () => {},
      markLastCardAsAudioCard: async () => {},
      createSentenceCard: async () => {
        throw new Error("mine boom");
      },
    },
    getCurrentSecondarySubText: () => "sub2",
    showMpvOsd: (text) => osd.push(text),
    logError: (message, err) => logs.push({ message, err }),
  });

  await new Promise((resolve) => setImmediate(resolve));

  assert.equal(logs.length, 1);
  assert.equal(logs[0]?.message, "mineSentenceMultiple failed:");
  assert.equal((logs[0]?.err as Error).message, "mine boom");
  assert.ok(osd.some((entry) => entry.includes("Mine sentence failed: mine boom")));
});
