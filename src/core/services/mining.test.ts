import test from "node:test";
import assert from "node:assert/strict";
import {
  copyCurrentSubtitle,
  handleMineSentenceDigit,
  handleMultiCopyDigit,
  mineSentenceCard,
} from "./mining";

test("copyCurrentSubtitle reports tracker and subtitle guards", () => {
  const osd: string[] = [];
  const copied: string[] = [];

  copyCurrentSubtitle({
    subtitleTimingTracker: null,
    writeClipboardText: (text) => copied.push(text),
    showMpvOsd: (text) => osd.push(text),
  });
  assert.equal(osd.at(-1), "Subtitle tracker not available");

  copyCurrentSubtitle({
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

test("copyCurrentSubtitle copies current subtitle text", () => {
  const osd: string[] = [];
  const copied: string[] = [];

  copyCurrentSubtitle({
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

test("mineSentenceCard handles missing integration and disconnected mpv", async () => {
  const osd: string[] = [];

  assert.equal(
    await mineSentenceCard({
    ankiIntegration: null,
    mpvClient: null,
    showMpvOsd: (text) => osd.push(text),
    }),
    false,
  );
  assert.equal(osd.at(-1), "AnkiConnect integration not enabled");

  assert.equal(
    await mineSentenceCard({
    ankiIntegration: {
      updateLastAddedFromClipboard: async () => {},
      triggerFieldGroupingForLastAddedCard: async () => {},
      markLastCardAsAudioCard: async () => {},
      createSentenceCard: async () => false,
    },
    mpvClient: {
      connected: false,
      currentSubText: "line",
      currentSubStart: 1,
      currentSubEnd: 2,
    },
    showMpvOsd: (text) => osd.push(text),
    }),
    false,
  );

  assert.equal(osd.at(-1), "MPV not connected");
});

test("mineSentenceCard creates sentence card from mpv subtitle state", async () => {
  const created: Array<{
    sentence: string;
    startTime: number;
    endTime: number;
    secondarySub?: string;
  }> = [];

  const createdCard = await mineSentenceCard({
    ankiIntegration: {
      updateLastAddedFromClipboard: async () => {},
      triggerFieldGroupingForLastAddedCard: async () => {},
      markLastCardAsAudioCard: async () => {},
      createSentenceCard: async (sentence, startTime, endTime, secondarySub) => {
        created.push({ sentence, startTime, endTime, secondarySub });
        return true;
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

  assert.equal(createdCard, true);
  assert.deepEqual(created, [
    {
      sentence: "subtitle line",
      startTime: 10,
      endTime: 12,
      secondarySub: "secondary line",
    },
  ]);
});

test("handleMultiCopyDigit copies available history and reports truncation", () => {
  const osd: string[] = [];
  const copied: string[] = [];

  handleMultiCopyDigit(5, {
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

test("handleMineSentenceDigit reports async create failures", async () => {
  const osd: string[] = [];
  const logs: Array<{ message: string; err: unknown }> = [];
  let cardsMined = 0;

  handleMineSentenceDigit(2, {
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
    onCardsMined: (count) => {
      cardsMined += count;
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  assert.equal(logs.length, 1);
  assert.equal(logs[0]?.message, "mineSentenceMultiple failed:");
  assert.equal((logs[0]?.err as Error).message, "mine boom");
  assert.ok(osd.some((entry) => entry.includes("Mine sentence failed: mine boom")));
  assert.equal(cardsMined, 0);
});

test("handleMineSentenceDigitService increments successful card count", async () => {
  const osd: string[] = [];
  let cardsMined = 0;

  handleMineSentenceDigit(2, {
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
      createSentenceCard: async () => true,
    },
    getCurrentSecondarySubText: () => "sub2",
    showMpvOsd: (text) => osd.push(text),
    logError: () => {},
    onCardsMined: (count) => {
      cardsMined += count;
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  assert.equal(cardsMined, 1);
});
