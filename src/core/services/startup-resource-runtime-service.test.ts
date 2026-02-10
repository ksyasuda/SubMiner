import test from "node:test";
import assert from "node:assert/strict";
import {
  createMecabTokenizerAndCheckRuntimeService,
  createSubtitleTimingTrackerRuntimeService,
} from "./startup-resource-runtime-service";

test("createMecabTokenizerAndCheckRuntimeService sets tokenizer and checks availability", async () => {
  const calls: string[] = [];
  let assigned: unknown = null;
  await createMecabTokenizerAndCheckRuntimeService({
    createMecabTokenizer: () => ({
      checkAvailability: async () => {
        calls.push("checkAvailability");
      },
    }),
    setMecabTokenizer: (tokenizer) => {
      assigned = tokenizer;
      calls.push("setMecabTokenizer");
    },
  });
  assert.equal(assigned !== null, true);
  assert.deepEqual(calls, ["setMecabTokenizer", "checkAvailability"]);
});

test("createSubtitleTimingTrackerRuntimeService sets created tracker", () => {
  const tracker = { id: "x" };
  let assigned: unknown = null;
  createSubtitleTimingTrackerRuntimeService({
    createSubtitleTimingTracker: () => tracker,
    setSubtitleTimingTracker: (value) => {
      assigned = value;
    },
  });
  assert.equal(assigned, tracker);
});
