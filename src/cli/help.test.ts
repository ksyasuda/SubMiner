import test from "node:test";
import assert from "node:assert/strict";
import { printHelp } from "./help";

test("printHelp includes configured texthooker port", () => {
  const original = console.log;
  let output = "";
  console.log = (value?: unknown) => {
    output += String(value);
  };

  try {
    printHelp(7777);
  } finally {
    console.log = original;
  }

  assert.match(output, /--help\s+Show this help/);
  assert.match(output, /default: 7777/);
  assert.match(output, /--refresh-known-words/);
  assert.match(output, /--anilist-status/);
  assert.match(output, /--anilist-retry-queue/);
});
