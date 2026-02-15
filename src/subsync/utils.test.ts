import test from "node:test";
import assert from "node:assert/strict";
import { codecToExtension } from "./utils";

test("codecToExtension maps stream/web formats to ffmpeg extractable extensions", () => {
  assert.equal(codecToExtension("subrip"), "srt");
  assert.equal(codecToExtension("webvtt"), "vtt");
  assert.equal(codecToExtension("vtt"), "vtt");
  assert.equal(codecToExtension("ttml"), "ttml");
});

test("codecToExtension returns null for unsupported codecs", () => {
  assert.equal(codecToExtension("unsupported-codec"), null);
});
