import test from "node:test";
import assert from "node:assert/strict";
import { resolveCurrentAudioStreamIndex } from "./mpv-service";

test("resolveCurrentAudioStreamIndex returns selected ff-index when no current track id", () => {
  assert.equal(
    resolveCurrentAudioStreamIndex(
      [
        { type: "audio", id: 1, selected: false, "ff-index": 1 },
        { type: "audio", id: 2, selected: true, "ff-index": 3 },
      ],
      null,
    ),
    3,
  );
});

test("resolveCurrentAudioStreamIndex prefers matching current audio track id", () => {
  assert.equal(
    resolveCurrentAudioStreamIndex(
      [
        { type: "audio", id: 1, selected: true, "ff-index": 3 },
        { type: "audio", id: 2, selected: false, "ff-index": 6 },
      ],
      2,
    ),
    6,
  );
});

test("resolveCurrentAudioStreamIndex returns null when tracks are not an array", () => {
  assert.equal(
    resolveCurrentAudioStreamIndex(null, null),
    null,
  );
});
