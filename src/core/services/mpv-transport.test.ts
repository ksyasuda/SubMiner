import test from "node:test";
import assert from "node:assert/strict";
import { getMpvReconnectDelay } from "./mpv-transport";

test("getMpvReconnectDelay follows existing reconnect ramp", () => {
  assert.equal(getMpvReconnectDelay(0, true), 1000);
  assert.equal(getMpvReconnectDelay(1, true), 1000);
  assert.equal(getMpvReconnectDelay(2, true), 2000);
  assert.equal(getMpvReconnectDelay(4, true), 5000);
  assert.equal(getMpvReconnectDelay(7, true), 10000);

  assert.equal(getMpvReconnectDelay(0, false), 200);
  assert.equal(getMpvReconnectDelay(2, false), 500);
  assert.equal(getMpvReconnectDelay(4, false), 1000);
  assert.equal(getMpvReconnectDelay(6, false), 2000);
});
