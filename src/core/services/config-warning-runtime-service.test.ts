import test from "node:test";
import assert from "node:assert/strict";
import {
  formatConfigWarningRuntimeService,
  logConfigWarningRuntimeService,
} from "./config-warning-runtime-service";

test("formatConfigWarningRuntimeService formats warning line", () => {
  const message = formatConfigWarningRuntimeService({
    path: "ankiConnect.enabled",
    value: "oops",
    fallback: true,
    message: "invalid type",
  });
  assert.equal(
    message,
    '[config] ankiConnect.enabled: invalid type value="oops" fallback=true',
  );
});

test("logConfigWarningRuntimeService delegates to logger", () => {
  const logs: string[] = [];
  logConfigWarningRuntimeService(
    {
      path: "x.y",
      value: 1,
      fallback: 2,
      message: "bad",
    },
    (line) => logs.push(line),
  );
  assert.equal(logs.length, 1);
  assert.match(logs[0], /^\[config\] x\.y: bad /);
});
