import test from "node:test";
import assert from "node:assert/strict";
import { createAppLoggingRuntimeService } from "./app-logging-runtime-service";

test("createAppLoggingRuntimeService routes logs and formats config warnings", () => {
  const lines: string[] = [];
  const logger = {
    log: (line: string) => lines.push(`log:${line}`),
    warn: (line: string) => lines.push(`warn:${line}`),
    error: (line: string) => lines.push(`error:${line}`),
  };

  const runtime = createAppLoggingRuntimeService(logger);
  runtime.logInfo("hello");
  runtime.logWarning("careful");
  runtime.logNoRunningInstance();
  runtime.logConfigWarning({
    path: "x.y",
    value: "bad",
    fallback: "good",
    message: "invalid",
  });

  assert.equal(lines[0], "log:hello");
  assert.equal(lines[1], "warn:careful");
  assert.equal(lines[2], "error:No running instance. Use --start to launch the app.");
  assert.match(lines[3], /^warn:\[config\] x\.y: invalid /);
});
