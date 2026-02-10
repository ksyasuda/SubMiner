import test from "node:test";
import assert from "node:assert/strict";
import { createNumericShortcutRuntimeService } from "./numeric-shortcut-runtime-service";

test("createNumericShortcutRuntimeService creates sessions wired to globalShortcut", () => {
  const registered: string[] = [];
  const unregistered: string[] = [];
  const osd: string[] = [];
  const handlers = new Map<string, () => void>();

  const runtime = createNumericShortcutRuntimeService({
    globalShortcut: {
      register: (accelerator, callback) => {
        registered.push(accelerator);
        handlers.set(accelerator, callback);
        return true;
      },
      unregister: (accelerator) => {
        unregistered.push(accelerator);
        handlers.delete(accelerator);
      },
    },
    showMpvOsd: (text) => {
      osd.push(text);
    },
    setTimer: () => setTimeout(() => {}, 1000),
    clearTimer: (timer) => clearTimeout(timer),
  });

  const session = runtime.createSession();
  session.start({
    timeoutMs: 5000,
    onDigit: () => {},
    messages: {
      prompt: "Select count",
      timeout: "Timed out",
    },
  });

  assert.equal(session.isActive(), true);
  assert.ok(registered.includes("1"));
  assert.ok(registered.includes("Escape"));
  assert.equal(osd[0], "Select count");

  handlers.get("Escape")?.();
  assert.equal(session.isActive(), false);
  assert.ok(unregistered.includes("Escape"));
});
