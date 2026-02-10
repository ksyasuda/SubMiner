import test from "node:test";
import assert from "node:assert/strict";
import {
  createNumericShortcutRuntimeService,
  createNumericShortcutSessionService,
} from "./numeric-shortcut-service";

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

test("numeric shortcut session handles digit selection and unregisters shortcuts", () => {
  const handlers = new Map<string, () => void>();
  const unregistered: string[] = [];
  const osd: string[] = [];
  const session = createNumericShortcutSessionService({
    registerShortcut: (accelerator, handler) => {
      handlers.set(accelerator, handler);
      return true;
    },
    unregisterShortcut: (accelerator) => {
      unregistered.push(accelerator);
      handlers.delete(accelerator);
    },
    setTimer: () => setTimeout(() => {}, 0),
    clearTimer: (timer) => clearTimeout(timer),
    showMpvOsd: (text) => {
      osd.push(text);
    },
  });

  const digits: number[] = [];
  session.start({
    timeoutMs: 5000,
    onDigit: (digit) => {
      digits.push(digit);
    },
    messages: {
      prompt: "Pick a digit",
      timeout: "Timed out",
    },
  });

  assert.equal(session.isActive(), true);
  assert.equal(osd[0], "Pick a digit");
  assert.ok(handlers.has("3"));
  handlers.get("3")?.();

  assert.deepEqual(digits, [3]);
  assert.equal(session.isActive(), false);
  assert.ok(unregistered.includes("Escape"));
  assert.ok(unregistered.includes("1"));
  assert.ok(unregistered.includes("9"));
});

test("numeric shortcut session emits timeout message", () => {
  const osd: string[] = [];
  const session = createNumericShortcutSessionService({
    registerShortcut: () => true,
    unregisterShortcut: () => {},
    setTimer: (handler) => {
      handler();
      return setTimeout(() => {}, 0);
    },
    clearTimer: (timer) => clearTimeout(timer),
    showMpvOsd: (text) => {
      osd.push(text);
    },
  });

  session.start({
    timeoutMs: 5000,
    onDigit: () => {},
    messages: {
      prompt: "Pick a digit",
      timeout: "Timed out",
      cancelled: "Aborted",
    },
  });

  assert.equal(session.isActive(), false);
  assert.ok(osd.includes("Timed out"));
});

test("numeric shortcut session handles escape cancellation", () => {
  const handlers = new Map<string, () => void>();
  const osd: string[] = [];
  const session = createNumericShortcutSessionService({
    registerShortcut: (accelerator, handler) => {
      handlers.set(accelerator, handler);
      return true;
    },
    unregisterShortcut: (accelerator) => {
      handlers.delete(accelerator);
    },
    setTimer: () => setTimeout(() => {}, 10000),
    clearTimer: (timer) => clearTimeout(timer),
    showMpvOsd: (text) => {
      osd.push(text);
    },
  });

  session.start({
    timeoutMs: 5000,
    onDigit: () => {},
    messages: {
      prompt: "Pick a digit",
      timeout: "Timed out",
      cancelled: "Aborted",
    },
  });
  handlers.get("Escape")?.();
  assert.equal(session.isActive(), false);
  assert.ok(osd.includes("Aborted"));
});
