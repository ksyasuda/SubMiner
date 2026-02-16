import test from "node:test";
import assert from "node:assert/strict";
import {
  playNextSubtitleRuntimeService,
  replayCurrentSubtitleRuntimeService,
  sendMpvCommandRuntimeService,
  setMpvSubVisibilityRuntimeService,
  showMpvOsdRuntimeService,
} from "./mpv-service";

test("showMpvOsdRuntimeService sends show-text when connected", () => {
  const commands: (string | number)[][] = [];
  showMpvOsdRuntimeService(
    {
      connected: true,
      send: ({ command }) => {
        commands.push(command);
      },
    },
    "hello",
  );
  assert.deepEqual(commands, [["show-text", "hello", "3000"]]);
});

test("showMpvOsdRuntimeService logs fallback when disconnected", () => {
  const logs: string[] = [];
  showMpvOsdRuntimeService(
    {
      connected: false,
      send: () => {},
    },
    "hello",
    (line) => {
      logs.push(line);
    },
  );
  assert.deepEqual(logs, ["OSD (MPV not connected): hello"]);
});

test("mpv runtime command wrappers call expected client methods", () => {
  const calls: string[] = [];
  const client = {
    connected: true,
    send: ({ command }: { command: (string | number)[] }) => {
      calls.push(`send:${command.join(",")}`);
    },
    replayCurrentSubtitle: () => {
      calls.push("replay");
    },
    playNextSubtitle: () => {
      calls.push("next");
    },
    setSubVisibility: (visible: boolean) => {
      calls.push(`subVisible:${visible}`);
    },
  };

  replayCurrentSubtitleRuntimeService(client);
  playNextSubtitleRuntimeService(client);
  sendMpvCommandRuntimeService(client, ["script-message", "x"]);
  setMpvSubVisibilityRuntimeService(client, false);

  assert.deepEqual(calls, [
    "replay",
    "next",
    "send:script-message,x",
    "subVisible:false",
  ]);
});
