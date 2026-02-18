import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";

import type { MergedToken } from "../types";
import { PartOfSpeech } from "../types.js";
import { computeWordClass } from "./subtitle-render.js";

function createToken(overrides: Partial<MergedToken>): MergedToken {
  return {
    surface: "",
    reading: "",
    headword: "",
    startPos: 0,
    endPos: 0,
    partOfSpeech: PartOfSpeech.other,
    isMerged: true,
    isKnown: false,
    isNPlusOneTarget: false,
    ...overrides,
  };
}

function extractClassBlock(cssText: string, selector: string): string {
  const ruleRegex = /([^{}]+)\{([^}]*)\}/g;
  let match: RegExpExecArray | null = null;
  let fallbackBlock = "";

  while ((match = ruleRegex.exec(cssText)) !== null) {
    const selectorsBlock = match[1]?.trim() ?? "";
    const selectorBlock = match[2] ?? "";

    const selectors = selectorsBlock
      .split(",")
      .map((entry) => entry.trim())
      .filter((entry) => entry.length > 0);

    if (selectors.includes(selector)) {
      if (selectors.length === 1) {
        return selectorBlock;
      }

      if (!fallbackBlock) {
        fallbackBlock = selectorBlock;
      }
    }
  }

  if (fallbackBlock) {
    return fallbackBlock;
  }

  return "";
}

test("computeWordClass preserves known and n+1 classes while adding JLPT classes", () => {
  const knownJlpt = createToken({
    isKnown: true,
    jlptLevel: "N1",
    surface: "猫",
  });
  const nPlusOneJlpt = createToken({
    isNPlusOneTarget: true,
    jlptLevel: "N2",
    surface: "犬",
  });

  assert.equal(computeWordClass(knownJlpt), "word word-known word-jlpt-n1");
  assert.equal(
    computeWordClass(nPlusOneJlpt),
    "word word-n-plus-one word-jlpt-n2",
  );
});

test("computeWordClass does not add frequency class to known or N+1 terms", () => {
  const known = createToken({
    isKnown: true,
    frequencyRank: 10,
    surface: "既知",
  });
  const nPlusOne = createToken({
    isNPlusOneTarget: true,
    frequencyRank: 10,
    surface: "目標",
  });
  const frequency = createToken({
    frequencyRank: 10,
    surface: "頻度",
  });

  assert.equal(
    computeWordClass(known, {
      enabled: true,
      topX: 100,
      mode: "single",
      singleColor: "#000000",
      bandedColors: [
        "#000000",
        "#000000",
        "#000000",
        "#000000",
        "#000000",
      ] as const,
    }),
    "word word-known",
  );
  assert.equal(
    computeWordClass(nPlusOne, {
      enabled: true,
      topX: 100,
      mode: "single",
      singleColor: "#000000",
      bandedColors: [
        "#000000",
        "#000000",
        "#000000",
        "#000000",
        "#000000",
      ] as const,
    }),
    "word word-n-plus-one",
  );
  assert.equal(
    computeWordClass(frequency, {
      enabled: true,
      topX: 100,
      mode: "single",
      singleColor: "#000000",
      bandedColors: [
        "#000000",
        "#000000",
        "#000000",
        "#000000",
        "#000000",
      ] as const,
    }),
    "word word-frequency-single",
  );
});

test("computeWordClass adds frequency class for single mode when rank is within topX", () => {
  const token = createToken({
    surface: "猫",
    frequencyRank: 50,
  });

  const actual = computeWordClass(token, {
    enabled: true,
    topX: 100,
    mode: "single",
    singleColor: "#000000",
    bandedColors: [
      "#000000",
      "#000000",
      "#000000",
      "#000000",
      "#000000",
    ] as const,
  });

  assert.equal(actual, "word word-frequency-single");
});

test("computeWordClass adds frequency class when rank equals topX", () => {
  const token = createToken({
    surface: "水",
    frequencyRank: 100,
  });

  const actual = computeWordClass(token, {
    enabled: true,
    topX: 100,
    mode: "single",
    singleColor: "#000000",
    bandedColors: [
      "#000000",
      "#000000",
      "#000000",
      "#000000",
      "#000000",
    ] as const,
  });

  assert.equal(actual, "word word-frequency-single");
});

test("computeWordClass adds frequency class for banded mode", () => {
  const token = createToken({
    surface: "犬",
    frequencyRank: 250,
  });

  const actual = computeWordClass(token, {
    enabled: true,
    topX: 1000,
    mode: "banded",
    singleColor: "#000000",
    bandedColors: [
      "#111111",
      "#222222",
      "#333333",
      "#444444",
      "#555555",
    ] as const,
  });

  assert.equal(actual, "word word-frequency-band-2");
});

test("computeWordClass uses configured band count for banded mode", () => {
  const token = createToken({
    surface: "犬",
    frequencyRank: 2,
  });

  const actual = computeWordClass(token, {
    enabled: true,
    topX: 4,
    mode: "banded",
    singleColor: "#000000",
    bandedColors: ["#111111", "#222222", "#333333", "#444444", "#555555"],
  } as any);

  assert.equal(actual, "word word-frequency-band-3");
});

test("computeWordClass skips frequency class when rank is out of topX", () => {
  const token = createToken({
    surface: "犬",
    frequencyRank: 1200,
  });

  const actual = computeWordClass(token, {
    enabled: true,
    topX: 1000,
    mode: "single",
    singleColor: "#000000",
    bandedColors: [
      "#000000",
      "#000000",
      "#000000",
      "#000000",
      "#000000",
    ] as const,
  });

  assert.equal(actual, "word");
});

test("JLPT CSS rules use underline-only styling in renderer stylesheet", () => {
  const distCssPath = path.join(process.cwd(), "dist", "renderer", "style.css");
  const srcCssPath = path.join(process.cwd(), "src", "renderer", "style.css");

  const cssPath = fs.existsSync(distCssPath) ? distCssPath : srcCssPath;
  if (!fs.existsSync(cssPath)) {
    assert.fail(
      "JLPT CSS file missing. Run `pnpm run build` first, or ensure src/renderer/style.css exists.",
    );
  }

  const cssText = fs.readFileSync(cssPath, "utf-8");

  for (let level = 1; level <= 5; level += 1) {
    const block = extractClassBlock(
      cssText,
      `#subtitleRoot .word.word-jlpt-n${level}`,
    );
    assert.ok(block.length > 0, `word-jlpt-n${level} class should exist`);
    assert.match(block, /text-decoration-line:\s*underline;/);
    assert.match(block, /text-decoration-thickness:\s*2px;/);
    assert.match(block, /text-underline-offset:\s*4px;/);
    assert.match(block, /color:\s*inherit;/);
  }

  for (let band = 1; band <= 5; band += 1) {
    const block = extractClassBlock(
      cssText,
      band === 1
        ? "#subtitleRoot .word.word-frequency-single"
        : `#subtitleRoot .word.word-frequency-band-${band}`,
    );
    assert.ok(
      block.length > 0,
      `frequency class word-frequency-${band === 1 ? "single" : `band-${band}`} should exist`,
    );
    assert.match(block, /color:\s*var\(/);
  }
});
