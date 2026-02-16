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

function extractClassBlock(cssText: string, level: number): string {
  const selector = `#subtitleRoot .word.word-jlpt-n${level}`;
  const start = cssText.indexOf(selector);
  if (start < 0) return "";

  const openBrace = cssText.indexOf("{", start);
  if (openBrace < 0) return "";
  const closeBrace = cssText.indexOf("}", openBrace);
  if (closeBrace < 0) return "";

  return cssText.slice(openBrace + 1, closeBrace);
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

test("JLPT CSS rules use underline-only styling in renderer stylesheet", () => {
  const distCssPath = path.join(process.cwd(), "dist", "renderer", "style.css");
  const srcCssPath = path.join(process.cwd(), "src", "renderer", "style.css");

  const cssPath = fs.existsSync(distCssPath)
    ? distCssPath
    : srcCssPath;
  if (!fs.existsSync(cssPath)) {
    assert.fail(
      "JLPT CSS file missing. Run `pnpm run build` first, or ensure src/renderer/style.css exists.",
    );
  }

  const cssText = fs.readFileSync(cssPath, "utf-8");

  for (let level = 1; level <= 5; level += 1) {
    const block = extractClassBlock(cssText, level);
    assert.ok(block.length > 0, `word-jlpt-n${level} class should exist`);
    assert.match(block, /text-decoration-line:\s*underline;/);
    assert.match(block, /text-decoration-thickness:\s*2px;/);
    assert.match(block, /text-underline-offset:\s*4px;/);
    assert.match(block, /color:\s*inherit;/);
  }
});
