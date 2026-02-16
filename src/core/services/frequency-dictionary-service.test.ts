import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

import { createFrequencyDictionaryLookupService } from "./frequency-dictionary-service";

test("createFrequencyDictionaryLookupService logs parse errors and returns no-op for invalid dictionaries", async () => {
  const logs: string[] = [];
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "subminer-frequency-dict-"));
  const bankPath = path.join(tempDir, "term_meta_bank_1.json");
  fs.writeFileSync(bankPath, "{ invalid json");

  const lookup = await createFrequencyDictionaryLookupService({
    searchPaths: [tempDir],
    log: (message) => {
      logs.push(message);
    },
  });

  const rank = lookup("猫");

  assert.equal(rank, null);
  assert.equal(
    logs.some((entry) =>
      entry.includes("Failed to parse frequency dictionary file as JSON") &&
      entry.includes("term_meta_bank_1.json")
    ),
    true,
  );
});

test("createFrequencyDictionaryLookupService continues with no-op lookup when search path is missing", async () => {
  const logs: string[] = [];
  const missingPath = path.join(os.tmpdir(), "subminer-frequency-dict-missing-dir");
  const lookup = await createFrequencyDictionaryLookupService({
    searchPaths: [missingPath],
    log: (message) => {
      logs.push(message);
    },
  });

  assert.equal(lookup("猫"), null);
  assert.equal(
    logs.some((entry) => entry.includes(`Frequency dictionary not found.`)),
    true,
  );
});
