/*
 * SubMiner - All-in-one sentence mining overlay
 * Copyright (C) 2024 sudacode
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

import { spawn, execSync } from "child_process";
import { PartOfSpeech, Token, MecabStatus } from "./types";

export { PartOfSpeech };

function mapPartOfSpeech(pos1: string): PartOfSpeech {
  switch (pos1) {
    case "名詞":
      return PartOfSpeech.noun;
    case "動詞":
      return PartOfSpeech.verb;
    case "形容詞":
      return PartOfSpeech.i_adjective;
    case "形状詞":
    case "形容動詞":
      return PartOfSpeech.na_adjective;
    case "助詞":
      return PartOfSpeech.particle;
    case "助動詞":
      return PartOfSpeech.bound_auxiliary;
    case "記号":
    case "補助記号":
      return PartOfSpeech.symbol;
    default:
      return PartOfSpeech.other;
  }
}

export function parseMecabLine(line: string): Token | null {
  if (!line || line === "EOS" || line.trim() === "") {
    return null;
  }

  const tabIndex = line.indexOf("\t");
  if (tabIndex === -1) {
    return null;
  }

  const surface = line.substring(0, tabIndex);
  const featureString = line.substring(tabIndex + 1);
  const features = featureString.split(",");

  const pos1 = features[0] || "";
  const pos2 = features[1] || "";
  const pos3 = features[2] || "";
  const pos4 = features[3] || "";
  const inflectionType = features[4] || "";
  const inflectionForm = features[5] || "";
  const lemma = features[6] || surface;
  const reading = features[7] || "";
  const pronunciation = features[8] || "";

  return {
    word: surface,
    partOfSpeech: mapPartOfSpeech(pos1),
    pos1,
    pos2,
    pos3,
    pos4,
    inflectionType,
    inflectionForm,
    headword: lemma !== "*" ? lemma : surface,
    katakanaReading: reading !== "*" ? reading : "",
    pronunciation: pronunciation !== "*" ? pronunciation : "",
  };
}

export class MecabTokenizer {
  private mecabPath: string | null = null;
  private available: boolean = false;
  private enabled: boolean = true;

  async checkAvailability(): Promise<boolean> {
    try {
      const result = execSync("which mecab", { encoding: "utf-8" }).trim();
      if (result) {
        this.mecabPath = result;
        this.available = true;
        console.log("MeCab found at:", this.mecabPath);
        return true;
      }
    } catch (err) {
      console.log("MeCab not found on system");
    }

    this.available = false;
    return false;
  }

  async tokenize(text: string): Promise<Token[] | null> {
    if (!this.available || !this.enabled || !text) {
      return null;
    }

    return new Promise((resolve) => {
      const mecab = spawn("mecab", [], {
        stdio: ["pipe", "pipe", "pipe"],
      });

      let stdout = "";
      let stderr = "";

      mecab.stdout.on("data", (data: Buffer) => {
        stdout += data.toString();
      });

      mecab.stderr.on("data", (data: Buffer) => {
        stderr += data.toString();
      });

      mecab.on("close", (code: number | null) => {
        if (code !== 0) {
          console.error("MeCab process exited with code:", code);
          if (stderr) {
            console.error("MeCab stderr:", stderr);
          }
          resolve(null);
          return;
        }

        const lines = stdout.split("\n");
        const tokens: Token[] = [];

        for (const line of lines) {
          const token = parseMecabLine(line);
          if (token) {
            tokens.push(token);
          }
        }

        resolve(tokens);
      });

      mecab.on("error", (err: Error) => {
        console.error("Failed to spawn MeCab:", err.message);
        resolve(null);
      });

      mecab.stdin.write(text);
      mecab.stdin.end();
    });
  }

  getStatus(): MecabStatus {
    return {
      available: this.available,
      enabled: this.enabled,
      path: this.mecabPath,
    };
  }

  setEnabled(enabled: boolean): void {
    this.enabled = enabled;
  }
}

export { mapPartOfSpeech };
