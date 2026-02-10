import * as http from "http";
import * as https from "https";
import * as path from "path";
import * as childProcess from "child_process";
import {
  JimakuApiResponse,
  JimakuConfig,
  JimakuMediaInfo,
} from "../types";

function execCommand(
  command: string,
): Promise<{ stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    childProcess.exec(command, { timeout: 10000 }, (err, stdout, stderr) => {
      if (err) {
        reject(err);
        return;
      }
      resolve({ stdout, stderr });
    });
  });
}

export async function resolveJimakuApiKey(
  config: JimakuConfig,
): Promise<string | null> {
  if (config.apiKey && config.apiKey.trim()) {
    console.log("[jimaku] API key found in config");
    return config.apiKey.trim();
  }
  if (config.apiKeyCommand && config.apiKeyCommand.trim()) {
    try {
      const { stdout } = await execCommand(config.apiKeyCommand);
      const key = stdout.trim();
      console.log(
        `[jimaku] apiKeyCommand result: ${key.length > 0 ? "key obtained" : "empty output"}`,
      );
      return key.length > 0 ? key : null;
    } catch (err) {
      console.error(
        "Failed to run jimaku.apiKeyCommand:",
        (err as Error).message,
      );
      return null;
    }
  }
  console.log(
    "[jimaku] No API key configured (neither apiKey nor apiKeyCommand set)",
  );
  return null;
}

function getRetryAfter(headers: http.IncomingHttpHeaders): number | undefined {
  const value = headers["x-ratelimit-reset-after"];
  if (!value) return undefined;
  const raw = Array.isArray(value) ? value[0] : value;
  const parsed = Number.parseFloat(raw);
  if (!Number.isFinite(parsed)) return undefined;
  return parsed;
}

export async function jimakuFetchJson<T>(
  endpoint: string,
  query: Record<string, string | number | boolean | null | undefined>,
  options: { baseUrl: string; apiKey: string },
): Promise<JimakuApiResponse<T>> {
  const url = new URL(endpoint, options.baseUrl);
  for (const [key, value] of Object.entries(query)) {
    if (value === null || value === undefined) continue;
    url.searchParams.set(key, String(value));
  }

  console.log(`[jimaku] GET ${url.toString()}`);
  const transport = url.protocol === "https:" ? https : http;

  return new Promise((resolve) => {
    const req = transport.request(
      url,
      {
        method: "GET",
        headers: {
          Authorization: options.apiKey,
          "User-Agent": "SubMiner",
        },
      },
      (res) => {
        let data = "";
        res.on("data", (chunk) => {
          data += chunk.toString();
        });
        res.on("end", () => {
          const status = res.statusCode || 0;
          console.log(`[jimaku] Response HTTP ${status} for ${endpoint}`);
          if (status >= 200 && status < 300) {
            try {
              const parsed = JSON.parse(data) as T;
              resolve({ ok: true, data: parsed });
            } catch {
              console.error(`[jimaku] JSON parse error: ${data.slice(0, 200)}`);
              resolve({
                ok: false,
                error: { error: "Failed to parse Jimaku response JSON." },
              });
            }
            return;
          }

          let errorMessage = `Jimaku API error (HTTP ${status})`;
          try {
            const parsed = JSON.parse(data) as { error?: string };
            if (parsed && parsed.error) {
              errorMessage = parsed.error;
            }
          } catch {
            // Ignore parse errors.
          }
          console.error(`[jimaku] API error: ${errorMessage}`);

          resolve({
            ok: false,
            error: {
              error: errorMessage,
              code: status || undefined,
              retryAfter:
                status === 429 ? getRetryAfter(res.headers) : undefined,
            },
          });
        });
      },
    );

    req.on("error", (err) => {
      console.error(`[jimaku] Network error: ${(err as Error).message}`);
      resolve({
        ok: false,
        error: { error: `Jimaku request failed: ${(err as Error).message}` },
      });
    });

    req.end();
  });
}

function matchEpisodeFromName(name: string): {
  season: number | null;
  episode: number | null;
  index: number | null;
  confidence: "high" | "medium" | "low";
} {
  const seasonEpisode = name.match(/S(\d{1,2})E(\d{1,3})/i);
  if (seasonEpisode && seasonEpisode.index !== undefined) {
    return {
      season: Number.parseInt(seasonEpisode[1], 10),
      episode: Number.parseInt(seasonEpisode[2], 10),
      index: seasonEpisode.index,
      confidence: "high",
    };
  }

  const alt = name.match(/(\d{1,2})x(\d{1,3})/i);
  if (alt && alt.index !== undefined) {
    return {
      season: Number.parseInt(alt[1], 10),
      episode: Number.parseInt(alt[2], 10),
      index: alt.index,
      confidence: "high",
    };
  }

  const epOnly = name.match(/(?:^|[\s._-])E(?:P)?(\d{1,3})(?:\b|[\s._-])/i);
  if (epOnly && epOnly.index !== undefined) {
    return {
      season: null,
      episode: Number.parseInt(epOnly[1], 10),
      index: epOnly.index,
      confidence: "medium",
    };
  }

  const numeric = name.match(/(?:^|[-–—]\s*)(\d{1,3})\s*[-–—]/);
  if (numeric && numeric.index !== undefined) {
    return {
      season: null,
      episode: Number.parseInt(numeric[1], 10),
      index: numeric.index,
      confidence: "medium",
    };
  }

  return { season: null, episode: null, index: null, confidence: "low" };
}

function detectSeasonFromDir(mediaPath: string): number | null {
  const parent = path.basename(path.dirname(mediaPath));
  const match = parent.match(/(?:Season|S)\s*(\d{1,2})/i);
  if (!match) return null;
  const parsed = Number.parseInt(match[1], 10);
  return Number.isFinite(parsed) ? parsed : null;
}

function cleanupTitle(value: string): string {
  return value
    .replace(/^[\s-–—]+/, "")
    .replace(/[\s-–—]+$/, "")
    .replace(/\s+/g, " ")
    .trim();
}

export function parseMediaInfo(mediaPath: string | null): JimakuMediaInfo {
  if (!mediaPath) {
    return {
      title: "",
      season: null,
      episode: null,
      confidence: "low",
      filename: "",
      rawTitle: "",
    };
  }

  const filename = path.basename(mediaPath);
  let name = filename.replace(/\.[^/.]+$/, "");
  name = name.replace(/\[[^\]]*]/g, " ");
  name = name.replace(/\(\d{4}\)/g, " ");
  name = name.replace(/[._]/g, " ");
  name = name.replace(/[–—]/g, "-");
  name = name.replace(/\s+/g, " ").trim();

  const parsed = matchEpisodeFromName(name);
  let titlePart = name;
  if (parsed.index !== null) {
    titlePart = name.slice(0, parsed.index);
  }

  const seasonFromDir = parsed.season ?? detectSeasonFromDir(mediaPath);
  const title = cleanupTitle(titlePart || name);

  return {
    title,
    season: seasonFromDir,
    episode: parsed.episode,
    confidence: parsed.confidence,
    filename,
    rawTitle: name,
  };
}
