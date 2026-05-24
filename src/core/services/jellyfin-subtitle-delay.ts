import * as fs from 'fs';
import * as path from 'path';

type JellyfinSubtitleDelayStore = {
  version?: unknown;
  delays?: unknown;
};

type JellyfinSubtitleDelayParams = {
  filePath: string;
  itemId: string;
  streamIndex: number;
};

type SaveJellyfinSubtitleDelayParams = JellyfinSubtitleDelayParams & {
  delaySeconds: number;
};

function storeKey(itemId: string, streamIndex: number): string {
  return JSON.stringify([itemId, streamIndex]);
}

function readDelayMap(filePath: string): Record<string, number> {
  try {
    if (!fs.existsSync(filePath)) return {};
    const parsed = JSON.parse(fs.readFileSync(filePath, 'utf-8')) as JellyfinSubtitleDelayStore;
    if (
      !parsed ||
      typeof parsed !== 'object' ||
      !parsed.delays ||
      typeof parsed.delays !== 'object'
    ) {
      return {};
    }
    const delays: Record<string, number> = {};
    for (const [key, value] of Object.entries(parsed.delays as Record<string, unknown>)) {
      if (typeof value === 'number' && Number.isFinite(value)) {
        delays[key] = value;
      }
    }
    return delays;
  } catch {
    return {};
  }
}

export function loadJellyfinSubtitleDelay(params: JellyfinSubtitleDelayParams): number | null {
  const delay = readDelayMap(params.filePath)[storeKey(params.itemId, params.streamIndex)];
  return typeof delay === 'number' && Number.isFinite(delay) ? delay : null;
}

export function saveJellyfinSubtitleDelay(params: SaveJellyfinSubtitleDelayParams): boolean {
  if (!Number.isFinite(params.delaySeconds)) return false;
  try {
    const delays = readDelayMap(params.filePath);
    delays[storeKey(params.itemId, params.streamIndex)] = params.delaySeconds;
    const dir = path.dirname(params.filePath);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    fs.writeFileSync(params.filePath, JSON.stringify({ version: 1, delays }, null, 2));
    return true;
  } catch {
    return false;
  }
}
