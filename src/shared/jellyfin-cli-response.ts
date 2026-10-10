import fs from 'node:fs';
import path from 'node:path';

// Structured replies the app writes to `--jellyfin-response-path`. The launcher passes a temp
// path and waits for the file instead of scraping logs, because the command usually runs in an
// already-running app instance whose output never reaches the launcher's helper process.

export type JellyfinCliLibrary = {
  id: string;
  name: string;
  collectionType: string;
};

export type JellyfinCliItem = {
  id: string;
  title: string;
  type: string;
};

export type JellyfinPreviewAuthPayload = {
  serverUrl: string;
  accessToken: string;
  userId: string;
};

export type JellyfinCliResponse =
  | { libraries: JellyfinCliLibrary[] }
  | { items: JellyfinCliItem[] }
  | { error: string }
  | JellyfinPreviewAuthPayload;

// Writes via rename so a polling reader never parses a half-written file.
export function writeJellyfinCliResponse(
  responsePath: string,
  response: JellyfinCliResponse,
): void {
  fs.mkdirSync(path.dirname(responsePath), { recursive: true });
  const tempPath = `${responsePath}.${process.pid}.tmp`;
  fs.writeFileSync(tempPath, JSON.stringify(response), 'utf-8');
  fs.renameSync(tempPath, responsePath);
}
