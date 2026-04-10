import fs from 'node:fs';
import path from 'node:path';
import type { PluginSessionBindingsArtifact } from '../../types';

export function getSessionBindingsArtifactPath(configDir: string): string {
  return path.join(configDir, 'session-bindings.json');
}

export function writeSessionBindingsArtifact(
  configDir: string,
  artifact: PluginSessionBindingsArtifact,
): string {
  const artifactPath = getSessionBindingsArtifactPath(configDir);
  fs.mkdirSync(configDir, { recursive: true });
  fs.writeFileSync(artifactPath, `${JSON.stringify(artifact, null, 2)}\n`, 'utf8');
  return artifactPath;
}
