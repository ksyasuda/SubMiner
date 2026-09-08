import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, open, rm } from 'node:fs/promises';
import path from 'node:path';
import { publishSubtitleGenerationFile } from './subtitle-generation-files';
import type { SubtitleGenerationProgress } from '../../shared/subtitle-generation';

export async function downloadSubtitleGenerationArtifact(input: {
  url: string;
  size: number;
  sha256: string;
  destination: string;
  label: string;
  onProgress?: (progress: SubtitleGenerationProgress) => void;
  signal?: AbortSignal;
}): Promise<string> {
  input.signal?.throwIfAborted();
  await mkdir(path.dirname(input.destination), { recursive: true });
  const temporaryDirectory = await mkdtemp(
    path.join(path.dirname(input.destination), '.download-'),
  );
  const temporaryPath = path.join(temporaryDirectory, 'model.bin');
  input.onProgress?.({ stage: 'download', percent: 0, message: `Downloading ${input.label}...` });
  try {
    const response = await fetch(input.url, { signal: input.signal });
    if (!response.ok || !response.body) {
      throw new Error(`Model download failed: HTTP ${response.status}`);
    }
    const file = await open(temporaryPath, 'wx');
    const reader = response.body.getReader();
    const digest = createHash('sha256');
    let received = 0;
    let previousPercent = -1;
    try {
      while (true) {
        input.signal?.throwIfAborted();
        const { done, value } = await reader.read();
        if (done) break;
        received += value.byteLength;
        if (received > input.size) throw new Error('Downloaded model exceeds expected size.');
        digest.update(value);
        await file.writeFile(value);
        const percent = Math.min(99, Math.floor((received / input.size) * 100));
        if (percent !== previousPercent) {
          previousPercent = percent;
          input.onProgress?.({
            stage: 'download',
            percent,
            message: `Downloading ${input.label}...`,
          });
        }
      }
      if (received !== input.size || digest.digest('hex') !== input.sha256) {
        throw new Error('Downloaded model failed integrity verification. Try downloading again.');
      }
      await file.sync();
    } finally {
      await reader.cancel().catch(() => undefined);
      await file.close();
    }
    input.signal?.throwIfAborted();
    await publishSubtitleGenerationFile(temporaryPath, input.destination);
    input.onProgress?.({ stage: 'download', percent: 100, message: `${input.label} is ready.` });
    return input.destination;
  } finally {
    await rm(temporaryDirectory, { recursive: true, force: true });
  }
}
