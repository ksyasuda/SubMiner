import { execFile } from 'node:child_process';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import type {
  SubtitleGenerationAcceleration,
  SubtitleGenerationToolStatus,
} from '../../shared/subtitle-generation';

function probe(command: string, args: string[], env: NodeJS.ProcessEnv) {
  return new Promise<{ code: number; stdout: string; stderr: string } | null>((resolve) => {
    execFile(
      command,
      args,
      { env, timeout: 3000, killSignal: 'SIGKILL', maxBuffer: 64 * 1024, windowsHide: true },
      (error, stdout, stderr) => {
        const code = error?.code ?? 0;
        if (typeof code !== 'number' || error?.killed || error?.signal) {
          resolve(null);
          return;
        }
        resolve({ code, stdout, stderr });
      },
    );
  });
}

/** Require both a working NVIDIA driver and CUDA device discovery in the selected Whisper binary. */
export async function detectSubtitleGenerationAcceleration(
  whisper: SubtitleGenerationToolStatus,
  env: NodeJS.ProcessEnv = process.env,
): Promise<SubtitleGenerationAcceleration> {
  const unavailable: SubtitleGenerationAcceleration = { kind: 'unavailable' };
  if (whisper.kind === 'missing') return unavailable;
  let directory: string | undefined;
  try {
    const nvidia = await probe('nvidia-smi', ['--query-gpu=name', '--format=csv,noheader'], env);
    const gpuName = nvidia?.stdout.trim().split(/\r?\n/)[0]?.trim();
    if (nvidia?.code !== 0 || !gpuName) return unavailable;

    directory = await mkdtemp(path.join(tmpdir(), 'subminer-cuda-check-'));
    const model = path.join(directory, 'probe.bin');
    await writeFile(model, Buffer.alloc(4));
    // whisper.cpp discovers backends before checking model magic. This deliberately invalid
    // local file stops before allocating a model or decoding audio, even on a fresh install.
    const result = await probe(whisper.path, ['-m', model, '-f', model], env);
    if (
      result &&
      result.code !== 0 &&
      /ggml_cuda_init:\s+found\s+[1-9]\d*\s+CUDA devices?\b/i.test(result.stderr) &&
      /invalid model data \(bad magic\)/i.test(result.stderr)
    ) {
      return { kind: 'nvidia-cuda', gpuName };
    }
    return unavailable;
  } catch {
    // Detection is advisory; unsupported builds and driver failures must not block generation.
    return unavailable;
  } finally {
    if (directory) await rm(directory, { recursive: true, force: true }).catch(() => {});
  }
}
