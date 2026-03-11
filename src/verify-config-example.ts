import fs from 'node:fs';
import { DEFAULT_CONFIG, generateConfigTemplate } from './config';
import { resolveConfigExampleOutputPaths } from './generate-config-example';

export type ConfigExampleVerificationResult = {
  docsSiteDetected: boolean;
  missingPaths: string[];
  outputPaths: string[];
  stalePaths: string[];
  template: string;
};

export function verifyConfigExampleArtifacts(options?: {
  cwd?: string;
  docsSiteDirName?: string;
  template?: string;
  deps?: {
    existsSync?: (candidate: string) => boolean;
    readFileSync?: (candidate: string, encoding: BufferEncoding) => string;
  };
}): ConfigExampleVerificationResult {
  const existsSync = options?.deps?.existsSync ?? fs.existsSync;
  const readFileSync = options?.deps?.readFileSync ?? fs.readFileSync;
  const template = options?.template ?? generateConfigTemplate(DEFAULT_CONFIG);
  const outputPaths = resolveConfigExampleOutputPaths({
    cwd: options?.cwd,
    docsSiteDirName: options?.docsSiteDirName,
    existsSync,
  });
  const missingPaths: string[] = [];
  const stalePaths: string[] = [];

  for (const outputPath of outputPaths) {
    if (!existsSync(outputPath)) {
      missingPaths.push(outputPath);
      continue;
    }

    if (readFileSync(outputPath, 'utf-8') !== template) {
      stalePaths.push(outputPath);
    }
  }

  return {
    docsSiteDetected: outputPaths.length > 1,
    missingPaths,
    outputPaths,
    stalePaths,
    template,
  };
}

function main(): void {
  const result = verifyConfigExampleArtifacts();

  if (result.missingPaths.length === 0 && result.stalePaths.length === 0) {
    console.log('[OK] config example artifacts verified');
    for (const outputPath of result.outputPaths) {
      console.log(`     ${outputPath}`);
    }
    if (!result.docsSiteDetected) {
      console.log('     docs-site not present; skipped docs artifact check');
    }
    return;
  }

  console.error('[FAIL] config example artifacts are out of sync');
  for (const missingPath of result.missingPaths) {
    console.error(`       missing: ${missingPath}`);
  }
  for (const stalePath of result.stalePaths) {
    console.error(`       stale: ${stalePath}`);
  }
  console.error('       run: bun run generate:config-example');
  process.exitCode = 1;
}

if (require.main === module) {
  main();
}
