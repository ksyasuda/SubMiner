import { fileURLToPath } from 'node:url';
import { resolve } from 'node:path';
import { spawn, spawnSync } from 'node:child_process';
import { collectLaneFiles } from './test-lanes.ts';

// Runs a test lane with per-file process isolation: one `bun test` process per
// test file so a hanging test or leaked global in one file cannot poison the
// rest of the lane. Use --single-process for the old all-in-one-process mode.
//
// Usage: bun scripts/run-test-lane.mjs <lane> [--jobs N] [--timeout-secs N] [--single-process]

const repoRoot = resolve(fileURLToPath(new URL('..', import.meta.url)));

// Cap per-file buffered output so a long or noisy test cannot grow the string
// without bound and exhaust memory.
const MAX_OUTPUT_BYTES = 1024 * 1024;

// Track spawned `bun test` children so we can kill them if the runner is
// interrupted, avoiding orphaned in-flight test processes.
const activeChildren = new Set();

function terminateChildren() {
  for (const child of activeChildren) {
    child.kill('SIGKILL');
  }
  activeChildren.clear();
}

for (const signal of ['SIGINT', 'SIGTERM']) {
  process.on(signal, () => {
    terminateChildren();
    process.exit(130);
  });
}

function parseArgs(argv) {
  const options = { lane: undefined, jobs: 1, timeoutSecs: 300, singleProcess: false };
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === '--jobs') {
      options.jobs = Math.max(1, Number(argv[(index += 1)]) || 1);
    } else if (arg === '--timeout-secs') {
      options.timeoutSecs = Math.max(1, Number(argv[(index += 1)]) || 300);
    } else if (arg === '--single-process') {
      options.singleProcess = true;
    } else if (!arg.startsWith('--') && options.lane === undefined) {
      options.lane = arg;
    } else {
      process.stderr.write(`Unknown argument: ${arg}\n`);
      process.exit(1);
    }
  }
  return options;
}

function runFile(file, timeoutSecs) {
  return new Promise((resolvePromise) => {
    const child = spawn('bun', ['test', `./${file}`], { cwd: repoRoot });
    activeChildren.add(child);
    let output = '';
    let truncated = false;
    let timedOut = false;
    const append = (chunk) => {
      if (truncated) return;
      output += chunk;
      if (output.length > MAX_OUTPUT_BYTES) {
        output = `${output.slice(0, MAX_OUTPUT_BYTES)}\n[output truncated at ${MAX_OUTPUT_BYTES} bytes]\n`;
        truncated = true;
      }
    };
    child.stdout.on('data', append);
    child.stderr.on('data', append);
    const timer = setTimeout(() => {
      timedOut = true;
      child.kill('SIGKILL');
    }, timeoutSecs * 1000);
    child.on('close', (code) => {
      clearTimeout(timer);
      activeChildren.delete(child);
      resolvePromise({ file, code: timedOut ? 124 : (code ?? 1), output, timedOut });
    });
    child.on('error', (error) => {
      clearTimeout(timer);
      activeChildren.delete(child);
      resolvePromise({ file, code: 1, output: String(error), timedOut: false });
    });
  });
}

async function runIsolated(files, options) {
  const failures = [];
  let nextIndex = 0;
  let completed = 0;

  async function worker() {
    while (nextIndex < files.length) {
      const file = files[nextIndex];
      nextIndex += 1;
      const result = await runFile(file, options.timeoutSecs);
      completed += 1;
      if (result.code !== 0) {
        failures.push(result);
        const reason = result.timedOut ? `timed out after ${options.timeoutSecs}s` : 'failed';
        process.stderr.write(`\n[${completed}/${files.length}] ${file} ${reason}\n`);
        process.stderr.write(result.output);
      }
    }
  }

  await Promise.all(Array.from({ length: Math.min(options.jobs, files.length) }, worker));

  if (failures.length > 0) {
    process.stderr.write(`\n${failures.length} of ${files.length} test files failed:\n`);
    for (const failure of failures) {
      process.stderr.write(`  ${failure.file}${failure.timedOut ? ' (timeout)' : ''}\n`);
    }
    return 1;
  }
  process.stdout.write(`All ${files.length} test files passed.\n`);
  return 0;
}

function runSingleProcess(files) {
  const result = spawnSync('bun', ['test', ...files.map((file) => `./${file}`)], {
    cwd: repoRoot,
    stdio: 'inherit',
  });
  if (result.error) {
    throw result.error;
  }
  return result.status ?? 1;
}

const options = parseArgs(process.argv.slice(2));

if (!options.lane) {
  process.stderr.write('Missing test lane name\n');
  process.exit(1);
}

let files;
try {
  files = collectLaneFiles(repoRoot, options.lane);
} catch (error) {
  process.stderr.write(`${error instanceof Error ? error.message : error}\n`);
  process.exit(1);
}

if (options.singleProcess) {
  process.exit(runSingleProcess(files));
}

process.exit(await runIsolated(files, options));
