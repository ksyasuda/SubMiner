import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const env = { ...process.env };

for (const name of [
  'CSC_LINK',
  'CSC_KEY_PASSWORD',
  'WIN_CSC_LINK',
  'WIN_CSC_KEY_PASSWORD',
  'CSC_NAME',
  'WIN_CSC_NAME',
]) {
  delete env[name];
}

env.CSC_IDENTITY_AUTO_DISCOVERY = 'false';

const electronBuilderCli = fileURLToPath(
  new URL('../node_modules/electron-builder/out/cli/cli.js', import.meta.url),
);

const result = spawnSync(
  process.execPath,
  [electronBuilderCli, '--win', 'nsis', 'zip', '--publish', 'never'],
  {
    stdio: 'inherit',
    env,
  },
);

if (result.error) {
  throw result.error;
}

process.exit(result.status ?? 1);
