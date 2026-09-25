import assert from 'node:assert/strict';
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { test } from 'bun:test';

test.each([false, true])(
  'AUR downloads handle empty release metadata, unavailable=%s',
  async (unavailable) => {
    const workflow = await readFile(
      new URL('../.github/workflows/release.yml', import.meta.url),
      'utf8',
    );
    const step = workflow
      .split('      - name: Download release assets for AUR\n')[1]
      ?.split('\n      - name:')[0];
    const script = step?.split('        run: |\n')[1]?.replace(/^          /gm, '');
    assert.ok(script, 'AUR download step must have a shell script');

    const workspace = await mkdtemp(path.join(os.tmpdir(), 'subminer-aur-download-'));
    const requests: string[] = [];
    const files = new Map([
      ['SubMiner-0.20.0.AppImage', 'appimage bytes'],
      ['subminer', 'launcher bytes'],
      ['subminer-assets.tar.gz', 'optional assets bytes'],
    ]);
    const server = Bun.serve({
      hostname: '127.0.0.1',
      port: 0,
      fetch(request) {
        const pathname = new URL(request.url).pathname;
        requests.push(pathname);
        if (unavailable || requests.length === 1) return new Response('try again', { status: 503 });
        const name = pathname.split('/').at(-1);
        const body = name ? files.get(name) : undefined;
        return new Response(body ?? 'not found', { status: body ? 200 : 404 });
      },
    });

    try {
      const bin = path.join(workspace, 'bin');
      await mkdir(bin);
      await writeFile(
        path.join(bin, 'gh'),
        '#!/bin/sh\necho "no assets to download" >&2\nexit 1\n',
        { mode: 0o755 },
      );
      const output = path.join(workspace, 'output');
      const proc = Bun.spawn(['bash', '-c', script], {
        cwd: workspace,
        env: {
          ...process.env,
          PATH: `${bin}${path.delimiter}${process.env.PATH}`,
          RELEASE_VERSION: 'v0.20.0',
          GITHUB_SERVER_URL: server.url.origin,
          GITHUB_REPOSITORY: 'ksyasuda/SubMiner',
          GITHUB_OUTPUT: output,
        },
        stdout: 'pipe',
        stderr: 'pipe',
      });
      const [status, stderr, stdout] = await Promise.all([
        proc.exited,
        new Response(proc.stderr).text(),
        new Response(proc.stdout).text(),
      ]);
      assert.equal(status, 0, stderr);
      if (unavailable) {
        assert.equal(requests.length, 4, 'failed downloads stop after three retries');
        assert.match(await readFile(output, 'utf8'), /^skip=true$/m);
        assert.match(stdout, /::warning::Unable to download/);
        await assert.rejects(
          readFile(path.join(workspace, '.tmp/aur-release-assets/SubMiner-0.20.0.AppImage')),
          { code: 'ENOENT' },
        );
        return;
      }
      for (const [name, body] of files) {
        assert.equal(
          await readFile(path.join(workspace, '.tmp/aur-release-assets', name), 'utf8'),
          body,
        );
        assert.ok(requests.includes(`/ksyasuda/SubMiner/releases/download/v0.20.0/${name}`));
      }
      assert.equal(requests.length, 4, 'the first failed download must be retried');
      assert.match(await readFile(output, 'utf8'), /^skip=false$/m);
    } finally {
      server.stop(true);
      await rm(workspace, { recursive: true, force: true });
    }
  },
  15_000,
);
