import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { downloadYoutubeSubtitleTrack, downloadYoutubeSubtitleTracks } from './track-download';

async function withTempDir<T>(fn: (dir: string) => Promise<T>): Promise<T> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-youtube-track-download-'));
  try {
    return await fn(dir);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

function makeFakeYtDlpScript(dir: string): string {
  const scriptPath = path.join(dir, 'yt-dlp');
const script = `#!/usr/bin/env node
const fs = require('node:fs');
const path = require('node:path');

const args = process.argv.slice(2);
let outputTemplate = '';
const wantsAutoSubs = args.includes('--write-auto-subs');
const wantsManualSubs = args.includes('--write-subs');
const subLangIndex = args.indexOf('--sub-langs');
const subLang = subLangIndex >= 0 ? args[subLangIndex + 1] || '' : '';
const subLangs = subLang ? subLang.split(',').filter(Boolean) : [];
for (let i = 0; i < args.length; i += 1) {
  if (args[i] === '-o' && typeof args[i + 1] === 'string') {
    outputTemplate = args[i + 1];
    i += 1;
  }
}

if (process.env.YTDLP_EXPECT_AUTO_SUBS === '1' && !wantsAutoSubs) {
  process.exit(2);
}
if (process.env.YTDLP_EXPECT_MANUAL_SUBS === '1' && !wantsManualSubs) {
  process.exit(3);
}
if (process.env.YTDLP_EXPECT_SUB_LANG && subLang !== process.env.YTDLP_EXPECT_SUB_LANG) {
  process.exit(4);
}

const prefix = outputTemplate.replace(/\.%\([^)]+\)s$/, '');
if (!prefix) {
  process.exit(1);
}
fs.mkdirSync(path.dirname(prefix), { recursive: true });

if (process.env.YTDLP_FAKE_MODE === 'multi') {
  for (const lang of subLangs) {
    fs.writeFileSync(\`\${prefix}.\${lang}.vtt\`, 'WEBVTT\\n');
  }
} else if (process.env.YTDLP_FAKE_MODE === 'rolling-auto') {
  fs.writeFileSync(
    \`\${prefix}.vtt\`,
    [
      'WEBVTT',
      '',
      '00:00:01.000 --> 00:00:02.000',
      '今日は',
      '',
      '00:00:02.000 --> 00:00:03.000',
      '今日はいい天気ですね',
      '',
      '00:00:03.000 --> 00:00:04.000',
      '今日はいい天気ですね本当に',
      '',
    ].join('\\n'),
  );
} else if (process.env.YTDLP_FAKE_MODE === 'multi-primary-only-fail') {
  const primaryLang = subLangs[0];
  if (primaryLang) {
    fs.writeFileSync(\`\${prefix}.\${primaryLang}.vtt\`, 'WEBVTT\\n');
  }
  process.stderr.write("ERROR: Unable to download video subtitles for 'en': HTTP Error 429: Too Many Requests\\n");
  process.exit(1);
} else if (process.env.YTDLP_FAKE_MODE === 'both') {
  fs.writeFileSync(\`\${prefix}.vtt\`, 'WEBVTT\\n');
  fs.writeFileSync(\`\${prefix}.orig.webp\`, 'webp');
} else if (process.env.YTDLP_FAKE_MODE === 'webp-only') {
  fs.writeFileSync(\`\${prefix}.orig.webp\`, 'webp');
} else {
  fs.writeFileSync(\`\${prefix}.vtt\`, 'WEBVTT\\n');
}
process.exit(0);
`;
  fs.writeFileSync(scriptPath, script, 'utf8');
  fs.chmodSync(scriptPath, 0o755);
  return scriptPath;
}

async function withFakeYtDlp<T>(
  mode: 'both' | 'webp-only' | 'multi' | 'multi-primary-only-fail' | 'rolling-auto',
  fn: (dir: string, binDir: string) => Promise<T>,
): Promise<T> {
  return await withTempDir(async (root) => {
    const binDir = path.join(root, 'bin');
    fs.mkdirSync(binDir, { recursive: true });
    makeFakeYtDlpScript(binDir);

    const originalPath = process.env.PATH ?? '';
    process.env.PATH = `${binDir}${path.delimiter}${originalPath}`;
    process.env.YTDLP_FAKE_MODE = mode;
    try {
      return await fn(root, binDir);
    } finally {
      process.env.PATH = originalPath;
      delete process.env.YTDLP_FAKE_MODE;
    }
  });
}

async function withFakeYtDlpExpectations<T>(
  expectations: Partial<Record<'YTDLP_EXPECT_AUTO_SUBS' | 'YTDLP_EXPECT_MANUAL_SUBS' | 'YTDLP_EXPECT_SUB_LANG', string>>,
  fn: () => Promise<T>,
): Promise<T> {
  const previous = {
    YTDLP_EXPECT_AUTO_SUBS: process.env.YTDLP_EXPECT_AUTO_SUBS,
    YTDLP_EXPECT_MANUAL_SUBS: process.env.YTDLP_EXPECT_MANUAL_SUBS,
    YTDLP_EXPECT_SUB_LANG: process.env.YTDLP_EXPECT_SUB_LANG,
  };
  Object.assign(process.env, expectations);
  try {
    return await fn();
  } finally {
    for (const [key, value] of Object.entries(previous)) {
      if (value === undefined) {
        delete process.env[key];
      } else {
        process.env[key] = value;
      }
    }
  }
}

async function withStubFetch<T>(
  handler: (url: string) => Promise<Response> | Response,
  fn: () => Promise<T>,
): Promise<T> {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async (input: string | URL | Request) => {
    const url =
      typeof input === 'string'
        ? input
        : input instanceof URL
          ? input.toString()
          : input.url;
    return await handler(url);
  }) as typeof fetch;
  try {
    return await fn();
  } finally {
    globalThis.fetch = originalFetch;
  }
}

test('downloadYoutubeSubtitleTrack prefers subtitle files over later webp artifacts', async () => {
  if (process.platform === 'win32') {
    return;
  }

  await withFakeYtDlp('both', async (root) => {
    const result = await downloadYoutubeSubtitleTrack({
      targetUrl: 'https://www.youtube.com/watch?v=abc123',
      outputDir: path.join(root, 'out'),
      track: {
        id: 'auto:ja-orig',
        language: 'ja',
        sourceLanguage: 'ja-orig',
        kind: 'auto',
        label: 'Japanese (auto)',
      },
      mode: 'download',
    });

    assert.equal(path.extname(result.path), '.vtt');
    assert.match(path.basename(result.path), /^auto-ja-orig\./);
  });
});

test('downloadYoutubeSubtitleTrack ignores stale subtitle files from prior runs', async () => {
  if (process.platform === 'win32') {
    return;
  }

  await withFakeYtDlp('webp-only', async (root) => {
    const outputDir = path.join(root, 'out');
    fs.mkdirSync(outputDir, { recursive: true });
    fs.writeFileSync(path.join(outputDir, 'auto-ja.vtt'), 'stale subtitle');

    await assert.rejects(
      async () =>
        await downloadYoutubeSubtitleTrack({
          targetUrl: 'https://www.youtube.com/watch?v=abc123',
          outputDir,
          track: {
            id: 'auto:ja',
            language: 'ja',
            sourceLanguage: 'ja',
            kind: 'auto',
            label: 'Japanese (auto)',
          },
          mode: 'download',
        }),
      /No subtitle file was downloaded/,
    );
  });
});

test('downloadYoutubeSubtitleTrack uses auto subtitle flags and raw source language for auto tracks', async () => {
  if (process.platform === 'win32') {
    return;
  }

  await withFakeYtDlp('both', async (root) => {
    await withFakeYtDlpExpectations(
      {
        YTDLP_EXPECT_AUTO_SUBS: '1',
        YTDLP_EXPECT_SUB_LANG: 'ja-orig',
      },
      async () => {
        const result = await downloadYoutubeSubtitleTrack({
          targetUrl: 'https://www.youtube.com/watch?v=abc123',
          outputDir: path.join(root, 'out'),
          track: {
            id: 'auto:ja-orig',
            language: 'ja',
            sourceLanguage: 'ja-orig',
            kind: 'auto',
            label: 'Japanese (auto)',
          },
          mode: 'download',
        });

        assert.equal(path.extname(result.path), '.vtt');
      },
    );
  });
});

test('downloadYoutubeSubtitleTrack keeps manual subtitle flag for manual tracks', async () => {
  if (process.platform === 'win32') {
    return;
  }

  await withFakeYtDlp('both', async (root) => {
    await withFakeYtDlpExpectations(
      {
        YTDLP_EXPECT_MANUAL_SUBS: '1',
        YTDLP_EXPECT_SUB_LANG: 'ja',
      },
      async () => {
        const result = await downloadYoutubeSubtitleTrack({
          targetUrl: 'https://www.youtube.com/watch?v=abc123',
          outputDir: path.join(root, 'out'),
          track: {
            id: 'manual:ja',
            language: 'ja',
            sourceLanguage: 'ja',
            kind: 'manual',
            label: 'Japanese (manual)',
          },
          mode: 'download',
        });

        assert.equal(path.extname(result.path), '.vtt');
      },
    );
  });
});

test('downloadYoutubeSubtitleTrack prefers direct download URL when available', async () => {
  await withTempDir(async (root) => {
    await withStubFetch(
      async (url) => {
        assert.equal(url, 'https://example.com/subs/ja.vtt');
        return new Response('WEBVTT\n', { status: 200 });
      },
      async () => {
        const result = await downloadYoutubeSubtitleTrack({
          targetUrl: 'https://www.youtube.com/watch?v=abc123',
          outputDir: path.join(root, 'out'),
          track: {
            id: 'auto:ja-orig',
            language: 'ja',
            sourceLanguage: 'ja-orig',
            kind: 'auto',
            label: 'Japanese (auto)',
            downloadUrl: 'https://example.com/subs/ja.vtt',
            fileExtension: 'vtt',
          },
          mode: 'download',
        });

        assert.equal(path.basename(result.path), 'auto-ja-orig.ja-orig.vtt');
        assert.equal(fs.readFileSync(result.path, 'utf8'), 'WEBVTT\n');
      },
    );
  });
});

test('downloadYoutubeSubtitleTrack converts srv3 auto subtitles into regular vtt', async () => {
  await withTempDir(async (root) => {
    await withStubFetch(
      async (url) => {
        assert.equal(url, 'https://example.com/subs/ja.srv3');
        return new Response(
          [
            '<timedtext><body>',
            '<p t="1000" d="2500">今日は</p>',
            '<p t="2000" d="2500">今日はいい天気ですね</p>',
            '<p t="3500" d="2500">今日はいい天気ですね本当に</p>',
            '</body></timedtext>',
          ].join(''),
          { status: 200 },
        );
      },
      async () => {
        const result = await downloadYoutubeSubtitleTrack({
          targetUrl: 'https://www.youtube.com/watch?v=abc123',
          outputDir: path.join(root, 'out'),
          track: {
            id: 'auto:ja-orig',
            language: 'ja',
            sourceLanguage: 'ja-orig',
            kind: 'auto',
            label: 'Japanese (auto)',
            downloadUrl: 'https://example.com/subs/ja.srv3',
            fileExtension: 'srv3',
          },
          mode: 'download',
        });

        assert.equal(path.basename(result.path), 'auto-ja-orig.ja-orig.vtt');
        assert.equal(
          fs.readFileSync(result.path, 'utf8'),
          [
            'WEBVTT',
            '',
            '00:00:01.000 --> 00:00:01.999',
            '今日は',
            '',
            '00:00:02.000 --> 00:00:03.499',
            'いい天気ですね',
            '',
            '00:00:03.500 --> 00:00:06.000',
            '本当に',
            '',
          ].join('\n'),
        );
      },
    );
  });
});

test('downloadYoutubeSubtitleTracks downloads primary and secondary in one invocation', async () => {
  if (process.platform === 'win32') {
    return;
  }

  await withFakeYtDlp('multi', async (root) => {
    const outputDir = path.join(root, 'out');
    const result = await downloadYoutubeSubtitleTracks({
      targetUrl: 'https://www.youtube.com/watch?v=abc123',
      outputDir,
      tracks: [
        {
          id: 'auto:ja-orig',
          language: 'ja',
          sourceLanguage: 'ja-orig',
          kind: 'auto',
          label: 'Japanese (auto)',
        },
        {
          id: 'auto:en',
          language: 'en',
          sourceLanguage: 'en',
          kind: 'auto',
          label: 'English (auto)',
        },
      ],
      mode: 'download',
    });

    assert.match(path.basename(result.get('auto:ja-orig') ?? ''), /\.ja-orig\.vtt$/);
    assert.match(path.basename(result.get('auto:en') ?? ''), /\.en\.vtt$/);
  });
});

test('downloadYoutubeSubtitleTracks preserves successfully downloaded primary file on partial failure', async () => {
  if (process.platform === 'win32') {
    return;
  }

  await withFakeYtDlp('multi-primary-only-fail', async (root) => {
    const outputDir = path.join(root, 'out');
    const result = await downloadYoutubeSubtitleTracks({
      targetUrl: 'https://www.youtube.com/watch?v=abc123',
      outputDir,
      tracks: [
        {
          id: 'auto:ja-orig',
          language: 'ja',
          sourceLanguage: 'ja-orig',
          kind: 'auto',
          label: 'Japanese (auto)',
        },
        {
          id: 'auto:en',
          language: 'en',
          sourceLanguage: 'en',
          kind: 'auto',
          label: 'English (auto)',
        },
      ],
      mode: 'download',
    });

    assert.match(path.basename(result.get('auto:ja-orig') ?? ''), /\.ja-orig\.vtt$/);
    assert.equal(result.has('auto:en'), false);
  });
});

test('downloadYoutubeSubtitleTracks prefers direct download URLs when available', async () => {
  await withTempDir(async (root) => {
    const seen: string[] = [];
    await withStubFetch(
      async (url) => {
        seen.push(url);
        return new Response(`WEBVTT\n${url}\n`, { status: 200 });
      },
      async () => {
        const result = await downloadYoutubeSubtitleTracks({
          targetUrl: 'https://www.youtube.com/watch?v=abc123',
          outputDir: path.join(root, 'out'),
          tracks: [
            {
              id: 'auto:ja-orig',
              language: 'ja',
              sourceLanguage: 'ja-orig',
              kind: 'auto',
              label: 'Japanese (auto)',
              downloadUrl: 'https://example.com/subs/ja.vtt',
              fileExtension: 'vtt',
            },
            {
              id: 'auto:en',
              language: 'en',
              sourceLanguage: 'en',
              kind: 'auto',
              label: 'English (auto)',
              downloadUrl: 'https://example.com/subs/en.vtt',
              fileExtension: 'vtt',
            },
          ],
          mode: 'download',
        });

        assert.deepEqual(seen, [
          'https://example.com/subs/ja.vtt',
          'https://example.com/subs/en.vtt',
        ]);
        assert.match(path.basename(result.get('auto:ja-orig') ?? ''), /\.ja-orig\.vtt$/);
        assert.match(path.basename(result.get('auto:en') ?? ''), /\.en\.vtt$/);
      },
    );
  });
});
