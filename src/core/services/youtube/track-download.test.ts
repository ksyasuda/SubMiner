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
  const script = `#!/usr/bin/env bun
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
  if (process.platform === 'win32') {
    fs.writeFileSync(scriptPath + '.cmd', `@echo off\r\nbun "${scriptPath}" %*\r\n`, 'utf8');
  }
  return scriptPath;
}

function makeFakeYtDlpShellScript(dir: string): string {
  const scriptPath = path.join(dir, 'yt-dlp');
  const script = `#!/bin/sh
PATH=/usr/bin:/bin:/usr/local/bin
has_auto_subs=0
wants_auto_subs=0
wants_manual_subs=0
sub_lang=''
output_template=''

while [ "$#" -gt 0 ]; do
  case "$1" in
    --write-auto-subs)
      wants_auto_subs=1
      ;;
    --write-subs)
      wants_manual_subs=1
      ;;
    --sub-langs)
      sub_lang="$2"
      shift
      ;;
    -o)
      output_template="$2"
      shift
      ;;
  esac
  shift
done

if [ "$YTDLP_EXPECT_AUTO_SUBS" = "1" ] && [ "$wants_auto_subs" != "1" ]; then
  exit 2
fi
if [ "$YTDLP_EXPECT_MANUAL_SUBS" = "1" ] && [ "$wants_manual_subs" != "1" ]; then
  exit 3
fi
if [ -n "$YTDLP_EXPECT_SUB_LANG" ] && [ "$sub_lang" != "$YTDLP_EXPECT_SUB_LANG" ]; then
  exit 4
fi

prefix="$(printf '%s' "$output_template" | sed 's/\.%(ext)s$//')"
if [ -z "$prefix" ]; then
  exit 1
fi
mkdir -p "$(dirname \"$prefix\")"

if [ "$YTDLP_FAKE_MODE" = "multi" ]; then
  OLD_IFS="$IFS"
  IFS=","
  for lang in $sub_lang; do
    if [ -n "$lang" ]; then
      printf 'WEBVTT\\n' > "\${prefix}.\${lang}.vtt"
    fi
  done
  IFS="$OLD_IFS"
elif [ "$YTDLP_FAKE_MODE" = "rolling-auto" ]; then
  cat <<'EOF' > "\${prefix}.vtt"
WEBVTT

00:00:01.000 --> 00:00:02.000
今日は

00:00:02.000 --> 00:00:03.000
今日はいい天気ですね

00:00:03.000 --> 00:00:04.000
今日はいい天気ですね本当に
EOF
elif [ "$YTDLP_FAKE_MODE" = "multi-primary-only-fail" ]; then
  primary_lang="\${sub_lang%%,*}"
  if [ -n "$primary_lang" ]; then
    printf 'WEBVTT\\n' > "\${prefix}.\${primary_lang}.vtt"
  fi
  printf "ERROR: Unable to download video subtitles for 'en': HTTP Error 429: Too Many Requests\\n" 1>&2
  exit 1
elif [ "$YTDLP_FAKE_MODE" = "both" ]; then
  printf 'WEBVTT\\n' > "\${prefix}.vtt"
  printf 'webp' > "\${prefix}.orig.webp"
elif [ "$YTDLP_FAKE_MODE" = "webp-only" ]; then
  printf 'webp' > "\${prefix}.orig.webp"
else
  printf 'WEBVTT\\n' > "\${prefix}.vtt"
fi
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
    if (process.platform === 'win32') {
      makeFakeYtDlpScript(binDir);
    } else {
      makeFakeYtDlpShellScript(binDir);
    }

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

async function withFakeYtDlpCommand<T>(
  mode: 'both' | 'webp-only' | 'multi' | 'multi-primary-only-fail' | 'rolling-auto',
  fn: (dir: string, binDir: string) => Promise<T>,
): Promise<T> {
  return await withTempDir(async (root) => {
    const binDir = path.join(root, 'bin');
    fs.mkdirSync(binDir, { recursive: true });

    const originalPath = process.env.PATH;
    const originalCommand = process.env.SUBMINER_YTDLP_BIN;
    process.env.PATH = '';
    process.env.YTDLP_FAKE_MODE = mode;
    process.env.SUBMINER_YTDLP_BIN =
      process.platform === 'win32' ? path.join(binDir, 'yt-dlp.cmd') : path.join(binDir, 'yt-dlp');
    if (process.platform === 'win32') {
      makeFakeYtDlpScript(binDir);
    } else {
      makeFakeYtDlpShellScript(binDir);
    }
    try {
      return await fn(root, binDir);
    } finally {
      if (originalPath === undefined) {
        delete process.env.PATH;
      } else {
        process.env.PATH = originalPath;
      }
      delete process.env.YTDLP_FAKE_MODE;
      if (originalCommand === undefined) {
        delete process.env.SUBMINER_YTDLP_BIN;
      } else {
        process.env.SUBMINER_YTDLP_BIN = originalCommand;
      }
    }
  });
}

async function withFakeYtDlpExpectations<T>(
  expectations: Partial<
    Record<'YTDLP_EXPECT_AUTO_SUBS' | 'YTDLP_EXPECT_MANUAL_SUBS' | 'YTDLP_EXPECT_SUB_LANG', string>
  >,
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
      typeof input === 'string' ? input : input instanceof URL ? input.toString() : input.url;
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
    });

    assert.equal(path.extname(result.path), '.vtt');
    assert.match(path.basename(result.path), /^auto-ja-orig\./);
  });
});

test('downloadYoutubeSubtitleTrack honors SUBMINER_YTDLP_BIN when yt-dlp is not on PATH', async () => {
  if (process.platform === 'win32') {
    return;
  }

  await withFakeYtDlpCommand('both', async (root) => {
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
        });

        assert.equal(path.extname(result.path), '.vtt');
      },
    );
  });
});

test('downloadYoutubeSubtitleTrack normalizes rolling auto-caption vtt output from yt-dlp', async () => {
  if (process.platform === 'win32') {
    return;
  }

  await withFakeYtDlp('rolling-auto', async (root) => {
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
    });

    assert.equal(
      fs.readFileSync(result.path, 'utf8'),
      [
        'WEBVTT',
        '',
        '00:00:01.000 --> 00:00:02.000',
        '今日は',
        '',
        '00:00:02.000 --> 00:00:03.000',
        'いい天気ですね',
        '',
        '00:00:03.000 --> 00:00:04.000',
        '本当に',
        '',
      ].join('\n'),
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
        });

        assert.equal(path.basename(result.path), 'auto-ja-orig.ja-orig.vtt');
        assert.equal(fs.readFileSync(result.path, 'utf8'), 'WEBVTT\n');
      },
    );
  });
});

test('downloadYoutubeSubtitleTrack sanitizes metadata source language in filenames', async () => {
  await withTempDir(async (root) => {
    await withStubFetch(
      async () => new Response('WEBVTT\n', { status: 200 }),
      async () => {
        const result = await downloadYoutubeSubtitleTrack({
          targetUrl: 'https://www.youtube.com/watch?v=abc123',
          outputDir: path.join(root, 'out'),
          track: {
            id: 'auto:../../ja-orig',
            language: 'ja',
            sourceLanguage: '../ja-orig/../../evil',
            kind: 'auto',
            label: 'Japanese (auto)',
            downloadUrl: 'https://example.com/subs/ja.vtt',
            fileExtension: 'vtt',
          },
        });

        assert.equal(path.dirname(result.path), path.join(root, 'out'));
        assert.equal(path.basename(result.path), 'auto-ja-orig.ja-orig-evil.vtt');
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

test('downloadYoutubeSubtitleTracks keeps duplicate source-language direct downloads distinct', async () => {
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
              downloadUrl: 'https://example.com/subs/ja-auto.vtt',
              fileExtension: 'vtt',
            },
            {
              id: 'manual:ja-orig',
              language: 'ja',
              sourceLanguage: 'ja-orig',
              kind: 'manual',
              label: 'Japanese (manual)',
              downloadUrl: 'https://example.com/subs/ja-manual.vtt',
              fileExtension: 'vtt',
            },
          ],
        });

        assert.deepEqual(seen, [
          'https://example.com/subs/ja-auto.vtt',
          'https://example.com/subs/ja-manual.vtt',
        ]);
        assert.notEqual(result.get('auto:ja-orig'), result.get('manual:ja-orig'));
      },
    );
  });
});
