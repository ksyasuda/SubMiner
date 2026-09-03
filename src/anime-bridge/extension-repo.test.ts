import test from 'node:test';
import assert from 'node:assert/strict';
import {
  extensionFileName,
  fetchRepoCatalogue,
  fetchRepoIndex,
  isValidRepoUrl,
  parseRepoIndex,
  repoBaseUrl,
} from './extension-repo';

const INDEX = 'https://repo.example/anime/index.min.json';

function animeEntry(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    name: 'Aniyomi: Example Source',
    pkg: 'eu.kanade.tachiyomi.animeextension.all.example',
    apk: 'example-v1.2.3.apk',
    lang: 'all',
    code: 12,
    version: '1.2.3',
    nsfw: 0,
    sources: [{ name: 'Example', lang: 'en' }],
    ...overrides,
  };
}

test('any https url naming a json index is accepted', () => {
  assert.equal(isValidRepoUrl(INDEX), true);
  assert.equal(isValidRepoUrl('  ' + INDEX + '  '), true);
  // Repos are free to name the index; index.min.json is only a convention.
  assert.equal(isValidRepoUrl('https://repo.example/anime/index.json'), true);
  assert.equal(
    isValidRepoUrl('https://manatan-community.github.io/extensions/video.min.json'),
    true,
  );
  // Plain http would let a network attacker swap the APK list.
  assert.equal(isValidRepoUrl('http://repo.example/anime/index.min.json'), false);
  assert.equal(isValidRepoUrl('https://repo.example/anime/'), false);
  assert.equal(isValidRepoUrl('https://repo.example/index.min.json.txt'), false);
  assert.equal(isValidRepoUrl('https://repo.example'), false);
  assert.equal(isValidRepoUrl(''), false);
});

test('repoBaseUrl strips the index file name', () => {
  assert.equal(repoBaseUrl(INDEX), 'https://repo.example/anime');
  assert.equal(
    repoBaseUrl('https://manatan-community.github.io/extensions/video.min.json'),
    'https://manatan-community.github.io/extensions',
  );
});

test('parseRepoIndex builds apk and icon urls from the repo root', () => {
  const [extension] = parseRepoIndex(INDEX, [animeEntry()]);

  assert.equal(extension?.pkg, 'eu.kanade.tachiyomi.animeextension.all.example');
  assert.equal(extension?.apkUrl, 'https://repo.example/anime/apk/example-v1.2.3.apk');
  assert.equal(
    extension?.iconUrl,
    'https://repo.example/anime/icon/eu.kanade.tachiyomi.animeextension.all.example.png',
  );
  assert.equal(extension?.repoUrl, INDEX);
  assert.equal(extension?.versionCode, 12);
  assert.deepEqual(extension?.sourceNames, ['Example']);
});

test('the Aniyomi name prefix is stripped', () => {
  const [extension] = parseRepoIndex(INDEX, [animeEntry()]);
  assert.equal(extension?.name, 'Example Source');
});

test('manga packages are excluded', () => {
  const entries = [animeEntry(), animeEntry({ pkg: 'eu.kanade.tachiyomi.extension.en.somemanga' })];
  const parsed = parseRepoIndex(INDEX, entries);
  assert.equal(parsed.length, 1);
  assert.match(parsed[0]!.pkg, /animeextension/);
});

test('malformed entries are skipped rather than failing the repo', () => {
  const parsed = parseRepoIndex(INDEX, [
    null,
    'nonsense',
    animeEntry({ apk: undefined }),
    animeEntry({ pkg: undefined }),
    animeEntry(),
  ]);
  assert.equal(parsed.length, 1);
});

test('a package name that is not a plain identifier is rejected', () => {
  // The package name becomes the on-disk file name, and a repo index is
  // unauthenticated: path separators here would write outside the extensions
  // directory even though the prefix check passes.
  const parsed = parseRepoIndex(INDEX, [
    animeEntry({ pkg: 'eu.kanade.tachiyomi.animeextension/../../../../etc/cron.d/x' }),
    animeEntry({ pkg: 'eu.kanade.tachiyomi.animeextension\\..\\evil' }),
    animeEntry({ pkg: 'eu.kanade.tachiyomi.animeextension.all.ok' }),
  ]);
  assert.deepEqual(
    parsed.map((extension) => extension.pkg),
    ['eu.kanade.tachiyomi.animeextension.all.ok'],
  );
});

test('an apk file name with path characters is rejected', () => {
  const parsed = parseRepoIndex(INDEX, [animeEntry({ apk: '../../../etc/passwd' })]);
  assert.deepEqual(parsed, []);
});

test('parseRepoIndex tolerates a non-array payload', () => {
  assert.deepEqual(parseRepoIndex(INDEX, { message: 'Not Found' }), []);
  assert.deepEqual(parseRepoIndex(INDEX, null), []);
});

test('missing optional fields fall back to safe defaults', () => {
  const [extension] = parseRepoIndex(INDEX, [
    { pkg: 'eu.kanade.tachiyomi.animeextension.all.bare', apk: 'bare.apk' },
  ]);
  assert.equal(extension?.name, 'eu.kanade.tachiyomi.animeextension.all.bare');
  assert.equal(extension?.lang, 'all');
  assert.equal(extension?.versionCode, 0);
  assert.equal(extension?.nsfw, false);
  assert.deepEqual(extension?.sourceNames, []);
});

test('nsfw is read from the numeric flag', () => {
  assert.equal(parseRepoIndex(INDEX, [animeEntry({ nsfw: 1 })])[0]?.nsfw, true);
  assert.equal(parseRepoIndex(INDEX, [animeEntry({ nsfw: 0 })])[0]?.nsfw, false);
});

test('fetchRepoIndex rejects an invalid url before making a request', async () => {
  let called = false;
  const fetchImpl = (async () => {
    called = true;
    return new Response('[]');
  }) as typeof fetch;

  await assert.rejects(() => fetchRepoIndex('http://insecure/index.min.json', { fetchImpl }));
  assert.equal(called, false);
});

test('fetchRepoIndex surfaces a non-ok response', async () => {
  const fetchImpl = (async () => new Response('', { status: 404 })) as typeof fetch;
  await assert.rejects(() => fetchRepoIndex(INDEX, { fetchImpl }), /404/);
});

test('fetchRepoIndex applies a deadline when the caller supplies no signal', async () => {
  let receivedSignal: AbortSignal | undefined;
  const fetchImpl = (async (_input: RequestInfo | URL, init?: RequestInit) => {
    receivedSignal = init?.signal instanceof AbortSignal ? init.signal : undefined;
    return new Response('[]');
  }) as typeof fetch;

  await fetchRepoIndex(INDEX, { fetchImpl, timeoutMs: 50 });

  assert.ok(receivedSignal, 'repository request should receive a deadline signal');
});

test('fetchRepoCatalogue merges repos and keeps the highest version code', async () => {
  const second = 'https://other.example/anime/index.min.json';
  const fetchImpl = (async (input: RequestInfo | URL) => {
    const url = String(input);
    if (url === INDEX) {
      return new Response(JSON.stringify([animeEntry({ code: 12, version: '1.2.3' })]));
    }
    return new Response(JSON.stringify([animeEntry({ code: 20, version: '2.0.0' })]));
  }) as typeof fetch;

  const catalogue = await fetchRepoCatalogue([INDEX, second], { fetchImpl });

  assert.equal(catalogue.extensions.length, 1);
  assert.equal(catalogue.extensions[0]?.versionCode, 20);
  assert.equal(catalogue.extensions[0]?.repoUrl, second);
  assert.deepEqual(catalogue.failures, []);
});

test('one failing repo does not hide the others', async () => {
  const broken = 'https://broken.example/anime/index.min.json';
  const fetchImpl = (async (input: RequestInfo | URL) => {
    if (String(input) === broken) throw new Error('ENOTFOUND');
    return new Response(JSON.stringify([animeEntry()]));
  }) as typeof fetch;

  const catalogue = await fetchRepoCatalogue([broken, INDEX], { fetchImpl });

  assert.equal(catalogue.extensions.length, 1);
  assert.equal(catalogue.failures.length, 1);
  assert.equal(catalogue.failures[0]?.repoUrl, broken);
  assert.match(catalogue.failures[0]?.error ?? '', /ENOTFOUND/);
});

test('an empty repo list yields an empty catalogue without any request', async () => {
  let called = false;
  const fetchImpl = (async () => {
    called = true;
    return new Response('[]');
  }) as typeof fetch;

  const catalogue = await fetchRepoCatalogue([], { fetchImpl });
  assert.deepEqual(catalogue, { extensions: [], failures: [] });
  assert.equal(called, false);
});

test('extensions are stored under their package name so updates replace in place', () => {
  assert.equal(
    extensionFileName('eu.kanade.tachiyomi.animeextension.all.example'),
    'eu.kanade.tachiyomi.animeextension.all.example.apk',
  );
});
