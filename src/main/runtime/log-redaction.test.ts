import assert from 'node:assert/strict';
import { performance } from 'node:perf_hooks';
import test from 'node:test';
import { LOG_EXPORT_REDACTION_RULES, redactLogExportText } from './log-redaction';

test('log export redaction rules expose auditable fixtures', () => {
  assert.deepEqual(
    LOG_EXPORT_REDACTION_RULES.map((rule) => rule.name),
    [
      'signed-youtube-media-url-query',
      'url-sensitive-components',
      'sensitive-headers',
      'yt-dlp-cookie-args',
      'json-secret-values',
      'generic-sensitive-key-values',
      'home-path-usernames',
      'email-addresses',
      'ip-addresses',
    ],
  );

  const names = new Set<string>();
  for (const rule of LOG_EXPORT_REDACTION_RULES) {
    assert.equal(names.has(rule.name), false, `duplicate rule name: ${rule.name}`);
    names.add(rule.name);
    assert.ok(rule.pattern.length > 0, `${rule.name} pattern`);
    assert.ok(rule.replacement.length > 0, `${rule.name} replacement`);
    assert.ok(rule.risk.length > 0, `${rule.name} risk`);
    assert.ok(rule.examples.length > 0, `${rule.name} examples`);

    for (const example of rule.examples) {
      assert.equal(rule.redact(example.input), example.expected, `${rule.name} direct`);
      assert.equal(redactLogExportText(example.input), example.expected, rule.name);
    }
  }
});

test('redactLogExportText redacts parsed URL query values without swallowing punctuation', () => {
  const masked = redactLogExportText(
    [
      'failed URL (https://example.test/watch?access_token=tok123).',
      'callback subminer://anilist-setup?access_token=ani-token&state=ok',
      'encoded https://example.test/watch?client%5Fsecret=secret-value&plain=ok!',
    ].join('\n'),
  );

  assert.match(masked, /\(https:\/\/example\.test\/watch\?access_token=<redacted>\)\./);
  assert.match(masked, /subminer:\/\/anilist-setup\?access_token=<redacted>&state=ok/);
  assert.match(masked, /client%5Fsecret=<redacted>&plain=ok!/);
  assert.doesNotMatch(masked, /tok123/);
  assert.doesNotMatch(masked, /ani-token/);
  assert.doesNotMatch(masked, /secret-value/);
});

test('redactLogExportText redacts URL credentials containing at signs', () => {
  const masked = redactLogExportText('GET https://alice:p@ss@example.test/watch?state=ok');

  assert.equal(masked, 'GET https://<credentials>@example.test/watch?state=ok');
  assert.doesNotMatch(masked, /alice/);
  assert.doesNotMatch(masked, /p@ss/);
});

test('redactLogExportText redacts signed YouTube media URL query strings', () => {
  const masked = redactLogExportText(
    [
      'ffmpeg failed for (https://rr1---sn-a5mekn6r.googlevideo.com/videoplayback?expire=1777777777&ei=EI_VALUE&id=o-SECRETID&itag=251&ip=203.0.113.42&source=youtube&requiressl=yes&n=NSECRET&sparams=expire,ei,ip,id,itag,source,requiressl&sig=SIGSECRET&lsig=LSIGSECRET).',
      'manifest https://manifest.googlevideo.com/videoplayback?expire=1777777777&signature=SIG2#frag',
    ].join('\n'),
  );

  assert.match(
    masked,
    /\(https:\/\/rr1---sn-a5mekn6r\.googlevideo\.com\/videoplayback\?<redacted>\)\./,
  );
  assert.match(masked, /https:\/\/manifest\.googlevideo\.com\/videoplayback\?<redacted>#frag/);
  assert.doesNotMatch(masked, /1777777777/);
  assert.doesNotMatch(masked, /EI_VALUE/);
  assert.doesNotMatch(masked, /o-SECRETID/);
  assert.doesNotMatch(masked, /203\.0\.113\.42/);
  assert.doesNotMatch(masked, /NSECRET/);
  assert.doesNotMatch(masked, /SIGSECRET/);
  assert.doesNotMatch(masked, /LSIGSECRET/);
});

test('redactLogExportText redacts nested JSON secret values', () => {
  const masked = redactLogExportText(
    'json {"nested":{"token":123,"safe":"ok"},"items":[{"password":true},{"client_secret":null}]}',
  );

  assert.match(masked, /"token":"<redacted>"/);
  assert.match(masked, /"safe":"ok"/);
  assert.match(masked, /"password":"<redacted>"/);
  assert.match(masked, /"client_secret":"<redacted>"/);
  assert.doesNotMatch(masked, /123/);
  assert.doesNotMatch(masked, /true/);
  assert.doesNotMatch(masked, /null/);
});

test('redactLogExportText redacts IPv6 addresses with zone identifiers', () => {
  const masked = redactLogExportText('connected [fe80::1%en0]:443 and fe80::2%eth0');

  assert.equal(masked, 'connected [<ip>]:443 and <ip>');
});

test('redactLogExportText keeps malformed JSON scans bounded', () => {
  const malformedLine = `${'{'.repeat(20_000)} token=secret`;
  const start = performance.now();
  const masked = redactLogExportText(`${malformedLine}\njson {"token":"secret","safe":"ok"}`);
  const elapsedMs = performance.now() - start;

  assert.match(masked, /token=<redacted>/);
  assert.match(masked, /json {"token":"<redacted>","safe":"ok"}/);
  assert.doesNotMatch(masked, /token=secret/);
  assert.doesNotMatch(masked, /"token":"secret"/);
  assert.ok(elapsedMs < 100, `expected bounded scan under 100ms, got ${elapsedMs.toFixed(1)}ms`);
});
