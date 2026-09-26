import fs from 'node:fs';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const source = path.join(repoRoot, 'vendor', 'hachidori');
const output = path.join(repoRoot, 'build', 'hachidori');
const provenance = JSON.parse(fs.readFileSync(path.join(source, 'SOURCE.json'), 'utf8'));

// Upstream commits the engine binaries. Verify the pinned bytes before staging
// so ordinary app builds need neither Emscripten nor network access.
for (const [file, expected] of Object.entries(provenance.artifacts)) {
  const actual = createHash('sha256')
    .update(fs.readFileSync(path.join(source, file)))
    .digest('hex');
  if (actual !== expected) throw new Error(`Hachidori artifact checksum mismatch: ${file}`);
}
const extension = path.join(source, 'extension');
const manifest = JSON.parse(fs.readFileSync(path.join(extension, 'manifest.json'), 'utf8'));
for (const file of [
  manifest.options_page,
  manifest.background.service_worker,
  'offscreen.html',
  ...manifest.content_scripts.flatMap(({ js, css }) => [...js, ...css]),
]) {
  if (!fs.existsSync(path.join(extension, file)))
    throw new Error(`Missing Hachidori asset: ${file}`);
}
fs.rmSync(output, { recursive: true, force: true });
fs.mkdirSync(output, { recursive: true });
fs.cpSync(extension, output, { recursive: true });
// Host configuration belongs in the staged copy, leaving the fork usable in Chrome.
const hostConfiguration = {
  'overlay-mode.js': [
    ['export const OVERLAY_MODE = false;', 'export const OVERLAY_MODE = true;'],
    ['customJavaScript: !IS_FIREFOX,', 'customJavaScript: false,'],
  ],
  // Overlay hosts seed the lookup highlight off because their Anki screenshot is
  // the see-through viewport. SubMiner captures media from mpv, and without the
  // highlight the sidebar shows nothing for the word being looked up.
  'setup-state.js': [
    [
      'lookupMode: "hover",\n  sourceHighlightEnabled: false,',
      'lookupMode: "hover",\n  sourceHighlightEnabled: true,',
    ],
  ],
};
for (const [file, replacements] of Object.entries(hostConfiguration)) {
  const filePath = path.join(output, file);
  let text = fs.readFileSync(filePath, 'utf8');
  for (const [original, replacement] of replacements) {
    if (!text.includes(original))
      throw new Error(`Hachidori host configuration changed upstream: ${original}`);
    text = text.replace(original, replacement);
  }
  fs.writeFileSync(filePath, text);
}
manifest.permissions = manifest.permissions.filter((permission) => permission !== 'userScripts');
// A fixed key pins the extension ID, so Hachidori's storage origin (dictionaries,
// settings, setup state) survives userData path changes. Electron otherwise
// derives the ID from the unpacked directory path.
manifest.key =
  'MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA40zTTbhCbrKnDOqj17tosDChP60VXuuMjR0tFFR/kAet21jrYn0JKPzqqNpXMGaJsx1chnuMOgaP0zuOT+gv9DNVV83wGLhQL7cC4LCoGN648WqnE2jYqBPj6hdD5A/N3uPeKFRAHkiC9HKoj2Et6sYZFTqag79QqQkkUz17X2VpjuTGmTP3lvhyEwhJWQKzOU0k6dTn6fKB2ZSqTebpw/JS2G8jf58vHlQOcfLhIiolKCG9DfMuo11aoG1s3AsFmQmw9VVyp4wgyQ80UI4lyBdqboanJxEiZ211+HKWPEgtG6VML1+Uo59qWS3uBU1apxE1+7E51UBEHaTsxajFIQIDAQAB';
fs.writeFileSync(path.join(output, 'manifest.json'), JSON.stringify(manifest, null, 2) + '\n');
for (const file of ['LICENSE', 'SOURCE.json', 'README.md']) {
  fs.copyFileSync(path.join(source, file), path.join(output, file));
}
process.stdout.write(`Hachidori ${provenance.revision} staged in ${output}\n`);
