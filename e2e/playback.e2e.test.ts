import assert from 'node:assert/strict';
import test, { after, before } from 'node:test';
import type { CdpPage } from './harness/cdp';
import { FAKE_ANKI_DECK, FAKE_ANKI_MODEL } from './harness/fake-anki';
import { FIXTURE_CUES } from './harness/fixtures';
import { OVERLAY_PAGE, startE2eSession, type E2eSession } from './harness/session';
import { waitUntil } from './harness/wait';

let session: E2eSession;
let overlay: CdpPage;

before(
  async () => {
    session = await startE2eSession();
    overlay = await session.openPage(OVERLAY_PAGE);
  },
  { timeout: 120_000 },
);

after(() => session?.dispose(), { timeout: 30_000 });

type RenderedToken = { surface: string; headword: string | null };

// Parks playback inside a cue and waits for the overlay to show that line.
async function showCue(cue: (typeof FIXTURE_CUES)[number]): Promise<RenderedToken[]> {
  await session.mpv.command('seek', cue.at, 'absolute+exact');
  await overlay.waitFor(
    `document.getElementById('subtitleRoot').textContent === ${JSON.stringify(cue.text)}
      && document.querySelector('#subtitleRoot .word') !== null`,
    { description: `overlay to render "${cue.text}" tokenized` },
  );
  return overlay.evaluate<RenderedToken[]>(
    `Array.from(document.querySelectorAll('#subtitleRoot .word'), (word) => ({
      surface: word.textContent,
      headword: word.getAttribute('data-headword'),
    }))`,
  );
}

function hasToken(tokens: RenderedToken[], surface: string, headword: string): boolean {
  return tokens.some((token) => token.surface === surface && token.headword === headword);
}

test(
  'overlay renders the current mpv subtitle line as dictionary tokens',
  { timeout: 30_000 },
  async () => {
    const tokens = await showCue(FIXTURE_CUES[0]);

    assert.ok(hasToken(tokens, '今日', '今日'), JSON.stringify(tokens));
    assert.ok(hasToken(tokens, '天気', '天気'), JSON.stringify(tokens));
  },
);

test(
  'seeking to another cue replaces the line and deinflects verbs to their headword',
  { timeout: 30_000 },
  async () => {
    const tokens = await showCue(FIXTURE_CUES[2]);

    assert.ok(hasToken(tokens, '行きます', '行く'), JSON.stringify(tokens));
  },
);

test(
  'mining a sentence card sends the line with generated audio and image to Anki',
  { timeout: 60_000 },
  async () => {
    const cue = FIXTURE_CUES[1];
    await showCue(cue);
    session.anki.requests.length = 0;

    await session.runAppCommand('--mine-sentence');
    const update = await waitUntil(
      () => session.anki.requests.find((request) => request.action === 'updateNoteFields'),
      { description: 'media fields to be written to the mined note', timeoutMs: 30_000 },
    );

    const added = session.anki.requests.filter((request) => request.action === 'addNote');
    assert.equal(added.length, 1);
    assert.partialDeepStrictEqual(added[0]?.params.note, {
      deckName: FAKE_ANKI_DECK,
      modelName: FAKE_ANKI_MODEL,
      fields: { Sentence: cue.text },
    });

    const stored = session.anki.requests
      .filter((request) => request.action === 'storeMediaFile')
      .map((request) => String(request.params.filename));
    assert.ok(
      stored.some((name) => name.endsWith('.mp3')),
      `no audio in ${stored.join(', ')}`,
    );
    assert.ok(
      stored.some((name) => name.endsWith('.jpg')),
      `no image in ${stored.join(', ')}`,
    );

    const updatedFields = JSON.stringify(update.params);
    for (const filename of stored) assert.ok(updatedFields.includes(filename), filename);
  },
);

// Runs last: it leaves the lookup popup open over the overlay.
test(
  'holding Shift over a word opens the Yomitan popup with its dictionary entry',
  { timeout: 30_000 },
  async () => {
    await showCue(FIXTURE_CUES[0]);
    const word = await overlay.evaluate<{ x: number; y: number }>(
      `(() => {
        const words = Array.from(document.querySelectorAll('#subtitleRoot .word'));
        const box = words.find((word) => word.textContent === '天気').getBoundingClientRect();
        return { x: box.x + box.width / 2, y: box.y + box.height / 2 };
      })()`,
    );

    // Yomitan's default scan modifier is Shift, and it scans on mouse movement,
    // so approach the word, hold Shift, then move across it.
    const shift = { key: 'Shift', code: 'ShiftLeft', windowsVirtualKeyCode: 16 };
    const move = (x: number, modifiers = 0) =>
      overlay.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x, y: word.y, modifiers });
    await move(word.x - 30);
    await overlay.send('Input.dispatchKeyEvent', { type: 'rawKeyDown', ...shift, modifiers: 8 });
    await move(word.x - 2, 8);
    await move(word.x, 8);

    await overlay.waitFor(
      `document.querySelector('[data-subminer-yomitan-popup-visible="true"]') !== null`,
      { description: 'overlay to mark the Yomitan popup visible' },
    );
    const popup = await session.openPage('popup.html');
    const entry = await popup.waitFor<string>(
      `document.body.innerText.includes('weather') && document.body.innerText`,
      { description: 'popup to show the fixture dictionary entry' },
    );
    await overlay.send('Input.dispatchKeyEvent', { type: 'keyUp', ...shift });
    assert.match(entry, /天気/);
  },
);
