import assert from 'node:assert/strict';
import test, { after } from 'node:test';
import { connectCdpPage, listCdpTargets, type CdpPage } from './harness/cdp';
import { FAKE_ANKI_DECK, FAKE_ANKI_MODEL } from './harness/fake-anki';
import { FIXTURE_CUES } from './harness/fixtures';
import { OVERLAY_PAGE, startE2eSession, type E2eSession } from './harness/session';
import { waitUntil } from './harness/wait';

type Started = { session: E2eSession; overlay: CdpPage };

// Bun caps node:test hooks at 5s whatever their timeout option says, so the
// session boots inside whichever test runs first and every test budgets for it.
const TEST_TIMEOUT_MS = 120_000;
let started: Promise<Started> | undefined;

function useSession(): Promise<Started> {
  started ??= (async () => {
    const session = await startE2eSession();
    try {
      return { session, overlay: await session.openPage(OVERLAY_PAGE) };
    } catch (error) {
      await session.dispose();
      throw error;
    }
  })();
  return started;
}

after(async () => {
  const running = await started?.catch(() => undefined);
  await running?.session.dispose();
});

type RenderedToken = { surface: string; headword: string | null };

// Parks playback inside a cue and waits for the overlay to show that line.
async function showCue(cue: (typeof FIXTURE_CUES)[number]): Promise<RenderedToken[]> {
  const { session, overlay } = await useSession();
  await session.mpv.command('seek', cue.at, 'absolute+exact');
  // One expression reads the line and its tokens together, so a re-render
  // between two round trips cannot hand back a half-updated subtitle.
  return overlay
    .waitFor<RenderedToken[] | false>(
      `(() => {
      const root = document.getElementById('subtitleRoot');
      const words = Array.from(root.querySelectorAll('.word'), (word) => ({
        surface: word.textContent,
        headword: word.getAttribute('data-headword'),
      }));
      return root.textContent === ${JSON.stringify(cue.text)} && words.length > 0 && words;
    })()`,
      { description: `overlay to render "${cue.text}" tokenized` },
    )
    .catch(async (error: unknown) => {
      // Say what the overlay was showing, what main thinks the line is, and
      // what the renderer logged, so a dropped update can be placed.
      const state = await overlay.evaluate<string>(
        `(async () => JSON.stringify({
          visibility: document.visibilityState,
          text: document.getElementById('subtitleRoot').textContent,
          html: document.getElementById('subtitleRoot').innerHTML.slice(0, 200),
          mainLine: (await window.electronAPI.getCurrentSubtitle())?.text,
          errorToast: document.getElementById('overlayErrorToast')?.textContent,
        }))()`,
      );
      // Diagnostics must not replace the cue failure if the inspector is gone too.
      const windows = await session
        .mainProcess()
        .then((main) =>
          main.evaluate<string>(
            `JSON.stringify(process.mainModule.require('electron').BrowserWindow.getAllWindows().map((w) => ({
              id: w.id,
              title: w.getTitle(),
              url: w.webContents.getURL().split('/').pop(),
              visible: w.isVisible(),
              loading: w.webContents.isLoading(),
              documentLoaded: w.__subminerOverlayDocumentLoaded,
              contentReady: w.__subminerOverlayContentReady,
            })))`,
          ),
        )
        .catch((inspectorError: unknown) => `unavailable (${String(inspectorError)})`);
      const targets = (await listCdpTargets(session.cdpPort))
        .filter((t) => t.type === 'page')
        .map((t) => t.url.split('/').pop());
      const log = overlay.console.slice(-20).join('\n');
      throw new Error(
        `Overlay state: ${state}\nMain windows: ${windows}\nPage targets: ${JSON.stringify(targets)}\nRenderer console:\n${log}`,
        { cause: error },
      );
    });
}

function hasToken(tokens: RenderedToken[], surface: string, headword: string): boolean {
  return tokens.some((token) => token.surface === surface && token.headword === headword);
}

test(
  'overlay renders the current mpv subtitle line as dictionary tokens',
  { timeout: TEST_TIMEOUT_MS },
  async () => {
    const tokens = await showCue(FIXTURE_CUES[0]);

    assert.ok(hasToken(tokens, '今日', '今日'), JSON.stringify(tokens));
    assert.ok(hasToken(tokens, '天気', '天気'), JSON.stringify(tokens));
  },
);

test(
  'seeking to another cue replaces the line and deinflects verbs to their headword',
  { timeout: TEST_TIMEOUT_MS },
  async () => {
    const tokens = await showCue(FIXTURE_CUES[2]);

    assert.ok(hasToken(tokens, '行きます', '行く'), JSON.stringify(tokens));
  },
);

test(
  'mining a sentence card sends the line with generated audio and image to Anki',
  { timeout: TEST_TIMEOUT_MS },
  async () => {
    const { session } = await useSession();
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
  { timeout: TEST_TIMEOUT_MS },
  async () => {
    const { session, overlay } = await useSession();
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
    // Yomitan keeps more than one popup frame around, so read whichever
    // one holds the entry instead of trusting target order.
    let seen: string[] = [];
    const entry = await waitUntil(
      async () => {
        const targets = (await listCdpTargets(session.cdpPort)).filter((target) =>
          target.url.includes('popup.html'),
        );
        seen = [];
        for (const target of targets) {
          // A frame can go away between listing and connecting; skip it.
          const frame = await connectCdpPage(target).catch(() => undefined);
          if (!frame) continue;
          // textContent, not innerText: the entry counts once it is in the DOM,
          // even where a GPU-less runner never lays the frame out.
          const text = await frame.evaluate<string>('document.body?.textContent ?? ""');
          frame.close();
          seen.push(text);
        }
        return seen.find((text) => text.includes('weather'));
      },
      { description: 'a popup frame to show the fixture dictionary entry' },
    ).catch((error: unknown) => {
      const summary = seen.map((text) => text.replace(/\s+/g, ' ').trim().slice(0, 200));
      throw new Error(`Popup frames held: ${JSON.stringify(summary)}`, { cause: error });
    });
    await overlay.send('Input.dispatchKeyEvent', { type: 'keyUp', ...shift });
    assert.match(entry, /天気/);
  },
);
