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
    const session = await startE2eSession({ config: { controller: { enabled: true } } });
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

// Card-kind flags are written before media, so wait for the update carrying audio.
function findMediaUpdate(session: E2eSession) {
  return session.anki.requests.find(
    (request) =>
      request.action === 'updateNoteFields' && JSON.stringify(request.params).includes('[sound:'),
  );
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

// The other tests read the renderer over CDP, which works even if the window never appears.
// Windows keeps the overlay hidden until the tracker identifies mpv's window.
// Bun's node:test ignores the `skip` option, so this picks `test.skip` instead.
const testOnWindows = process.platform === 'win32' ? test : test.skip;
testOnWindows('visible overlay window appears over mpv', { timeout: TEST_TIMEOUT_MS }, async () => {
  const { session } = await useSession();
  await session.runAppCommand('--show-visible-overlay');
  const mpvPid = await session.mpv.command('get_property', 'pid');
  assert.equal(typeof mpvPid, 'number');
  const main = await session.mainProcess();
  await waitUntil(
    () =>
      main.evaluate<boolean>(
        `(() => {
          const electron = process.mainModule.require('electron');
          const overlay = electron.BrowserWindow.getAllWindows().find(
            (w) => w.webContents.getURL().includes(${JSON.stringify(OVERLAY_PAGE)}));
          if (!overlay?.isVisible() || overlay.getOpacity() !== 1 ||
              overlay.__subminerOverlayContentReady !== true) return false;
          // The inspector's mainModule can be Electron's bootstrap module. Resolve app
          // dependencies from the repository instead of that module's search paths.
          const appRequire = process.mainModule.require('node:module').createRequire(
            process.mainModule.require('node:path').resolve('package.json'));
          const koffi = appRequire('koffi');
          const getWindow = koffi.load('user32.dll').func(
            'intptr __stdcall GetWindow(intptr hwnd, uint uCmd)');
          const handle = overlay.getNativeWindowHandle();
          const hwnd = handle.length >= 8 ? Number(handle.readBigUInt64LE(0)) : handle.readUInt32LE(0);
          const owner = getWindow(hwnd, 4); // GW_OWNER
          const win32 = appRequire(
            process.mainModule.require('node:path').resolve('dist/window-trackers/win32.js'));
          return win32.findMpvWindows(${JSON.stringify(mpvPid)}).matches.some(
            (match) => match.hwnd === owner);
        })()`,
      ),
    { description: 'the visible overlay to be fully revealed and owned by the target mpv window' },
  );
});

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
    const update = await waitUntil(() => findMediaUpdate(session), {
      description: 'media fields to be written to the mined note',
      timeoutMs: 30_000,
    });

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

test(
  'mining a line retimed with sub-delay clips it where mpv plays it',
  { timeout: TEST_TIMEOUT_MS },
  async () => {
    const { session, overlay } = await useSession();
    const cue = FIXTURE_CUES[1];
    const delay = 2;
    await session.mpv.command('set_property', 'sub-delay', delay);
    try {
      // The previous test leaves this cue on screen, so blank the line first; otherwise the wait
      // below passes before the seek lands and mining can catch mpv's empty mid-seek sub-text.
      await session.mpv.command('seek', FIXTURE_CUES[0].start / 2, 'absolute+exact');
      await overlay.waitFor(`document.getElementById('subtitleRoot').textContent === ''`, {
        description: 'overlay to clear the line before the first delayed cue',
      });
      // Past the cue's own end, so the line is only on screen because of the delay.
      await session.mpv.command('seek', cue.end + delay / 2, 'absolute+exact');
      await overlay.waitFor(
        `document.getElementById('subtitleRoot').textContent === ${JSON.stringify(cue.text)}`,
        { description: 'overlay to show the delayed line' },
      );
      session.anki.requests.length = 0;

      await session.runAppCommand('--mine-sentence');
      await waitUntil(() => findMediaUpdate(session), {
        description: 'media fields to be written to the mined note',
        timeoutMs: 30_000,
      });

      // MiscInfo stamps the clip start: the cue's start on the video clock.
      const written = JSON.stringify(session.anki.requests);
      assert.ok(written.includes(`(00:00:0${cue.start + delay})`), written);
    } finally {
      await session.mpv.command('set_property', 'sub-delay', 0);
    }
  },
);

test(
  'open subtitle sidebar lists the parsed cues and follows playback',
  { timeout: TEST_TIMEOUT_MS },
  async () => {
    const { session, overlay } = await useSession();
    await showCue(FIXTURE_CUES[0]);
    await session.runAppCommand('--toggle-subtitle-sidebar');
    await overlay.waitFor(
      `document.querySelectorAll('#subtitleSidebarList .subtitle-sidebar-item').length === ${FIXTURE_CUES.length}`,
      { description: 'sidebar to list every fixture cue' },
    );

    for (const [index, cue] of FIXTURE_CUES.entries()) {
      await session.mpv.command('seek', cue.at, 'absolute+exact');
      await overlay.waitFor(
        `document.querySelector('#subtitleSidebarList .subtitle-sidebar-item.active')?.dataset.index === '${index}'`,
        { description: `sidebar to highlight cue ${index}` },
      );
    }

    await session.runAppCommand('--toggle-subtitle-sidebar');
    await overlay.waitFor(
      `document.getElementById('subtitleSidebarModal').classList.contains('hidden')`,
      { description: 'sidebar to close' },
    );
  },
);

test(
  'subtitle sidebar follows and seeks a sub-delay retimed track',
  { timeout: TEST_TIMEOUT_MS },
  async () => {
    const { session, overlay } = await useSession();
    const delay = 2;
    await session.mpv.command('set_property', 'sub-delay', delay);
    try {
      // At 9s cue 2 starts on the file's clock, but mpv is still showing cue 1.
      await session.mpv.command('seek', FIXTURE_CUES[2].start, 'absolute+exact');
      await overlay.waitFor(
        `document.getElementById('subtitleRoot').textContent === ${JSON.stringify(FIXTURE_CUES[1].text)}`,
        { description: 'overlay to show the delayed line' },
      );
      await session.runAppCommand('--toggle-subtitle-sidebar');
      await overlay.waitFor(
        `document.querySelector('#subtitleSidebarList .subtitle-sidebar-item.active')?.dataset.index === '1'`,
        { description: 'sidebar to highlight the delayed line' },
      );

      await overlay.evaluate(
        `document.querySelector('#subtitleSidebarList .subtitle-sidebar-item[data-index="2"]').click()`,
      );
      const target = FIXTURE_CUES[2].start + delay;
      await waitUntil(
        async () => {
          const time = Number(await session.mpv.command('get_property', 'time-pos'));
          return time >= target && time < target + 0.5;
        },
        { description: `click to seek to cue 2 at ${target}s on the video clock` },
      );

      await session.runAppCommand('--toggle-subtitle-sidebar');
      await overlay.waitFor(
        `document.getElementById('subtitleSidebarModal').classList.contains('hidden')`,
        { description: 'sidebar to close' },
      );
    } finally {
      await session.mpv.command('set_property', 'sub-delay', 0);
    }
  },
);

// macOS and Windows keep a hidden modal renderer alive, and hidden windows keep polling because
// background throttling is off. A fake pad goes to every renderer, as a real pad would.
// Linux has no such window, and under Xvfb without a window manager the overlay reports hidden,
// so requestAnimationFrame and with it the controller poll never run. Bun's node:test ignores the
// `skip` option, so this picks `test.skip` instead.
const testOutsideLinux = process.platform === 'linux' ? test.skip : test;
testOutsideLinux(
  'one controller press toggles mpv pause once across every window',
  { timeout: TEST_TIMEOUT_MS },
  async () => {
    const { session } = await useSession();
    const START_BUTTON = 9;
    const targets = (await listCdpTargets(session.cdpPort)).filter((target) =>
      target.url.includes('renderer/index.html'),
    );
    const renderers = await Promise.all(targets.map(connectCdpPage));
    const everyRenderer = (expression: string) =>
      Promise.all(renderers.map((renderer) => renderer.evaluate(expression)));
    const nextPolls = `new Promise((done) =>
      requestAnimationFrame(() => requestAnimationFrame(() => done(true))))`;
    try {
      await everyRenderer(`(() => {
        window.__fakePadPressed = false;
        Object.defineProperty(navigator, 'getGamepads', {
          configurable: true,
          value: () => [{
            id: 'E2E Pad (STANDARD GAMEPAD)',
            index: 0,
            connected: true,
            mapping: 'standard',
            axes: [0, 0, 0, 0],
            buttons: Array.from({ length: 17 }, (_, index) => {
              const pressed = window.__fakePadPressed && index === ${START_BUTTON};
              return { value: pressed ? 1 : 0, pressed, touched: pressed };
            }),
          }],
        });
        return true;
      })()`);
      await everyRenderer(nextPolls);
      await session.mpv.command('set_property', 'pause', true);

      await everyRenderer('window.__fakePadPressed = true');
      await everyRenderer(nextPolls);
      await everyRenderer('window.__fakePadPressed = false');
      await everyRenderer(nextPolls);
      // Long enough for a second `cycle pause` to reach mpv and cancel the first.
      await new Promise((done) => setTimeout(done, 300));

      assert.equal(await session.mpv.command('get_property', 'pause'), false);
    } finally {
      await everyRenderer('delete navigator.getGamepads');
      renderers.forEach((renderer) => renderer.close());
      await session.mpv.command('set_property', 'pause', true);
    }
  },
);

test(
  'controller combos survive the real save path and can be rebound to one button',
  { timeout: TEST_TIMEOUT_MS },
  async () => {
    const { overlay } = await useSession();
    const L1 = { kind: 'button', buttonIndex: 4 };
    const DPAD_LEFT = { kind: 'button', buttonIndex: 14 };
    const DPAD_RIGHT = { kind: 'button', buttonIndex: 15 };
    const save = (bindings: Record<string, unknown>) =>
      overlay.evaluate<Record<string, unknown>>(
        `(async () => {
          await window.electronAPI.saveControllerConfig(${JSON.stringify({ bindings })});
          return (await window.electronAPI.getControllerConfig()).bindings;
        })()`,
      );

    await save({ previousEntry: { ...DPAD_LEFT, hold: L1 } });
    const bothCombos = await save({ nextEntry: { ...DPAD_RIGHT, hold: L1 } });
    assert.deepEqual(bothCombos.previousEntry, { ...DPAD_LEFT, hold: L1 });
    assert.deepEqual(bothCombos.nextEntry, { ...DPAD_RIGHT, hold: L1 });

    const rebound = await save({ previousEntry: DPAD_LEFT, nextEntry: { kind: 'none' } });
    assert.deepEqual(rebound.previousEntry, DPAD_LEFT);
    assert.deepEqual(rebound.nextEntry, { kind: 'none' });

    await save({ previousEntry: { kind: 'none' } });
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
