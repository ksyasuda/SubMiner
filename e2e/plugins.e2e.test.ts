import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test, { after } from 'node:test';
import type { CdpPage } from './harness/cdp';
import { FIXTURE_CUES } from './harness/fixtures';
import { OVERLAY_PAGE, startE2eSession, type E2eSession } from './harness/session';
import { waitUntil } from './harness/wait';

// Bun caps node:test hooks at 5s whatever their timeout option says, so the
// session boots inside whichever test runs first and every test budgets for it.
const TEST_TIMEOUT_MS = 120_000;

// Opens a panel describing the subtitle context and a child process result, and
// logs when the panel's signal aborts so the test can see the close reach main.
const PROBE_PLUGIN = `
const { ipcMain } = require('electron');

exports.activate = (ctx) => {
  ipcMain.handle('e2e-probe:info', () => ({ dataDir: ctx.dataDir, greeting: ctx.settings.greeting }));
  ctx.onDispose(() => ipcMain.removeHandler('e2e-probe:info'));
  ctx.media.on('change', (media) => ctx.log.info('probe media ' + JSON.stringify(media)));
  ctx.anki.on('card-mined', ({ notes }) => ctx.log.info('probe mined ' + JSON.stringify(notes)));
  ctx.actions.register('subtitle', () => ctx.log.info('probe subtitle ' + JSON.stringify(ctx.subtitles.current())));
  ctx.actions.register('request', async () => {
    const title = await ctx.mpv.request(['get_property', 'media-title']);
    ctx.log.info('probe request ' + JSON.stringify(title));
    try { await ctx.mpv.request(['subminer-no-such-command']); }
    catch { ctx.log.info('probe request rejected'); }
  });
  ctx.actions.register('window', () => {
    ctx.ui.openWindow({ html: 'ui/index.html', preload: 'preload.js', title: 'E2E Probe Window' });
  });
  ctx.mpv.on('end-file', ({ reason }) => ctx.log.info('probe end-file ' + reason));
  ctx.actions.register('show', () => {
    const { before, current, after } = ctx.subtitles.context({ before: 1, after: 1 });
    const panel = ctx.ui.panel({ title: 'Probe', markdown: '_working_' });
    panel.signal.addEventListener('abort', () => ctx.log.info('probe panel aborted'));
    ctx
      .exec(process.execPath, ['-e', 'process.stdout.write("child-ok")'], {
        env: { ELECTRON_RUN_AS_NODE: '1' },
        signal: panel.signal,
      })
      .then((result) =>
        panel.update({
          markdown: [
            '## ' + ctx.settings.greeting,
            '',
            '- before: ' + before.map((line) => line.text).join('|'),
            '- current: **' + (current ? current.text : '') + '**',
            '- after: ' + after.map((line) => line.text).join('|'),
            '- exec: ' + result.stdout,
          ].join('\\n'),
        }),
      )
      .catch((error) => panel.update({ markdown: 'exec failed: ' + error.message }));
  });
};
`;

const PROBE_PRELOAD = `
const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('probe', { info: () => ipcRenderer.invoke('e2e-probe:info') });
`;

const PROBE_WINDOW_HTML = `<!doctype html>
<html>
  <head><title>E2E Probe Window</title></head>
  <body>
    <pre id="out">loading</pre>
    <script>
      window.probe.info().then((info) => {
        document.getElementById('out').textContent = JSON.stringify(info);
      });
    </script>
  </body>
</html>
`;

type Started = { session: E2eSession; overlay: CdpPage };
let started: Promise<Started> | undefined;

function useSession(): Promise<Started> {
  started ??= (async () => {
    const session = await startE2eSession({
      config: { plugins: { 'e2e-probe': { enabled: true, settings: { greeting: 'hello' } } } },
      files: {
        'plugins/e2e-probe/plugin.json': JSON.stringify({ name: 'E2E Probe', main: 'main.cjs' }),
        'plugins/e2e-probe/main.cjs': PROBE_PLUGIN,
        'plugins/e2e-probe/preload.js': PROBE_PRELOAD,
        'plugins/e2e-probe/ui/index.html': PROBE_WINDOW_HTML,
      },
    });
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

type PanelState = {
  title: string;
  heading: string;
  current: string;
  items: string[];
};

test(
  'a plugin action opens the plugin panel with subtitle context and child process output',
  { timeout: TEST_TIMEOUT_MS },
  async () => {
    const { session, overlay } = await useSession();
    const [previous, cue, next] = FIXTURE_CUES;
    await session.mpv.command('seek', cue.at, 'absolute+exact');
    await overlay.waitFor(
      `document.getElementById('subtitleRoot').textContent === ${JSON.stringify(cue.text)}`,
      { description: `overlay to show "${cue.text}"` },
    );

    await session.runAppCommand(
      '--session-action',
      JSON.stringify({
        actionId: 'runPluginAction',
        payload: { pluginActionId: 'e2e-probe/show' },
      }),
    );

    const modal = await session.openPage('renderer/index.html?layer=modal');
    const panel = await modal
      .waitFor<PanelState | false>(
        `(() => {
        const root = document.getElementById('pluginPanelModal');
        const body = document.getElementById('pluginPanelBody');
        if (root.classList.contains('hidden') || !/exec[: ]/.test(body.textContent)) return false;
        return {
          title: document.getElementById('pluginPanelTitle').textContent,
          heading: body.querySelector('h4')?.textContent ?? '',
          current: body.querySelector('strong')?.textContent ?? '',
          items: Array.from(body.querySelectorAll('li'), (item) => item.textContent),
        };
      })()`,
        { description: 'plugin panel to show the finished probe result' },
      )
      .catch(async (error: unknown) => {
        const body = await modal.evaluate<string>(
          `document.getElementById('pluginPanelBody').textContent`,
        );
        throw new Error(`Plugin panel body: ${JSON.stringify(body)}`, { cause: error });
      });

    assert.deepEqual(panel, {
      title: 'Probe',
      heading: 'hello',
      current: cue.text,
      items: [
        `before: ${previous.text}`,
        `current: ${cue.text}`,
        `after: ${next.text}`,
        'exec: child-ok',
      ],
    });

    // Closing hides the modal window (Linux destroys it), often before Chromium
    // acknowledges the key event, so the send is not awaited. The plugin's log line
    // proves the close reached main and aborted the handle.
    void modal
      .send('Input.dispatchKeyEvent', {
        type: 'rawKeyDown',
        key: 'Escape',
        code: 'Escape',
        windowsVirtualKeyCode: 27,
      })
      .catch(() => {});
    await waitUntil(
      () =>
        fs.readFileSync(path.join(session.root, 'app.log'), 'utf8').includes('probe panel aborted'),
      { description: 'closing the panel to abort the plugin handle in main' },
    );
  },
);

test(
  'a CLI plugin action opens a plugin window whose preload talks to the plugin',
  { timeout: TEST_TIMEOUT_MS },
  async () => {
    const { session } = await useSession();

    await session.runAppCommand('--plugin-action', 'e2e-probe/window');

    const page = await session.openPage('plugins/e2e-probe/ui/index.html');
    const info = await page.waitFor<{ dataDir: string; greeting: string } | false>(
      `(() => {
        const text = document.getElementById('out')?.textContent;
        return text && text !== 'loading' ? JSON.parse(text) : false;
      })()`,
      { description: 'plugin window to read plugin info over IPC' },
    );

    assert.equal(info.greeting, 'hello');
    assert.equal(path.basename(info.dataDir), 'e2e-probe');
    assert.equal(path.basename(path.dirname(info.dataDir)), 'plugin-data');
    assert.equal(fs.statSync(info.dataDir).isDirectory(), true);
  },
);

test('plugins receive mpv end-file events', { timeout: TEST_TIMEOUT_MS }, async () => {
  const { session } = await useSession();
  const mediaPath = await session.mpv.command('get_property', 'path');

  // Replacing the playing file ends it with reason "stop".
  await session.mpv.command('loadfile', String(mediaPath), 'replace');

  await waitUntil(
    () =>
      fs.readFileSync(path.join(session.root, 'app.log'), 'utf8').includes('probe end-file stop'),
    { description: 'the plugin to log the end-file reason' },
  );
  // time-pos is unavailable from end-file until the replacement opens, so this waits for
  // the reload to finish. Otherwise a later sub-add can land on the outgoing file and vanish.
  await waitUntil(
    () =>
      session.mpv.command('get_property', 'time-pos').then(
        () => true,
        () => false,
      ),
    { description: 'the replacement file to load' },
  );
});

test(
  'plugin subtitle lookup uses the subtitle clock for repeated retimed lines',
  { timeout: TEST_TIMEOUT_MS },
  async () => {
    const { session, overlay } = await useSession();
    // Unlike any fixture cue, so the overlay showing it proves main received this line rather
    // than still displaying one from an earlier test.
    const text = '今日は猫が好きです。';
    const subtitlePath = path.join(session.root, 'repeated.srt');
    fs.writeFileSync(
      subtitlePath,
      `1\n00:00:01,000 --> 00:00:03,000\n${text}\n\n2\n00:00:06,000 --> 00:00:08,000\n${text}\n`,
    );
    const originalSid = await session.mpv.command('get_property', 'sid');
    // Runs the probe action and returns the cue it logged on its nth run.
    const reportSubtitle = async (run: number): Promise<unknown> => {
      await session.runAppCommand('--plugin-action', 'e2e-probe/subtitle');
      const line = await waitUntil(
        () =>
          fs
            .readFileSync(path.join(session.root, 'app.log'), 'utf8')
            .split('\n')
            .filter((entry) => entry.includes('probe subtitle '))[run - 1],
        { description: `the plugin to report its current subtitle (run ${run})` },
      );
      assert.ok(line);
      return JSON.parse(line.slice(line.indexOf('probe subtitle ') + 'probe subtitle '.length));
    };
    try {
      await session.mpv.command('seek', 0, 'absolute+exact');
      await session.mpv.command('sub-add', subtitlePath, 'select');
      await session.mpv.command('set_property', 'sub-delay', 5);
      await session.runAppCommand('--toggle-subtitle-sidebar');
      await overlay.waitFor(
        `document.querySelectorAll('#subtitleSidebarList .subtitle-sidebar-item').length === 2`,
        { description: 'the repeated subtitle cues to be parsed' },
      );
      await session.mpv.command('seek', 7, 'absolute+exact');
      await overlay.waitFor(
        `document.getElementById('subtitleRoot').textContent === ${JSON.stringify(text)}`,
        { description: 'the first delayed line to be visible' },
      );
      assert.deepEqual(await reportSubtitle(1), { text, startTime: 1, endTime: 3 });

      // Same text again at 6-8 on the subtitle clock, so only the timings tell the cues apart.
      await session.mpv.command('seek', 12, 'absolute+exact');
      await overlay.waitFor(
        `document.querySelectorAll('#subtitleSidebarList .subtitle-sidebar-item')[1]?.classList.contains('active') === true`,
        { description: 'the second delayed line to become active' },
      );
      assert.deepEqual(await reportSubtitle(2), { text, startTime: 6, endTime: 8 });
    } finally {
      await session.mpv.command('set_property', 'sub-delay', 0);
      await session.mpv.command('set_property', 'sid', originalSid);
      await session.runAppCommand('--toggle-subtitle-sidebar');
    }
  },
);

test(
  'plugin requests report command failures and mined notes carry the captured source',
  { timeout: TEST_TIMEOUT_MS },
  async () => {
    const { session, overlay } = await useSession();
    const cue = FIXTURE_CUES[1];
    await session.mpv.command('seek', cue.at, 'absolute+exact');
    await overlay.waitFor(
      `document.getElementById('subtitleRoot').textContent === ${JSON.stringify(cue.text)}`,
      { description: 'the mined subtitle to be visible' },
    );
    await session.runAppCommand('--plugin-action', 'e2e-probe/request');
    await waitUntil(
      () =>
        fs
          .readFileSync(path.join(session.root, 'app.log'), 'utf8')
          .includes('probe request rejected'),
      { description: 'mpv to reject an invalid plugin request' },
    );
    await session.runAppCommand('--mine-sentence');
    const log = await waitUntil(
      () => {
        const text = fs.readFileSync(path.join(session.root, 'app.log'), 'utf8');
        return text.includes('probe mined ') ? text : null;
      },
      { description: 'the mining event to reach the plugin' },
    );
    assert.ok(log);
    const line = log.split('\n').find((entry) => entry.includes('probe mined '));
    assert.ok(line);
    const notes: unknown = JSON.parse(
      line.slice(line.indexOf('probe mined ') + 'probe mined '.length),
    );
    assert.ok(Array.isArray(notes));
    assert.equal(notes.length, 1);
    const note: unknown = notes[0];
    assert.partialDeepStrictEqual(note, {
      provenance: {
        subtitle: { text: cue.text },
        media: { path: await session.mpv.command('get_property', 'path') },
      },
    });
    assert.ok(note && typeof note === 'object' && 'originalNoteId' in note && 'noteId' in note);
    assert.equal(note.originalNoteId, note.noteId);
  },
);
