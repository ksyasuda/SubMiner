// Run after bun run build. This uses a temporary profile and small local ZIPs.
const { app, BrowserWindow, protocol, session } = require('electron');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const http = require('node:http');
const root = path.resolve(__dirname, '..');
const parser = require(path.join(root, 'dist/core/services/tokenizer/yomitan-parser-runtime.js'));
const { writeStoredZip } = require(path.join(root, 'dist/shared/stored-zip.js'));
const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-hachidori-parser-'));

protocol.registerSchemesAsPrivileged([
  {
    scheme: 'chrome-extension',
    privileges: {
      standard: true,
      secure: true,
      supportFetchAPI: true,
      corsEnabled: true,
      bypassCSP: true,
    },
  },
]);
app.setPath('userData', profile);
app.setAppPath(root);
app.disableHardwareAcceleration();
app.on('window-all-closed', () => {});
const deadline = setTimeout(() => {
  console.error('Hachidori parser verification timed out', profile);
  app.exit(1);
}, 120_000);

function fixture(name, title, bankName, entries, frequencyMode = 'rank-based') {
  const zipPath = path.join(profile, name + '.zip');
  writeStoredZip(zipPath, [
    {
      name: 'index.json',
      data: Buffer.from(JSON.stringify({ title, revision: '1', format: 3, frequencyMode })),
    },
    { name: bankName, data: Buffer.from(JSON.stringify(entries)) },
  ]);
  return zipPath;
}

app
  .whenReady()
  .then(async () => {
    const targetSession = session.fromPartition('persist:hachidori-check');
    const extension = await targetSession.extensions.loadExtension(
      path.join(root, 'build/hachidori'),
      { allowFileAccess: true },
    );
    // Keep a host window alive while the dictionary importer opens/closes its
    // temporary settings windows, matching the running app's window lifecycle.
    const host = new BrowserWindow({
      show: false,
      webPreferences: { session: targetSession, contextIsolation: true, nodeIntegration: false },
    });
    await host.loadURL(`chrome-extension://${extension.id}/settings.html`);
    let parserWindow = null;
    let readyPromise = null;
    let initPromise = null;
    const deps = {
      getYomitanExt: () => extension,
      getYomitanSession: () => targetSession,
      getYomitanParserWindow: () => parserWindow,
      setYomitanParserWindow: (value) => {
        parserWindow = value;
      },
      getYomitanParserReadyPromise: () => readyPromise,
      setYomitanParserReadyPromise: (value) => {
        readyPromise = value;
      },
      getYomitanParserInitPromise: () => initPromise,
      setYomitanParserInitPromise: (value) => {
        initPromise = value;
      },
    };
    const errors = [];
    const logger = {
      error: (...args) => {
        errors.push(args);
        console.error(...args);
      },
    };
    assert.deepEqual(await parser.getYomitanDictionaryInfo(deps, logger), []);
    assert.equal(
      await parser.syncYomitanDefaultAnkiServer('http://127.0.0.1:18766', deps, logger, {
        forceOverride: true,
        deck: 'Test Mining',
      }),
      true,
    );
    // Exercise discovery and revisioned options writes in the real extension,
    // without creating a note or changing the user's Anki collection.
    const fields = ['Term', 'Reading', 'Definition', 'Context', 'Pronunciation', 'Image'];
    const metadataServer = http.createServer((request, response) => {
      let body = '';
      request.on('data', (chunk) => {
        body += chunk;
      });
      request.on('end', () => {
        const message = JSON.parse(body);
        assert.equal(message.action, 'multi');
        const result = message.params.actions.map(({ action }) => ({
          result:
            action === 'deckNames'
              ? ['Test Mining']
              : action === 'modelNames'
                ? ['Custom Japanese']
                : fields,
          error: null,
        }));
        response.setHeader('Content-Type', 'application/json');
        response.end(JSON.stringify({ result, error: null }));
      });
    });
    await new Promise((resolve) => metadataServer.listen(0, '127.0.0.1', resolve));
    try {
      const url = `http://127.0.0.1:${metadataServer.address().port}`;
      const ankiConfig = {
        tags: ['SubMiner', 'Autofill'],
        fields: {
          word: 'Term',
          sentence: 'Context',
          wordAudio: 'Pronunciation',
          image: 'Image',
        },
      };
      assert.equal(
        await parser.syncYomitanDefaultAnkiServer(url, deps, logger, {
          forceOverride: true,
          deck: 'Test Mining',
          ankiConfig,
        }),
        true,
      );
      const anki = await parserWindow.webContents.executeJavaScript(
        `(async () => (await chrome.storage.local.get('options')).options.anki)()`,
      );
      assert.equal(anki.model, 'Custom Japanese');
      assert.equal(anki.deck, 'Test Mining');
      assert.deepEqual(anki.tags, ankiConfig.tags);
      assert.equal(anki.fieldTemplates.Term.value, '{expression}');
      assert.equal(anki.fieldTemplates.Pronunciation.value, '{audio}');
      assert.equal(anki.fieldTemplates.Context.value, '{sentence}');
      assert.deepEqual(anki.templates[0].fieldTemplates, anki.fieldTemplates);
      console.log('Hachidori Anki auto-population passed with an empty dictionary library');
    } finally {
      await new Promise((resolve) => metadataServer.close(resolve));
    }
    const archives = [
      fixture('terms', 'SubMiner Test Terms', 'term_bank_1.json', [
        ['食べる', 'たべる', '', 'v1', 0, ['to eat'], 1, ''],
      ]),
      fixture('names', 'SubMiner Character Dictionary (AniList 1)', 'term_bank_1.json', [
        ['ミナト', 'みなと', '', 'n', 0, ['name'], 1, ''],
      ]),
      fixture('frequency', 'SubMiner Test Frequency', 'term_meta_bank_1.json', [
        ['食べる', 'freq', { reading: 'たべる', frequency: 42 }],
        ['頻度だけ', 'freq', { reading: 'ひんどだけ', frequency: 120 }],
        ['頻度だけ', 'freq', { reading: 'べつのよみ', frequency: 250 }],
        ['頻度だけ', 'freq', 17],
      ]),
      fixture(
        'occurrences',
        'SubMiner Test Occurrences',
        'term_meta_bank_1.json',
        [['頻度だけ', 'freq', 9000]],
        'occurrence-based',
      ),
    ];
    for (const archive of archives)
      assert.equal(await parser.importYomitanDictionaryFromZip(archive, deps, logger), true);
    const dictionaries = await parser.getYomitanDictionaryInfo(deps, logger);
    assert.equal(dictionaries.length, 4);
    parser.clearYomitanParserCachesForWindow(parserWindow);
    const tokens = await parser.requestYomitanScanTokens('ミナト 食べた', deps, logger, {
      includeNameMatchMetadata: true,
      currentCharacterDictionaryMediaId: 1,
    });
    assert.equal(tokens[0].isNameMatch, true);
    assert.equal(tokens[1].headword, '食べる');
    assert.equal(tokens[1].reading, 'たべた');
    assert.equal(tokens[1].startPos, 4);
    assert.equal(tokens[1].endPos, 7);
    assert.equal(tokens[1].frequencyRank, 42);
    assert.deepEqual(tokens[1].wordClasses, ['v1']);
    const exact = await parser.requestYomitanTermFrequencies(
      [{ term: '食べる', reading: 'たべる' }],
      deps,
      logger,
    );
    assert.equal(exact.length, 1);
    assert.equal(exact[0].frequency, 42);
    assert.equal(exact[0].reading, 'たべる');
    assert.equal(exact[0].hasReading, false);
    const otherReading = await parser.requestYomitanTermFrequencies(
      [{ term: '食べる', reading: 'べつのよみ' }],
      deps,
      logger,
    );
    // The shared frequency pipeline retries a missing reading as a term-only query.
    assert.equal(otherReading[0]?.frequency, 42);
    const unmatched = await parser.requestYomitanTermFrequencies(
      [{ term: '頻度だけ', reading: null }],
      deps,
      logger,
    );
    assert.deepEqual(unmatched, []);
    assert.equal(await parser.getYomitanCurrentAnkiDeckName(deps, logger), 'Test Mining');
    assert.equal((await targetSession.extensions.getAllExtensions()).length, 1);
    assert.equal(
      await parser.syncYomitanDefaultAnkiServer('http://127.0.0.1:8765', deps, logger),
      true,
    );
    const directSettings = await parser.getYomitanSettingsFull(deps, logger);
    assert.equal(directSettings.profiles[0].options.anki.server, 'http://127.0.0.1:8765');
    const proxyState = await host.webContents.executeJavaScript(
      "chrome.storage.local.get('subminerAnkiProxyUrl')",
    );
    assert.equal(proxyState.subminerAnkiProxyUrl, null);
    for (const entry of await parser.getYomitanDictionaryInfo(deps, logger)) {
      assert.equal(await parser.deleteYomitanDictionaryByTitle(entry.title, deps, logger), true);
    }
    assert.deepEqual(await parser.getYomitanDictionaryInfo(deps, logger), []);
    assert.equal(errors.length, 0);
    console.log(
      'PASS Hachidori native import, scanner, character names, term-entry API frequencies, settings and removal',
    );
    clearTimeout(deadline);
    app.exit(0);
  })
  .catch((error) => {
    console.error(error);
    clearTimeout(deadline);
    app.exit(1);
  });
