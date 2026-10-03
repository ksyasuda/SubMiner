// Hachidori exposes its own storage and engine protocol. Adapt it only inside
// SubMiner's hidden extension windows so the shared scanner keeps its matching,
// character-name and frequency semantics without changing Hachidori's pages.
import { HACHIDORI_ANKI_SETTINGS_SCRIPT } from './hachidori-anki-settings';

export const HACHIDORI_SESSION_PARTITION = 'persist:hachidori';

export const HACHIDORI_PARSER_BRIDGE_SCRIPT = String.raw`
  (async () => {
    if (globalThis.__subminerDictionarySendMessage) return;
    const { createApiHost } = await import('./api-host.js');
    await import('./reader-options.js');
    const send = async (type, fields = {}, target = 'hoshidicts-offscreen') => {
      const reply = await chrome.runtime.sendMessage({
        target, type, requestId: crypto.randomUUID(), ...fields,
      });
      if (!reply || reply.ok !== true) {
        throw new Error(reply?.error || 'Hachidori returned an invalid response');
      }
      return reply;
    };
    // Hachidori reports an empty library while its dictionaries load and rejects
    // engine requests during imports. Wait for a settled engine so SubMiner never
    // caches an empty dictionary list or a token-less scan.
    const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
    const waitForEngineReady = async () => {
      const deadline = Date.now() + 120000;
      for (;;) {
        const reply = await send('hd_status').catch(() => null);
        if (reply?.ready === true && reply.loading !== true) return;
        if (Date.now() >= deadline) {
          throw new Error(reply?.ready === true
            ? 'Hachidori is still importing dictionaries' : 'Hachidori dictionary engine is not ready');
        }
        await sleep(250);
      }
    };
    const engine = async (type, fields = {}, target) => {
      for (let attempt = 0; ; attempt += 1) {
        await waitForEngineReady();
        try {
          return await send(type, fields, target);
        } catch (error) {
          if (attempt >= 3 || !/busy mutating/.test(error.message)) throw error;
          await sleep(250);
        }
      }
    };
    const readState = async () => (await engine('hd_state_read', {}, 'hoshidicts-worker')).state
      ?? { revision: 0, dictionaries: [] };
    const readOptions = async () => {
      const stored = (await chrome.storage.local.get('options')).options;
      return { ...globalThis.HDReaderOptions.normaliseOptions(stored), revision: stored?.revision ?? 0 };
    };
    globalThis.__subminerSetAnkiProxyUrl = async url => {
      const previous = (await chrome.storage.local.get('subminerAnkiProxyUrl')).subminerAnkiProxyUrl;
      await chrome.storage.local.set({ subminerAnkiProxyUrl: url });
      return typeof previous === 'string' ? previous : null;
    };
    ${HACHIDORI_ANKI_SETTINGS_SCRIPT}
    const api = createApiHost({
      engine: fields => engine(fields.type, fields),
      render: fields => send(fields.type, fields, 'hachidori-anki-render'),
      readDictionaries: async () => (await readState()).dictionaries,
      readAudioSources: async () => (await readOptions()).audioSources.filter(source => source.enabled),
      version: chrome.runtime.getManifest().version,
    });
    // Snapshot of the last projection SubMiner received, so a later write can
    // be reduced to the fields SubMiner actually changed.
    let lastProjection = null;
    async function optionsGetFull() {
      const [state, options] = await Promise.all([readState(), readOptions()]);
      const revisions = { dictionaries: state.revision, options: options.revision };
      lastProjection = {
        revisions,
        enabledByTitle: Object.fromEntries(state.dictionaries.map(entry => [entry.title, entry.enabled !== false])),
        server: options.anki.url,
        deckByFormat: Object.fromEntries(options.anki.templates.map(template => [template.id, template.deck])),
      };
      return {
        profileCurrent: 0,
        hachidoriRevisions: revisions,
        profiles: [{ name: 'Hachidori', options: {
          scanning: { length: options.scanLength },
          dictionaries: state.dictionaries.map((entry, index) => ({
            name: entry.title, alias: entry.displayName || entry.title,
            id: index, enabled: entry.enabled !== false,
          })),
          anki: {
            server: options.anki.url,
            cardFormats: options.anki.templates.map(template => ({
              id: template.id, type: 'term', enabled: true, deck: template.deck,
            })),
          },
        } }],
      };
    }
    async function setAllSettings(value) {
      const projected = value.profiles?.[0]?.options;
      const revisions = value.hachidoriRevisions;
      if (!projected || !revisions) throw new Error('Invalid Hachidori settings projection');
      const base = lastProjection
        && lastProjection.revisions.dictionaries === revisions.dictionaries
        && lastProjection.revisions.options === revisions.options ? lastProjection : null;
      const enabledByTitle = Object.fromEntries(projected.dictionaries.map(item => [item.name, item.enabled === true]));
      const deckByFormat = Object.fromEntries(projected.anki.cardFormats.map(format => [format.id, format.deck]));
      for (let attempt = 0; ; attempt += 1) {
        const [state, options] = await Promise.all([readState(), readOptions()]);
        const current = state.revision === revisions.dictionaries && options.revision === revisions.options;
        if (!current && !base) {
          throw new Error('Hachidori settings changed while SubMiner was updating them; retry the action');
        }
        // Once Hachidori has moved on, apply only SubMiner's own changes on top
        // of the newer settings instead of replaying the stale projection.
        const enabledFor = title => enabledByTitle[title] === true;
        const dictionaries = state.dictionaries.map(entry =>
          current || (base.enabledByTitle[entry.title] === true) !== enabledFor(entry.title)
            ? { ...entry, enabled: enabledFor(entry.title) } : entry);
        const templates = options.anki.templates.map(template => {
          const deck = deckByFormat[template.id];
          return deck !== undefined && (current || base.deckByFormat[template.id] !== deck)
            ? { ...template, deck } : template;
        });
        const url = current || base.server !== projected.anki.server ? projected.anki.server : options.anki.url;
        const anki = { ...options.anki, url, templates, deck: templates[0]?.deck ?? options.anki.deck };
        try {
          if (JSON.stringify(dictionaries) !== JSON.stringify(state.dictionaries)) {
            await engine('hd_apply_state', { baseRevision: state.revision, dictionaries });
          }
          if (JSON.stringify(anki) !== JSON.stringify(options.anki)) {
            await engine('hd_options_write', { baseRevision: options.revision, options: { anki } }, 'hoshidicts-worker');
          }
          return true;
        } catch (error) {
          // A write raced another Hachidori change; re-read once and reapply the delta.
          if (attempt > 0 || !base) throw error;
        }
      }
    }
    async function getTermFrequencies({ termReadingList, dictionaries }) {
      const terms = [...new Set(termReadingList.map(pair => pair.term))];
      const { results } = await api({ type: 'hd_api_term_entries', terms });
      const frequencies = [];
      for (const result of results) {
        const term = terms[result.index];
        const pairs = termReadingList.filter(pair => pair.term === term);
        for (const entry of result.dictionaryEntries) {
          for (const value of entry.frequencies) {
            const headword = entry.headwords[value.headwordIndex];
            if (!headword || headword.term !== term || !dictionaries.includes(value.dictionary)) continue;
            if (!pairs.some(pair => pair.reading === null || pair.reading === headword.reading)) continue;
            // Upstream does not expose the frequency entry's original reading.
            // Keep its API flag and associate the value with the matched headword.
            frequencies.push({ term, reading: headword.reading || null,
              hasReading: value.hasReading, dictionary: value.dictionary,
              frequency: value.frequency, displayValue: value.displayValue,
              displayValueParsed: value.displayValueParsed });
          }
        }
      }
      return frequencies;
    }
    // Hachidori's public tokenize API emits display furigana without headwords.
    // SubMiner's fallback requires one group per token and a dictionary form.
    async function parseText({ text, scanLength }) {
      const content = [];
      let position = 0;
      while (position < text.length) {
        const rest = text.slice(position);
        const reply = await engine('hd_lookup', { text: rest, maxResults: 1, scanLength });
        const result = reply.results[0];
        if (!result?.matched || !rest.startsWith(result.matched)) {
          const character = String.fromCodePoint(rest.codePointAt(0));
          content.push([{ text: character, reading: '' }]);
          position += character.length;
          continue;
        }
        const term = result.term;
        // Keep the inflected ending in the reading, just as the main scanner does.
        let stem = 0;
        while (stem < term.expression.length && stem < result.matched.length && term.expression[stem] === result.matched[stem]) stem += 1;
        const ending = term.expression.slice(stem);
        const reading = stem > 0 && term.reading.endsWith(ending)
          ? term.reading.slice(0, term.reading.length - ending.length) + result.matched.slice(stem)
          : term.reading;
        content.push([{ text: result.matched, reading, headwords: [[{ term: term.expression }]] }]);
        position += result.matched.length;
      }
      return [{ source: 'scanning-parser', index: 0, content }];
    }
    async function invoke(action, params) {
      switch (action) {
        case 'optionsGetFull': return optionsGetFull();
        case 'setAllSettings': return setAllSettings(params.value);
        case 'getDictionaryInfo': return (await readState()).dictionaries.map(entry => ({
          title: entry.title, revision: entry.revision, frequencyMode: entry.frequencyMode,
        }));
        case 'termsFind': {
          const reply = await api({ type: 'hd_api_term_entries', terms: [params.text] });
          return reply.results[0];
        }
        case 'parseText': return parseText(params);
        case 'getTermFrequencies': return getTermFrequencies(params);
        default: throw new Error('Unsupported Hachidori parser action: ' + action);
      }
    }
    globalThis.__subminerDictionarySendMessage = ({ action, params }, callback) => {
      void invoke(action, params).then(result => callback({ result }), error => callback({ error: { message: error.message } }));
    };
    async function importArchive(blob, fileName) {
      const blobUrl = URL.createObjectURL(blob);
      try {
        const reply = await engine('hd_import', { blobUrl, fileName });
        if (reply.report?.success !== true) throw new Error(reply.report?.error || 'Hachidori dictionary import failed');
      } finally {
        URL.revokeObjectURL(blobUrl);
      }
    }
    globalThis.__subminerYomitanSettingsAutomation = {
      ready: true,
      async importDictionaryArchiveUrl(url) {
        const response = await fetch(url);
        if (!response.ok) throw new Error('Could not read the dictionary archive');
        await importArchive(await response.blob(), 'subminer-dictionary.zip');
      },
      async importDictionaryArchiveBase64(base64, fileName) {
        const bytes = Uint8Array.from(atob(base64), character => character.charCodeAt(0));
        await importArchive(new Blob([bytes], { type: 'application/zip' }), fileName);
      },
      async deleteDictionary(title) {
        const dictionary = (await readState()).dictionaries.find(entry => entry.title === title);
        if (dictionary) await engine('hd_remove', { id: dictionary.id, title });
      },
    };
    globalThis.__subminerAddNote = async word => {
      const lookup = await engine('hd_lookup', { text: word, maxResults: 1 });
      const result = lookup.results[0];
      if (!result) return { noteId: null, duplicateNoteIds: [] };
      // Match the dictionary context supplied by Hachidori's popup to its
      // native glossary, alias, and frequency template renderers.
      const { dictionaries } = await readState();
      const frequencyModes = new Map(dictionaries.map(entry => [entry.title, entry.frequencyMode]));
      const term = { ...result.term, frequencies: result.term.frequencies.map(group =>
        ({ ...group, frequencyMode: frequencyModes.get(group.dictionary) })) };
      const status = await send('hd_anki_status', {}, 'hachidori-anki');
      const request = {
        ...result, term, generation: lookup.generation, sentence: word, searchQuery: word,
        matchOffset: 0, documentTitle: 'SubMiner', popupSelectionText: '',
        configKey: status.configKey, subminerEnrich: false,
        dictionaryAliases: Object.fromEntries(dictionaries.filter(entry => entry.displayName)
          .map(entry => [entry.title, entry.displayName])),
        dictionaryIds: Object.fromEntries(dictionaries.map(entry => [entry.title, entry.id])),
        frequencyDictionaries: dictionaries.filter(entry => entry.enabled !== false && entry.frequencyCount > 0)
          .map(entry => entry.title),
        captureUnavailable: ['screenshot', 'animation', 'audio'],
      };
      const preflight = await send('hd_anki_preflight', { request }, 'hachidori-anki');
      if (preflight.canAdd !== true) {
        if (preflight.state !== 'duplicate') throw new Error(preflight.error || 'Hachidori Anki mining is unavailable');
        return { noteId: null, duplicateNoteIds: preflight.noteIds ?? [] };
      }
      const submitted = await send('hd_anki_submit', { request }, 'hachidori-anki');
      if (submitted.state === 'added' || submitted.state === 'updated') {
        return { noteId: submitted.noteId, duplicateNoteIds: [] };
      }
      if (submitted.state === 'duplicate') return { noteId: null, duplicateNoteIds: submitted.noteIds ?? [] };
      throw new Error(submitted.error || 'Hachidori could not confirm the note write');
    };
  })();
`;
