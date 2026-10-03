import type { AnkiConnectConfig } from '../../../types';

// Only settings with the same meaning in both apps cross this boundary.
export function buildHachidoriAnkiHints(config: AnkiConnectConfig) {
  return {
    tags: config.tags,
    fields: {
      expression: config.fields?.word?.trim(),
      audio: (config.fields?.wordAudio || config.fields?.audio)?.trim(),
      sentence: config.fields?.sentence?.trim(),
      screenshot: config.fields?.image?.trim(),
    },
    model: config.isLapis?.enabled ? config.isLapis.sentenceCardModel?.trim() : undefined,
    family: config.isLapis?.enabled
      ? 'lapis'
      : config.isKiku?.enabled
        ? 'kiku'
        : config.isSenren?.enabled
          ? 'senren'
          : undefined,
  };
}

// Installed inside the hidden Hachidori settings page. Uses local option writes
// even when dictionaries are remote or their host is offline.
export const HACHIDORI_ANKI_SETTINGS_SCRIPT = String.raw`
    globalThis.__subminerSyncAnkiSettings = async ({ server, deck, forceOverride, hints }) => {
      const previousProxy = await globalThis.__subminerSetAnkiProxyUrl(forceOverride ? server : null);
      const { applyAnkiPreset, resolveAnkiTemplates } = await import('./anki-templates.js');
      const { ankiSetupFamily } = await import('./anki-setup.js');
      const { createAnkiGateway } = await import('./anki.js');
      const gateway = createAnkiGateway({ timeoutMs: 2000 });
      const discoveries = new Map();
      const discover = async (anki, model) => {
        const key = JSON.stringify([anki.url, anki.apiKey, model]);
        if (!discoveries.has(key)) discoveries.set(key, gateway.discover({ ...anki, model }));
        return discoveries.get(key);
      };
      for (let attempt = 0; attempt < 2; attempt += 1) {
        const options = await readOptions();
        let anki = options.anki;
        const canReplaceServer = forceOverride || !anki.url || anki.url === server
          || anki.url === 'http://127.0.0.1:8765' || anki.url === previousProxy;
        if (!canReplaceServer) return { updated: false, matched: false, reason: 'blocked-existing-server' };
        anki = { ...anki, url: server };
        const first = anki.templates[0];
        if (!first) return { updated: false, matched: false, reason: 'no-templates' };
        let template = { ...first, fields: { ...first.fields } };
        // SubMiner's new-card polling only watches its configured deck, so the
        // first template follows it like Yomitan's term card deck.
        if (deck) template.deck = deck;
        if (hints?.tags && JSON.stringify(first.tags) === JSON.stringify(['hachidori'])) {
          template.tags = [...hints.tags];
        }
        let pending = false;
        // Advanced templates include intentionally blank fields. Preserve them
        // as a unit rather than replacing them with inferred mappings.
        if (hints && first.fieldTemplates === null) {
          let discovery = await discover(anki, template.model || hints.model || '');
          if (!discovery.connected) {
            pending = true;
          } else {
            if (!template.model) {
              const candidates = hints.model
                ? discovery.models.filter(model => model === hints.model)
                : discovery.models.filter(model => !hints.family || ankiSetupFamily(model) === hints.family);
              const anchors = [hints.fields.expression, hints.fields.sentence].filter(Boolean);
              const matches = [];
              // Require a configured word and sentence field for an inferred
              // model. Never choose the first of several compatible note types.
              if (anchors.length === 2) for (const model of candidates) {
                const result = await discover(anki, model);
                if (!result.connected || result.errors.length) { pending = true; break; }
                if (anchors.every(field => result.fields.includes(field))) matches.push(result);
              }
              if (!pending && matches.length === 1) {
                discovery = matches[0];
                template.model = discovery.model;
              }
            }
            if (template.model && discovery.model === template.model && discovery.fields.length) {
              const fields = discovery.fields;
              for (const [semantic, field] of Object.entries(hints.fields)) {
                if (!template.fields[semantic] && field && fields.includes(field)) template.fields[semantic] = field;
              }
              if (Object.values(first.fields).every(value => !value)) {
                const preset = applyAnkiPreset(template, fields, ankiSetupFamily(template.model) || 'automatic');
                const configured = resolveAnkiTemplates(template, fields).templates;
                const markers = new Set(Object.entries(template.fields)
                  .filter(([, value]) => value).map(([key]) => '{' + key + '}'));
                for (const row of Object.values(preset.fieldTemplates)) {
                  for (const marker of markers) row.value = row.value.replaceAll(marker, '');
                }
                for (const [field, row] of Object.entries(configured)) {
                  if (row.value) preset.fieldTemplates[field] = row;
                }
                template = preset;
              }
            }
          }
        }
        // Hachidori retains a compatibility projection of its first template.
        // Updating both prevents its normalizer from restoring stale values.
        const templateConfig = Object.fromEntries(globalThis.HDReaderOptions.ANKI_TEMPLATE_CONFIG_KEYS
          .map(key => [key, template[key]]));
        anki = { ...anki, ...templateConfig,
          templates: [template, ...anki.templates.slice(1)] };
        const changed = JSON.stringify(anki) !== JSON.stringify(options.anki);
        try {
          if (changed) await send('hd_options_write', {
            baseRevision: options.revision, options: { anki },
          }, 'hoshidicts-worker');
          return { updated: changed, matched: !pending, pending };
        } catch (error) {
          if (attempt > 0) {
            // The server was never written, so the marker must not claim it.
            await globalThis.__subminerSetAnkiProxyUrl(previousProxy);
            throw error;
          }
          // Re-read after a concurrent settings save before filling anything.
        }
      }
    };
`;
