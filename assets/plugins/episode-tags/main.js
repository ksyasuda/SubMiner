// Tags each mined card with the show and episode it came from, e.g. `show::Frieren::S01E07`.
// The show and episode are parsed from the file name, or from mpv's media title for streams.
const path = require('node:path');

const DEFAULT_SETTINGS = {
  // Placeholders: {show}, {season}, {episode}, and {ep} (S01E07, or E07 without a season).
  // A tag stops before the first `::` level whose placeholder has no value.
  tags: ['show::{show}::{ep}'],
  notify: false,
};

// Tried in order; the first match wins. Runs on names with brackets and dots already removed.
const EPISODE_PATTERNS = [
  /^(?<show>.*?)\s*S(?<season>\d{1,2})\s?E(?<episode>\d{1,4})\b/i,
  /^(?<show>.*?)\s*(?:Season\s?(?<season>\d{1,2})\s*)?(?:Episode|Ep)\s?(?<episode>\d{1,4})\b/i,
  /^(?<show>.*?)\s+-\s+(?<episode>\d{1,4})(?:v\d)?\b/,
  /^(?<show>.*?)\s+(?<episode>\d{1,4})(?:v\d)?$/,
  /^(?<show>.*?)\s*第(?<episode>\d{1,4})話/,
];

// `[Group] Show Name - 07 (1080p).mkv` -> { show: 'Show Name', season: null, episode: '07' }
function parseMediaName(name) {
  const cleaned = name
    .replace(/\.[a-z0-9]{2,4}$/i, '')
    .replace(/\[[^\]]*\]|\([^)]*\)|【[^】]*】/g, ' ')
    .replace(/[._]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  for (const pattern of EPISODE_PATTERNS) {
    const groups = cleaned.match(pattern)?.groups;
    const show = groups?.show.replace(/[\s-]+$/, '');
    if (!show) continue;
    return { show, season: pad(groups.season), episode: pad(groups.episode) };
  }
  return { show: cleaned || null, season: null, episode: null };
}

function pad(number) {
  return number ? String(Number(number)).padStart(2, '0') : null;
}

// Anki tags can't hold spaces, and a `::` in a value would add a hierarchy level.
function tagValue(value) {
  return value.replace(/::?/g, ' ').trim().replace(/\s+/g, '_').replace(/"/g, '');
}

function renderTag(template, values) {
  const levels = [];
  for (const level of template.split('::')) {
    let missing = false;
    const rendered = level.replace(/\{(\w+)\}/g, (_, key) => {
      const value = values[key];
      if (!value) missing = true;
      return value ? tagValue(value) : '';
    });
    if (missing) break;
    levels.push(rendered);
  }
  return levels.join('::');
}

function mediaName(mediaPath, title) {
  // A stream URL's last segment is usually an id, so streams use the title mpv shows.
  if (typeof mediaPath === 'string' && mediaPath && !/^[a-z]+:\/\//i.test(mediaPath)) {
    return path.basename(mediaPath);
  }
  return typeof title === 'string' && title ? title : null;
}

function renderTags(templates, name) {
  const { show, season, episode } = parseMediaName(name);
  const ep = episode ? `${season ? `S${season}` : ''}E${episode}` : null;
  const values = { show, season, episode, ep };
  return [...new Set(templates.map((template) => renderTag(template, values)))].filter(Boolean);
}

function activate(ctx) {
  const settings = { ...DEFAULT_SETTINGS, ...ctx.settings };
  ctx.anki.on('card-mined', async ({ notes }) => {
    // Notes from the same file share one addTags call.
    const notesByName = new Map();
    for (const { noteId, provenance } of notes) {
      const name = mediaName(provenance?.media?.path, provenance?.media?.title);
      if (!name) {
        ctx.log.debug(`Note ${noteId} has no known media source; not tagging it.`);
        continue;
      }
      if (!notesByName.has(name)) notesByName.set(name, new Set());
      notesByName.get(name).add(noteId);
    }

    for (const [name, notes] of notesByName) {
      const tags = renderTags(settings.tags, name);
      if (tags.length === 0) {
        ctx.log.debug(`No tags for "${name}".`);
        continue;
      }
      try {
        await ctx.anki.invoke('addTags', { notes: [...notes], tags: tags.join(' ') });
        ctx.log.info(`Tagged ${notes.size} note(s): ${tags.join(' ')}`);
        if (settings.notify) ctx.notify(`Tagged ${tags.join(' ')}`);
      } catch (error) {
        ctx.log.warn(`Tagging failed: ${error.message}`);
        ctx.notify(`Episode tagging failed: ${error.message}`, { variant: 'error' });
      }
    }
  });
}

module.exports = { activate };
