// Sends the current subtitle line, with a few lines of context, to `codex exec` and shows the
// answer in the SubMiner plugin panel. Registers two actions: `explain` and `translate`.
const { mkdtemp, readFile, rm } = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');

const DEFAULT_SETTINGS = {
  codexPath: 'codex',
  model: '',
  // Passed as -c model_reasoning_effort; low keeps answers to a few seconds. Empty uses your Codex config.
  reasoningEffort: 'low',
  before: 2,
  after: 1,
  timeoutMs: 120_000,
  pause: true,
  prompts: {},
};

const ACTIONS = {
  explain: {
    title: 'Grammar',
    prompt:
      'You are a Japanese tutor. Explain the line marked >>> for an intermediate learner. ' +
      'Use the other lines only as context. Reply in English markdown: a natural translation, ' +
      'then a bullet list of the grammar points and notable vocabulary. Be concise.',
  },
  translate: {
    title: 'Translation',
    prompt:
      'Translate the line marked >>> into natural English. Use the other lines only as context. ' +
      'Reply with the translation, then one short note only if the line is ambiguous.',
  },
};

function activate(ctx) {
  const settings = { ...DEFAULT_SETTINGS, ...ctx.settings };
  for (const [name, action] of Object.entries(ACTIONS)) {
    const prompt = settings.prompts?.[name] ?? action.prompt;
    ctx.actions.register(name, () => ask(ctx, settings, action.title, prompt));
  }
}

function quote(text) {
  return text
    .split('\n')
    .map((line) => `> ${line}`)
    .join('\n');
}

// Subtitles wrap across lines; flatten each cue so the >>> marker covers all of it.
function oneLine(text) {
  return text
    .split('\n')
    .map((part) => part.trim())
    .filter(Boolean)
    .join(' ');
}

function tail(text, lineCount) {
  return text.trim().split('\n').slice(-lineCount).join('\n');
}

async function ask(ctx, settings, title, instructions) {
  const { before, current, after } = ctx.subtitles.context({
    before: settings.before,
    after: settings.after,
  });
  if (!current) {
    ctx.notify('No subtitle line to send.', { variant: 'warning' });
    return;
  }
  if (settings.pause) ctx.mpv.command(['set', 'pause', 'yes']);

  const header = quote(current.text);
  const panel = ctx.ui.panel({ title, markdown: `${header}\n\n_Asking Codex..._` });
  const prompt = [
    instructions,
    '',
    'Subtitles:',
    ...before.map((line) => oneLine(line.text)),
    `>>> ${oneLine(current.text)}`,
    ...after.map((line) => oneLine(line.text)),
  ].join('\n');

  // Codex runs in an empty read-only directory so any command it tries cannot touch real files.
  const workDir = await mkdtemp(path.join(os.tmpdir(), 'subminer-ai-explain-'));
  const answerPath = path.join(workDir, 'answer.md');
  const args = [
    'exec',
    '--ephemeral',
    '--skip-git-repo-check',
    '--sandbox',
    'read-only',
    '--color',
    'never',
    '-C',
    workDir,
    '-o',
    answerPath,
  ];
  if (settings.model) args.push('-m', settings.model);
  if (settings.reasoningEffort) {
    args.push('-c', `model_reasoning_effort=${JSON.stringify(settings.reasoningEffort)}`);
  }
  args.push('-');

  try {
    const result = await ctx.exec(settings.codexPath, args, {
      stdin: prompt,
      timeoutMs: settings.timeoutMs,
      signal: panel.signal,
    });
    if (result.exitCode !== 0) {
      throw new Error(tail(result.stderr, 6) || `codex exited with code ${result.exitCode}`);
    }
    const answer = (await readFile(answerPath, 'utf8')).trim();
    panel.update({ markdown: `${header}\n\n${answer || '_Codex returned an empty answer._'}` });
  } catch (error) {
    // Closing the panel aborts the request; there is nobody left to show the error to.
    if (panel.signal.aborted) return;
    const message =
      error?.code === 'ENOENT'
        ? `Cannot find \`${settings.codexPath}\`. Set \`plugins.ai-explain.settings.codexPath\` to its full path.`
        : String(error?.message ?? error);
    ctx.log.warn(`Codex request failed: ${message}`);
    panel.update({
      markdown: `${header}\n\n**Codex request failed.**\n\n\`\`\`\n${message}\n\`\`\``,
    });
  } finally {
    await rm(workDir, { recursive: true, force: true });
  }
}

module.exports = { activate };
