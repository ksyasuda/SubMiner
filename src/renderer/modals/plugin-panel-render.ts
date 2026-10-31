export type PluginPanelInline =
  | { kind: 'text'; value: string }
  | { kind: 'code'; value: string }
  | { kind: 'strong'; value: string }
  | { kind: 'em'; value: string };

export type PluginPanelBlock =
  | { kind: 'heading'; level: number; inline: PluginPanelInline[] }
  | { kind: 'paragraph'; lines: PluginPanelInline[][] }
  | { kind: 'list'; ordered: boolean; items: PluginPanelInline[][] }
  | { kind: 'quote'; lines: PluginPanelInline[][] }
  | { kind: 'code'; value: string }
  | { kind: 'rule' };

// Underscore emphasis needs non-word characters around it, so snake_case names stay literal.
const INLINE_PATTERN =
  /`([^`]+)`|\*\*([^*]+)\*\*|(?<!\w)__([^_]+)__(?!\w)|\*([^*\s][^*]*)\*|(?<!\w)_([^_\s][^_]*)_(?!\w)/g;

export function parseInlineMarkdown(text: string): PluginPanelInline[] {
  const tokens: PluginPanelInline[] = [];
  let lastIndex = 0;
  for (const match of text.matchAll(INLINE_PATTERN)) {
    const index = match.index ?? 0;
    if (index > lastIndex) tokens.push({ kind: 'text', value: text.slice(lastIndex, index) });
    if (match[1] !== undefined) tokens.push({ kind: 'code', value: match[1] });
    else if (match[2] !== undefined || match[3] !== undefined) {
      tokens.push({ kind: 'strong', value: (match[2] ?? match[3])! });
    } else tokens.push({ kind: 'em', value: (match[4] ?? match[5])! });
    lastIndex = index + match[0].length;
  }
  if (lastIndex < text.length) tokens.push({ kind: 'text', value: text.slice(lastIndex) });
  return tokens;
}

const HEADING = /^(#{1,6})\s+(.*)$/;
const UNORDERED_ITEM = /^\s*[-*+]\s+(.*)$/;
const ORDERED_ITEM = /^\s*\d+[.)]\s+(.*)$/;
const QUOTE = /^>\s?(.*)$/;
const RULE = /^\s*([-*_])(\s*\1){2,}\s*$/;
const FENCE = /^\s*```/;

/**
 * Parses the markdown subset model replies lean on: headings, lists, quotes, fenced code, rules,
 * and paragraphs with code/bold/italic spans. Single newlines inside a paragraph are kept as
 * line breaks. Anything else, raw HTML included, stays literal text.
 */
export function parsePluginPanelMarkdown(markdown: string): PluginPanelBlock[] {
  const blocks: PluginPanelBlock[] = [];
  const lines = markdown.replace(/\r\n?/g, '\n').split('\n');
  let index = 0;

  while (index < lines.length) {
    const line = lines[index]!;
    if (!line.trim()) {
      index += 1;
      continue;
    }

    if (FENCE.test(line)) {
      const body: string[] = [];
      index += 1;
      while (index < lines.length && !FENCE.test(lines[index]!)) {
        body.push(lines[index]!);
        index += 1;
      }
      index += 1;
      blocks.push({ kind: 'code', value: body.join('\n') });
      continue;
    }

    const heading = HEADING.exec(line);
    if (heading) {
      blocks.push({
        kind: 'heading',
        level: heading[1]!.length,
        inline: parseInlineMarkdown(heading[2]!.trim()),
      });
      index += 1;
      continue;
    }

    if (RULE.test(line)) {
      blocks.push({ kind: 'rule' });
      index += 1;
      continue;
    }

    const ordered = ORDERED_ITEM.test(line);
    if (ordered || UNORDERED_ITEM.test(line)) {
      const pattern = ordered ? ORDERED_ITEM : UNORDERED_ITEM;
      const items: PluginPanelInline[][] = [];
      while (index < lines.length) {
        const item = pattern.exec(lines[index]!);
        if (!item) break;
        items.push(parseInlineMarkdown(item[1]!.trim()));
        index += 1;
      }
      blocks.push({ kind: 'list', ordered, items });
      continue;
    }

    if (QUOTE.test(line)) {
      const quoted: PluginPanelInline[][] = [];
      while (index < lines.length) {
        const quote = QUOTE.exec(lines[index]!);
        if (!quote) break;
        quoted.push(parseInlineMarkdown(quote[1]!));
        index += 1;
      }
      blocks.push({ kind: 'quote', lines: quoted });
      continue;
    }

    const paragraph: PluginPanelInline[][] = [];
    while (index < lines.length) {
      const next = lines[index]!;
      if (
        !next.trim() ||
        FENCE.test(next) ||
        HEADING.test(next) ||
        RULE.test(next) ||
        UNORDERED_ITEM.test(next) ||
        ORDERED_ITEM.test(next) ||
        QUOTE.test(next)
      ) {
        break;
      }
      paragraph.push(parseInlineMarkdown(next.trim()));
      index += 1;
    }
    blocks.push({ kind: 'paragraph', lines: paragraph });
  }

  return blocks;
}

function appendInline(parent: HTMLElement, tokens: PluginPanelInline[]): void {
  for (const token of tokens) {
    if (token.kind === 'text') {
      parent.append(document.createTextNode(token.value));
      continue;
    }
    const tag = token.kind === 'code' ? 'code' : token.kind;
    const element = document.createElement(tag);
    element.textContent = token.value;
    parent.append(element);
  }
}

function appendLines(parent: HTMLElement, lines: PluginPanelInline[][]): void {
  lines.forEach((line, index) => {
    if (index > 0) parent.append(document.createElement('br'));
    appendInline(parent, line);
  });
}

/** Builds DOM nodes from parsed blocks. Only text nodes carry content, so nothing is interpreted. */
export function renderPluginPanelMarkdown(container: HTMLElement, markdown: string): void {
  container.replaceChildren();
  for (const block of parsePluginPanelMarkdown(markdown)) {
    switch (block.kind) {
      case 'heading': {
        // The panel title is the top heading, so content headings start one step down.
        const element = document.createElement(`h${Math.min(6, block.level + 2)}`);
        appendInline(element, block.inline);
        container.append(element);
        break;
      }
      case 'paragraph': {
        const element = document.createElement('p');
        appendLines(element, block.lines);
        container.append(element);
        break;
      }
      case 'list': {
        const list = document.createElement(block.ordered ? 'ol' : 'ul');
        for (const item of block.items) {
          const element = document.createElement('li');
          appendInline(element, item);
          list.append(element);
        }
        container.append(list);
        break;
      }
      case 'quote': {
        const element = document.createElement('blockquote');
        appendLines(element, block.lines);
        container.append(element);
        break;
      }
      case 'code': {
        const pre = document.createElement('pre');
        const code = document.createElement('code');
        code.textContent = block.value;
        pre.append(code);
        container.append(pre);
        break;
      }
      case 'rule':
        container.append(document.createElement('hr'));
        break;
    }
  }
}
