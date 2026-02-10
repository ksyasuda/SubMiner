import DefaultTheme from 'vitepress/theme';
import { useRoute } from 'vitepress';
import { nextTick, onMounted, watch } from 'vue';
import '@catppuccin/vitepress/theme/macchiato/mauve.css';

let mermaidLoader: Promise<any> | null = null;

async function getMermaid() {
  if (!mermaidLoader) {
    mermaidLoader = import(
      /* @vite-ignore */ 'https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.esm.min.mjs'
    ).then((mod) => {
      const mermaid = mod.default ?? mod;
      mermaid.initialize({
        startOnLoad: false,
        securityLevel: 'loose',
      });
      return mermaid;
    });
  }
  return mermaidLoader;
}

async function renderMermaidBlocks() {
  if (typeof document === 'undefined') {
    return;
  }
  const blocks = Array.from(
    document.querySelectorAll<HTMLElement>('div.language-mermaid'),
  );
  if (blocks.length === 0) {
    return;
  }

  const mermaid = await getMermaid();
  const nodes: HTMLElement[] = [];

  for (const block of blocks) {
    if (block.dataset.mermaidRendered === 'true') {
      continue;
    }
    const code = block.querySelector('pre code');
    const source = code?.textContent?.trim();
    if (!source) {
      continue;
    }

    const mount = document.createElement('div');
    mount.className = 'mermaid';
    mount.textContent = source;

    block.replaceChildren(mount);
    block.dataset.mermaidRendered = 'true';
    nodes.push(mount);
  }

  if (nodes.length > 0) {
    await mermaid.run({ nodes });
  }
}

export default {
  ...DefaultTheme,
  setup() {
    const route = useRoute();
    const render = () => {
      nextTick(() => {
        renderMermaidBlocks().catch((error) => {
          console.error('Failed to render Mermaid diagram:', error);
        });
      });
    };

    onMounted(render);
    watch(() => route.path, render);
  },
};
