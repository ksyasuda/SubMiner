import DefaultTheme from 'vitepress/theme';
import { useRoute } from 'vitepress';
import { nextTick, onMounted, watch } from 'vue';
import '@catppuccin/vitepress/theme/macchiato/mauve.css';
import './mermaid-modal.css';

let mermaidLoader: Promise<any> | null = null;
const MERMAID_MODAL_ID = 'mermaid-diagram-modal';

function closeMermaidModal() {
  if (typeof document === 'undefined') {
    return;
  }

  const modal = document.getElementById(MERMAID_MODAL_ID);
  if (!modal) {
    return;
  }

  modal.classList.remove('is-open');
  document.body.classList.remove('mermaid-modal-open');
}

function ensureMermaidModal(): HTMLDivElement {
  const existing = document.getElementById(MERMAID_MODAL_ID);
  if (existing) {
    return existing as HTMLDivElement;
  }

  const modal = document.createElement('div');
  modal.id = MERMAID_MODAL_ID;
  modal.className = 'mermaid-modal';
  modal.innerHTML = `
    <div class="mermaid-modal__backdrop" data-mermaid-close="true"></div>
    <div class="mermaid-modal__dialog" role="dialog" aria-modal="true" aria-label="Expanded Mermaid diagram">
      <button class="mermaid-modal__close" type="button" aria-label="Close Mermaid diagram">Close</button>
      <div class="mermaid-modal__content"></div>
    </div>
  `;

  modal.addEventListener('click', (event) => {
    const target = event.target as HTMLElement | null;
    if (!target) {
      return;
    }

    if (target.closest('[data-mermaid-close="true"]') || target.closest('.mermaid-modal__close')) {
      closeMermaidModal();
    }
  });

  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && modal.classList.contains('is-open')) {
      closeMermaidModal();
    }
  });

  document.body.appendChild(modal);
  return modal;
}

function openMermaidModal(sourceNode: HTMLElement) {
  if (typeof document === 'undefined') {
    return;
  }

  const modal = ensureMermaidModal();
  const content = modal.querySelector<HTMLDivElement>('.mermaid-modal__content');
  if (!content) {
    return;
  }

  content.replaceChildren(sourceNode.cloneNode(true));
  modal.classList.add('is-open');
  document.body.classList.add('mermaid-modal-open');
}

function attachMermaidInteractions(nodes: HTMLElement[]) {
  for (const node of nodes) {
    if (node.dataset.mermaidInteractive === 'true') {
      continue;
    }

    const svg = node.querySelector<HTMLElement>('svg');
    if (!svg) {
      continue;
    }

    node.classList.add('mermaid-interactive');
    node.setAttribute('role', 'button');
    node.setAttribute('tabindex', '0');
    node.setAttribute('aria-label', 'Open Mermaid diagram in full view');

    const open = () => openMermaidModal(svg);
    node.addEventListener('click', open);
    node.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        open();
      }
    });

    node.dataset.mermaidInteractive = 'true';
  }
}

async function getMermaid() {
  if (!mermaidLoader) {
    mermaidLoader = import('mermaid').then((module) => {
      const mermaid = module.default;
      mermaid.initialize({
        startOnLoad: false,
        securityLevel: 'loose',
        theme: 'base',
        themeVariables: {
          background: '#24273a',
          primaryColor: '#363a4f',
          primaryTextColor: '#cad3f5',
          primaryBorderColor: '#c6a0f6',
          secondaryColor: '#494d64',
          secondaryTextColor: '#cad3f5',
          secondaryBorderColor: '#b7bdf8',
          tertiaryColor: '#5b6078',
          tertiaryTextColor: '#cad3f5',
          tertiaryBorderColor: '#8aadf4',
          lineColor: '#939ab7',
          textColor: '#cad3f5',
          mainBkg: '#363a4f',
          nodeBorder: '#c6a0f6',
          clusterBkg: '#1e2030',
          clusterBorder: '#494d64',
          edgeLabelBackground: '#24273a',
          labelTextColor: '#cad3f5',
        },
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
  const blocks = Array.from(document.querySelectorAll<HTMLElement>('div.language-mermaid'));
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
    attachMermaidInteractions(nodes);
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
