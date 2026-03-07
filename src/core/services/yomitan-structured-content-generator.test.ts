import assert from 'node:assert/strict';
import path from 'node:path';
import test from 'node:test';
import { pathToFileURL } from 'node:url';
import { resolveYomitanExtensionPath } from './yomitan-extension-paths';

class FakeStyle {
  private values = new Map<string, string>();

  set width(value: string) {
    this.values.set('width', value);
  }

  get width(): string {
    return this.values.get('width') ?? '';
  }

  set height(value: string) {
    this.values.set('height', value);
  }

  get height(): string {
    return this.values.get('height') ?? '';
  }

  set border(value: string) {
    this.values.set('border', value);
  }

  set borderRadius(value: string) {
    this.values.set('borderRadius', value);
  }

  set paddingTop(value: string) {
    this.values.set('paddingTop', value);
  }

  setProperty(name: string, value: string): void {
    this.values.set(name, value);
  }

  removeProperty(name: string): void {
    this.values.delete(name);
  }
}

class FakeNode {
  public childNodes: Array<FakeNode | FakeTextNode> = [];
  public className = '';
  public dataset: Record<string, string> = {};
  public style = new FakeStyle();
  public textContent: string | null = null;
  public title = '';
  public href = '';
  public rel = '';
  public target = '';
  public width = 0;
  public height = 0;
  public parentNode: FakeNode | null = null;

  constructor(public readonly tagName: string) {}

  appendChild(node: FakeNode | FakeTextNode): FakeNode | FakeTextNode {
    if (node instanceof FakeNode) {
      node.parentNode = this;
    }
    this.childNodes.push(node);
    return node;
  }

  addEventListener(): void {}

  closest(selector: string): FakeNode | null {
    if (!selector.startsWith('.')) {
      return null;
    }
    const className = selector.slice(1);
    let current: FakeNode | null = this;
    while (current) {
      if (current.className === className) {
        return current;
      }
      current = current.parentNode;
    }
    return null;
  }

  removeAttribute(name: string): void {
    if (name === 'src') {
      return;
    }
    if (name === 'href') {
      this.href = '';
    }
  }
}

class FakeImageElement extends FakeNode {
  public onload: (() => void) | null = null;
  public onerror: ((error: unknown) => void) | null = null;
  private _src = '';

  constructor() {
    super('img');
  }

  set src(value: string) {
    this._src = value;
    this.onload?.();
  }

  get src(): string {
    return this._src;
  }
}

class FakeCanvasElement extends FakeNode {
  constructor() {
    super('canvas');
  }
}

class FakeTextNode {
  constructor(public readonly data: string) {}
}

class FakeDocument {
  createElement(tagName: string): FakeNode {
    if (tagName === 'img') {
      return new FakeImageElement();
    }
    if (tagName === 'canvas') {
      return new FakeCanvasElement();
    }
    return new FakeNode(tagName);
  }

  createTextNode(data: string): FakeTextNode {
    return new FakeTextNode(data);
  }
}

function findFirstByClass(node: FakeNode, className: string): FakeNode | null {
  if (node.className === className) {
    return node;
  }
  for (const child of node.childNodes) {
    if (child instanceof FakeNode) {
      const result = findFirstByClass(child, className);
      if (result) {
        return result;
      }
    }
  }
  return null;
}

test('StructuredContentGenerator uses direct img loading for popup glossary images', async () => {
  const yomitanRoot = resolveYomitanExtensionPath({ cwd: process.cwd() });
  assert.ok(yomitanRoot, 'Run `bun run build:yomitan` before Yomitan integration tests.');

  const { DisplayContentManager } = await import(
    pathToFileURL(
      path.join(yomitanRoot, 'js', 'display', 'display-content-manager.js'),
    ).href
  );
  const { StructuredContentGenerator } = await import(
    pathToFileURL(
      path.join(yomitanRoot, 'js', 'display', 'structured-content-generator.js'),
    ).href
  );

  const createObjectURLCalls: string[] = [];
  const revokeObjectURLCalls: string[] = [];
  const originalHtmlImageElement = globalThis.HTMLImageElement;
  const originalHtmlCanvasElement = globalThis.HTMLCanvasElement;
  const originalCreateObjectURL = URL.createObjectURL;
  const originalRevokeObjectURL = URL.revokeObjectURL;
  globalThis.HTMLImageElement = FakeImageElement as unknown as typeof HTMLImageElement;
  globalThis.HTMLCanvasElement = FakeCanvasElement as unknown as typeof HTMLCanvasElement;
  URL.createObjectURL = (_blob: Blob) => {
    const value = 'blob:test-image';
    createObjectURLCalls.push(value);
    return value;
  };
  URL.revokeObjectURL = (value: string) => {
    revokeObjectURLCalls.push(value);
  };

  try {
    const manager = new DisplayContentManager({
      application: {
        api: {
          getMedia: async () => [
            {
              content: Buffer.from('png-bytes').toString('base64'),
              mediaType: 'image/png',
            },
          ],
        },
      },
    });

    const generator = new StructuredContentGenerator(
      manager,
      new FakeDocument(),
      {
        devicePixelRatio: 1,
        navigator: { userAgent: 'Mozilla/5.0' },
      },
    );

    const node = generator.createDefinitionImage(
      {
        tag: 'img',
        path: 'img/test.png',
        width: 8,
        height: 11,
        title: 'Alpha',
        background: true,
      },
      'SubMiner Character Dictionary',
    ) as FakeNode;

    await manager.executeMediaRequests();

    const imageNode = findFirstByClass(node, 'gloss-image');
    assert.ok(imageNode);
    assert.equal(imageNode.tagName, 'img');
    assert.equal((imageNode as FakeImageElement).src, 'blob:test-image');
    assert.equal(node.dataset.imageLoadState, 'loaded');
    assert.equal(node.dataset.hasImage, 'true');
    assert.deepEqual(createObjectURLCalls, ['blob:test-image']);

    manager.unloadAll();
    assert.deepEqual(revokeObjectURLCalls, ['blob:test-image']);
  } finally {
    globalThis.HTMLImageElement = originalHtmlImageElement;
    globalThis.HTMLCanvasElement = originalHtmlCanvasElement;
    URL.createObjectURL = originalCreateObjectURL;
    URL.revokeObjectURL = originalRevokeObjectURL;
  }
});
