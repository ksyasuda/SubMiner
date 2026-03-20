import * as fs from 'fs';
import * as http from 'http';
import * as path from 'path';
import { createLogger } from '../../logger';

const logger = createLogger('main:texthooker');

export type TexthookerBootstrapSettings = {
  enableKnownWordColoring: boolean;
  enableNPlusOneColoring: boolean;
  enableNameMatchColoring: boolean;
  enableFrequencyColoring: boolean;
  enableJlptColoring: boolean;
  characterDictionaryEnabled: boolean;
  knownWordColor: string;
  nPlusOneColor: string;
  nameMatchColor: string;
  hoverTokenColor: string;
  hoverTokenBackgroundColor: string;
  jlptColors: {
    N1: string;
    N2: string;
    N3: string;
    N4: string;
    N5: string;
  };
  frequencyDictionary: {
    singleColor: string;
    bandedColors: readonly [string, string, string, string, string];
  };
};

function buildTexthookerBootstrapScript(
  websocketUrl?: string,
  settings?: TexthookerBootstrapSettings,
): string {
  const statements: string[] = [];

  if (websocketUrl) {
    statements.push(
      `window.localStorage.setItem('bannou-texthooker-websocketUrl', ${JSON.stringify(websocketUrl)});`,
    );
  }

  if (settings) {
    const booleanStorageValue = (enabled: boolean): '"1"' | '"0"' => (enabled ? '"1"' : '"0"');
    statements.push(
      `window.localStorage.setItem('bannou-texthooker-enableKnownWordColoring', ${booleanStorageValue(settings.enableKnownWordColoring)});`,
      `window.localStorage.setItem('bannou-texthooker-enableNPlusOneColoring', ${booleanStorageValue(settings.enableNPlusOneColoring)});`,
      `window.localStorage.setItem('bannou-texthooker-enableNameMatchColoring', ${booleanStorageValue(settings.enableNameMatchColoring)});`,
      `window.localStorage.setItem('bannou-texthooker-enableFrequencyColoring', ${booleanStorageValue(settings.enableFrequencyColoring)});`,
      `window.localStorage.setItem('bannou-texthooker-enableJlptColoring', ${booleanStorageValue(settings.enableJlptColoring)});`,
      `window.localStorage.setItem('bannou-texthooker-characterDictionaryEnabled', ${booleanStorageValue(settings.characterDictionaryEnabled)});`,
    );
  }

  return statements.length > 0 ? `<script>${statements.join('')}</script>` : '';
}

function buildTexthookerBootstrapStyle(settings?: TexthookerBootstrapSettings): string {
  if (!settings) {
    return '';
  }

  const [band1, band2, band3, band4, band5] = settings.frequencyDictionary.bandedColors;

  return `<style id="subminer-texthooker-bootstrap-style">:root{--subminer-known-word-color:${settings.knownWordColor};--subminer-n-plus-one-color:${settings.nPlusOneColor};--subminer-name-match-color:${settings.nameMatchColor};--subminer-jlpt-n1-color:${settings.jlptColors.N1};--subminer-jlpt-n2-color:${settings.jlptColors.N2};--subminer-jlpt-n3-color:${settings.jlptColors.N3};--subminer-jlpt-n4-color:${settings.jlptColors.N4};--subminer-jlpt-n5-color:${settings.jlptColors.N5};--subminer-frequency-single-color:${settings.frequencyDictionary.singleColor};--subminer-frequency-band-1-color:${band1};--subminer-frequency-band-2-color:${band2};--subminer-frequency-band-3-color:${band3};--subminer-frequency-band-4-color:${band4};--subminer-frequency-band-5-color:${band5};--sm-token-hover-bg:${settings.hoverTokenBackgroundColor};--sm-token-hover-text:${settings.hoverTokenColor};}</style>`;
}

export function injectTexthookerBootstrapHtml(
  html: string,
  websocketUrl?: string,
  settings?: TexthookerBootstrapSettings,
): string {
  const bootstrapStyle = buildTexthookerBootstrapStyle(settings);
  const bootstrapScript = buildTexthookerBootstrapScript(websocketUrl, settings);

  if (!bootstrapStyle && !bootstrapScript) {
    return html;
  }

  if (html.includes('</head>')) {
    return html.replace('</head>', `${bootstrapStyle}${bootstrapScript}</head>`);
  }

  return `${bootstrapStyle}${bootstrapScript}${html}`;
}

export class Texthooker {
  constructor(
    private readonly getBootstrapSettings?: () => TexthookerBootstrapSettings | undefined,
  ) {}

  private server: http.Server | null = null;

  public isRunning(): boolean {
    return this.server !== null;
  }

  public start(port: number, websocketUrl?: string): http.Server | null {
    if (this.server) {
      return this.server;
    }

    const texthookerPath = this.getTexthookerPath();
    if (!texthookerPath) {
      logger.error('texthooker-ui not found');
      return null;
    }

    this.server = http.createServer((req, res) => {
      const urlPath = (req.url || '/').split('?')[0] ?? '/';
      const filePath = path.join(texthookerPath, urlPath === '/' ? 'index.html' : urlPath);

      const ext = path.extname(filePath);
      const mimeTypes: Record<string, string> = {
        '.html': 'text/html',
        '.js': 'application/javascript',
        '.css': 'text/css',
        '.json': 'application/json',
        '.png': 'image/png',
        '.svg': 'image/svg+xml',
        '.ttf': 'font/ttf',
        '.woff': 'font/woff',
        '.woff2': 'font/woff2',
      };

      fs.readFile(filePath, (err, data) => {
        if (err) {
          res.writeHead(404);
          res.end('Not found');
          return;
        }
        const bootstrapSettings = this.getBootstrapSettings?.();
        const responseData =
          urlPath === '/' || urlPath === '/index.html'
            ? Buffer.from(
                injectTexthookerBootstrapHtml(
                  data.toString('utf-8'),
                  websocketUrl,
                  bootstrapSettings,
                ),
              )
            : data;
        res.writeHead(200, { 'Content-Type': mimeTypes[ext] || 'text/plain' });
        res.end(responseData);
      });
    });

    this.server.listen(port, '127.0.0.1', () => {
      logger.info(`Texthooker server running at http://127.0.0.1:${port}`);
    });

    return this.server;
  }

  public stop(): void {
    if (this.server) {
      this.server.close();
      this.server = null;
    }
  }

  private getTexthookerPath(): string | null {
    const searchPaths = [
      path.join(__dirname, '..', '..', '..', 'vendor', 'texthooker-ui', 'docs'),
      path.join(process.resourcesPath, 'app', 'vendor', 'texthooker-ui', 'docs'),
    ];
    for (const candidate of searchPaths) {
      if (fs.existsSync(path.join(candidate, 'index.html'))) {
        return candidate;
      }
    }
    return null;
  }
}
