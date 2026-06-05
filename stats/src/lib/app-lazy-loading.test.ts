import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const APP_PATH = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../App.tsx');

test('App lazy-loads non-overview tabs and detail surfaces behind Suspense boundaries', () => {
  const source = fs.readFileSync(APP_PATH, 'utf8');

  assert.match(source, /\bSuspense\b/, 'expected Suspense boundary in App');
  assert.match(source, /lazy\(\(\) =>\s*import\('\.\/components\/anime\/AnimeTab'\)/);
  assert.match(source, /lazy\(\(\) =>\s*import\('\.\/components\/trends\/TrendsTab'\)/);
  assert.match(source, /lazy\(\(\) =>\s*import\('\.\/components\/vocabulary\/VocabularyTab'\)/);
  assert.match(source, /lazy\(\(\) =>\s*import\('\.\/components\/search\/SearchTab'\)/);
  assert.match(source, /lazy\(\(\) =>\s*import\('\.\/components\/sessions\/SessionsTab'\)/);
  assert.match(source, /lazy\(\(\) =>\s*import\('\.\/components\/library\/MediaDetailView'\)/);
  assert.match(source, /lazy\(\(\) =>\s*import\('\.\/components\/vocabulary\/WordDetailPanel'\)/);

  assert.doesNotMatch(source, /import \{ AnimeTab \} from '\.\/components\/anime\/AnimeTab';/);
  assert.doesNotMatch(source, /import \{ TrendsTab \} from '\.\/components\/trends\/TrendsTab';/);
  assert.doesNotMatch(
    source,
    /import \{ VocabularyTab \} from '\.\/components\/vocabulary\/VocabularyTab';/,
  );
  assert.doesNotMatch(source, /import \{ SearchTab \} from '\.\/components\/search\/SearchTab';/);
  assert.doesNotMatch(
    source,
    /import \{ SessionsTab \} from '\.\/components\/sessions\/SessionsTab';/,
  );
  assert.doesNotMatch(
    source,
    /import \{ MediaDetailView \} from '\.\/components\/library\/MediaDetailView';/,
  );
  assert.doesNotMatch(
    source,
    /import \{ WordDetailPanel \} from '\.\/components\/vocabulary\/WordDetailPanel';/,
  );
});
