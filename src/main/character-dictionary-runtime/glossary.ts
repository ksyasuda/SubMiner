import type { AnilistCharacterDictionaryCollapsibleSectionKey } from '../../types';
import { formatCharacterStats, parseCharacterDescription } from './description';
import type {
  CharacterDictionaryGlossaryEntry,
  CharacterDictionaryRole,
  CharacterRecord,
  VoiceActorRecord,
} from './types';

function roleLabel(role: CharacterDictionaryRole): string {
  if (role === 'main') return 'Protagonist';
  if (role === 'primary') return 'Main Character';
  if (role === 'side') return 'Side Character';
  return 'Minor Role';
}

function roleBadgeStyle(role: CharacterDictionaryRole): Record<string, string> {
  const base = {
    borderRadius: '4px',
    padding: '0.15em 0.5em',
    fontSize: '0.8em',
    fontWeight: 'bold',
    color: '#fff',
  };
  if (role === 'main') return { ...base, backgroundColor: '#4CAF50' };
  if (role === 'primary') return { ...base, backgroundColor: '#2196F3' };
  if (role === 'side') return { ...base, backgroundColor: '#FF9800' };
  return { ...base, backgroundColor: '#9E9E9E' };
}

function buildCollapsibleSection(
  title: string,
  open: boolean,
  body: Array<string | Record<string, unknown>> | string | Record<string, unknown>,
): Record<string, unknown> {
  return {
    tag: 'details',
    open,
    style: { marginTop: '0.4em' },
    content: [
      {
        tag: 'summary',
        style: { fontWeight: 'bold', fontSize: '0.95em', cursor: 'pointer' },
        content: title,
      },
      {
        tag: 'div',
        style: { padding: '0.25em 0 0 0.4em', fontSize: '0.9em' },
        content: body,
      },
    ],
  };
}

function buildVoicedByContent(
  voiceActors: VoiceActorRecord[],
  vaImagePaths: Map<number, string>,
): Record<string, unknown> {
  if (voiceActors.length === 1) {
    const va = voiceActors[0]!;
    const vaImgPath = vaImagePaths.get(va.id);
    const vaLabel = va.nativeName
      ? va.fullName
        ? `${va.nativeName} (${va.fullName})`
        : va.nativeName
      : va.fullName;

    if (vaImgPath) {
      return {
        tag: 'table',
        content: {
          tag: 'tr',
          content: [
            {
              tag: 'td',
              style: {
                verticalAlign: 'top',
                padding: '0',
                paddingRight: '0.4em',
                borderWidth: '0',
              },
              content: {
                tag: 'img',
                path: vaImgPath,
                width: 3,
                height: 3,
                sizeUnits: 'em',
                title: vaLabel,
                alt: vaLabel,
                collapsed: false,
                collapsible: false,
                background: true,
              },
            },
            {
              tag: 'td',
              style: { verticalAlign: 'middle', padding: '0', borderWidth: '0' },
              content: vaLabel,
            },
          ],
        },
      };
    }

    return { tag: 'div', content: vaLabel };
  }

  const items: Array<Record<string, unknown>> = [];
  for (const va of voiceActors) {
    const vaLabel = va.nativeName
      ? va.fullName
        ? `${va.nativeName} (${va.fullName})`
        : va.nativeName
      : va.fullName;
    items.push({ tag: 'li', content: vaLabel });
  }
  return { tag: 'ul', style: { marginTop: '0.15em' }, content: items };
}

function buildKnownNamesBlock(nameTerms: string[]): Record<string, unknown> | null {
  const visibleTerms = [...new Set(nameTerms.map((term) => term.trim()).filter(Boolean))];
  if (visibleTerms.length <= 1) {
    return null;
  }

  return {
    tag: 'div',
    style: { fontSize: '0.85em', marginBottom: '0.25em' },
    content: [
      {
        tag: 'div',
        style: { fontWeight: 'bold', color: '#d0d0d0', marginBottom: '0.1em' },
        content: 'Known names',
      },
      {
        tag: 'ul',
        style: { marginTop: '0', marginBottom: '0', paddingLeft: '1.2em' },
        content: visibleTerms.map((term) => ({
          tag: 'li',
          content: term,
        })),
      },
    ],
  };
}

export function createDefinitionGlossary(
  character: CharacterRecord,
  mediaTitle: string,
  imagePath: string | null,
  vaImagePaths: Map<number, string>,
  nameTerms: string[],
  getCollapsibleSectionOpenState: (
    section: AnilistCharacterDictionaryCollapsibleSectionKey,
  ) => boolean,
): CharacterDictionaryGlossaryEntry[] {
  const displayName = character.nativeName || character.fullName || `Character ${character.id}`;
  const { fields, text: descriptionText } = parseCharacterDescription(character.description);

  const content: Array<string | Record<string, unknown>> = [
    {
      tag: 'div',
      style: { fontWeight: 'bold', fontSize: '1.1em', marginBottom: '0.1em' },
      content: displayName,
    },
  ];

  const knownNamesBlock = buildKnownNamesBlock(nameTerms);
  if (knownNamesBlock) {
    content.push(knownNamesBlock);
  }

  if (imagePath) {
    content.push({
      tag: 'div',
      style: { marginTop: '0.3em', marginBottom: '0.3em' },
      content: {
        tag: 'img',
        path: imagePath,
        width: 8,
        height: 11,
        sizeUnits: 'em',
        title: displayName,
        alt: displayName,
        description: `${displayName} · ${mediaTitle}`,
        collapsed: false,
        collapsible: false,
        background: true,
      },
    });
  }

  content.push({
    tag: 'div',
    style: { fontSize: '0.8em', color: '#999', marginBottom: '0.2em' },
    content: `From: ${mediaTitle}`,
  });

  content.push({
    tag: 'div',
    style: { marginBottom: '0.15em' },
    content: {
      tag: 'span',
      style: roleBadgeStyle(character.role),
      content: roleLabel(character.role),
    },
  });

  const statsLine = formatCharacterStats(character);
  if (descriptionText) {
    content.push(
      buildCollapsibleSection(
        'Description',
        getCollapsibleSectionOpenState('description'),
        descriptionText,
      ),
    );
  }

  const fieldItems: Array<Record<string, unknown>> = [];
  if (statsLine) {
    fieldItems.push({
      tag: 'li',
      style: { fontWeight: 'bold' },
      content: statsLine,
    });
  }
  fieldItems.push(
    ...fields.map((field) => ({
      tag: 'li',
      content: `${field.key}: ${field.value}`,
    })),
  );
  if (fieldItems.length > 0) {
    content.push(
      buildCollapsibleSection(
        'Character Information',
        getCollapsibleSectionOpenState('characterInformation'),
        {
          tag: 'ul',
          style: { marginTop: '0.15em' },
          content: fieldItems,
        },
      ),
    );
  }

  if (character.voiceActors.length > 0) {
    content.push(
      buildCollapsibleSection(
        'Voiced by',
        getCollapsibleSectionOpenState('voicedBy'),
        buildVoicedByContent(character.voiceActors, vaImagePaths),
      ),
    );
  }

  return [
    {
      type: 'structured-content',
      content: { tag: 'div', content },
    },
  ];
}
