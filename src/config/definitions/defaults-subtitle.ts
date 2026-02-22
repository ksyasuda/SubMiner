import { ResolvedConfig } from '../../types';

export const SUBTITLE_DEFAULT_CONFIG: Pick<ResolvedConfig, 'subtitleStyle'> = {
  subtitleStyle: {
    enableJlpt: false,
    preserveLineBreaks: false,
    hoverTokenColor: '#c6a0f6',
    fontFamily:
      'M PLUS 1, Noto Sans CJK JP Regular, Noto Sans CJK JP, Hiragino Sans, Hiragino Kaku Gothic ProN, Yu Gothic, Arial Unicode MS, Arial, sans-serif',
    fontSize: 35,
    fontColor: '#cad3f5',
    fontWeight: 'normal',
    fontStyle: 'normal',
    backgroundColor: 'rgb(30, 32, 48, 0.88)',
    nPlusOneColor: '#c6a0f6',
    knownWordColor: '#a6da95',
    jlptColors: {
      N1: '#ed8796',
      N2: '#f5a97f',
      N3: '#f9e2af',
      N4: '#a6e3a1',
      N5: '#8aadf4',
    },
    frequencyDictionary: {
      enabled: false,
      sourcePath: '',
      topX: 1000,
      mode: 'single',
      singleColor: '#f5a97f',
      bandedColors: ['#ed8796', '#f5a97f', '#f9e2af', '#a6e3a1', '#8aadf4'],
    },
    secondary: {
      fontSize: 24,
      fontColor: '#ffffff',
      backgroundColor: 'transparent',
      fontWeight: 'normal',
      fontStyle: 'normal',
      fontFamily:
        'M PLUS 1, Noto Sans CJK JP Regular, Noto Sans CJK JP, Hiragino Sans, Hiragino Kaku Gothic ProN, Yu Gothic, Arial Unicode MS, Arial, sans-serif',
    },
  },
};
