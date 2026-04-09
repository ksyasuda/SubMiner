export const CHART_THEME = {
  tick: '#a5adcb',
  tooltipBg: '#363a4f',
  tooltipBorder: '#494d64',
  tooltipText: '#cad3f5',
  tooltipLabel: '#b8c0e0',
  barFill: '#8aadf4',
  grid: '#494d64',
  axisLine: '#494d64',
} as const;

export const CHART_DEFAULTS = {
  height: 160,
  tickFontSize: 11,
  margin: { top: 8, right: 8, bottom: 0, left: 0 },
  grid: { strokeDasharray: '3 3', vertical: false },
} as const;

export const TOOLTIP_CONTENT_STYLE = {
  background: CHART_THEME.tooltipBg,
  border: `1px solid ${CHART_THEME.tooltipBorder}`,
  borderRadius: 6,
  color: CHART_THEME.tooltipText,
  fontSize: 12,
};
