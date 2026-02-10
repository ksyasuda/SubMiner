# Overlay Positioning Flow

## 1) Window Bounds Flow (visible + invisible Electron windows)

```mermaid
flowchart LR
  A[Platform backend selection<br/>src/window-trackers/index.ts] --> B[Tracker emits geometry<br/>onWindowFound/onGeometryChange]
  B --> C[updateOverlayBounds<br/>src/main.ts]
  C --> D[mainWindow.setBounds]
  C --> E[invisibleWindow.setBounds]
```

## 2) Invisible Subtitle Layout Flow (mpv render metrics -> DOM layout)

```mermaid
flowchart LR
  A[mpv property changes<br/>sub-pos, sub-font-size, osd-dimensions, etc.] --> B[MpvIpcClient parses events<br/>src/main.ts]
  B --> C[updateMpvSubtitleRenderMetrics<br/>src/main.ts]
  C --> D[broadcast mpv-subtitle-render-metrics:set]
  D --> E[preload onMpvSubtitleRenderMetrics<br/>src/preload.ts]
  E --> F[renderer receives metrics event<br/>src/renderer/renderer.ts]
  F --> G[applyInvisibleSubtitleLayoutFromMpvMetrics]
  G --> H[subtitleContainer/subtitleRoot inline styles updated]
```

## 3) Visible Subtitle Manual Position Flow

```mermaid
flowchart LR
  A[User right-click drags subtitle] --> B[setupDragging<br/>src/renderer/renderer.ts]
  B --> C[applyYPercent]
  C --> D[saveSubtitlePosition IPC]
  D --> E[saveSubtitlePosition in main<br/>src/main.ts]
  E --> F[load/broadcast subtitle-position:set]
  F --> G[applyStoredSubtitlePosition in renderer]
```

## 4) Fallback Bounds Flow (tracker not ready)

```mermaid
flowchart LR
  A[windowTracker exists but not tracking] --> B["screen.getDisplayNearestPoint(cursor)"]
  B --> C[display.workArea]
  C --> D[updateOverlayBounds]
```

## Key Files

- `src/main.ts`
- `src/renderer/renderer.ts`
- `src/preload.ts`
- `src/window-trackers/base-tracker.ts`
- `src/window-trackers/index.ts`
