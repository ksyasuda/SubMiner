---
layout: home

title: SubMiner
titleTemplate: Immersion Mining Workflow for MPV

hero:
  name: SubMiner
  text: Immersion Mining for MPV
  tagline: Watch media, mine vocabulary, and craft anki cards without leaving the scene.
  image:
    src: /assets/SubMiner.png
    alt: SubMiner logo
  actions:
    - theme: brand
      text: Install
      link: /installation
    - theme: alt
      text: Explore workflow
      link: /mining-workflow

features:
  - icon:
      src: /assets/mpv.svg
      alt: mpv icon
    title: Built for mpv
    details: Tracks subtitles via mpv IPC in real time. Launch with the wrapper script or the mpv plugin — no external bridge needed.
    link: /usage
    linkText: How it works
  - icon:
      src: /assets/yomitan-icon.svg
      alt: Yomitan logo
    title: Bundled Yomitan
    details: Ships with a built-in Yomitan instance for instant word lookups and context-aware card creation directly from subtitle text.
    link: /mining-workflow
    linkText: Mining workflow
  - icon:
      src: /assets/anki-card.svg
      alt: Anki card icon
    title: Anki Card Enrichment
    details: Auto-fills card fields with sentence, audio clip, screenshot, and translation so you can focus on learning.
    link: /anki-integration
    linkText: Anki integration
  - icon:
      src: /assets/highlight.svg
      alt: Highlight icon
    title: Reading Annotations
    details: N+1 targeting, character-name matching, frequency highlighting, and JLPT tagging — all layered on subtitle text in real time.
    link: /subtitle-annotations
    linkText: Annotation details
  - icon:
      src: /assets/video.svg
      alt: Video playback icon
    title: YouTube & Whisper
    details: Play YouTube URLs or searches with native subtitles, or generate them with whisper.cpp and optional AI cleanup.
    link: /usage#youtube-playback
    linkText: YouTube playback
  - icon:
      src: /assets/jellyfin.svg
      alt: Jellyfin icon
    title: Jellyfin Integration
    details: Browse your Jellyfin library, pick media interactively, and play through mpv with full subtitle and mining support.
    link: /jellyfin-integration
    linkText: Jellyfin setup
  - icon:
      src: /assets/subtitle-download.svg
      alt: Subtitle download icon
    title: Subtitle Download & Sync
    details: Search and pull subtitles from Jimaku, then auto-sync timing with alass or ffsubsync — all from the overlay.
    link: /configuration#jimaku
    linkText: Jimaku integration
  - icon:
      src: /assets/tokenization.svg
      alt: Tracking chart icon
    title: Immersion Tracking
    details: Logs watch time, words encountered, and cards mined to SQLite with daily and monthly rollups for long-term progress tracking.
    link: /immersion-tracking
    linkText: Tracking details
  - icon:
      src: /assets/cross-platform.svg
      alt: Cross-platform icon
    title: Cross-Platform
    details: Runs on Linux (Hyprland, Sway, X11), macOS, and Windows with compositor-aware window positioning and platform-native integration.
    link: /installation
    linkText: Platform setup
---

<script setup>
const demoAssetVersion = '20260223-2';
</script>

<div class="landing-shell">
  <section class="workflow-section">
    <h2>How it fits together</h2>
    <div class="workflow-steps">
      <div class="workflow-step" style="animation-delay: 0ms">
        <div class="step-number">01</div>
        <div class="step-title">Start</div>
        <div class="step-desc">Launch with the wrapper or existing mpv setup and keep subtitles in sync.</div>
      </div>
      <div class="workflow-connector" aria-hidden="true"></div>
      <div class="workflow-step" style="animation-delay: 60ms">
        <div class="step-number">02</div>
        <div class="step-title">Lookup</div>
        <div class="step-desc">Hover or click a token in the interactive overlay to open Yomitan context.</div>
      </div>
      <div class="workflow-connector" aria-hidden="true"></div>
      <div class="workflow-step" style="animation-delay: 120ms">
        <div class="step-number">03</div>
        <div class="step-title">Mine</div>
        <div class="step-desc">Create cards from Yomitan or mine sentence cards directly from subtitle lines.</div>
      </div>
      <div class="workflow-connector" aria-hidden="true"></div>
      <div class="workflow-step" style="animation-delay: 180ms">
        <div class="step-number">04</div>
        <div class="step-title">Enrich</div>
        <div class="step-desc">Automatically attach timing-accurate audio, sentence text, and visual evidence.</div>
      </div>
      <div class="workflow-connector" aria-hidden="true"></div>
      <div class="workflow-step" style="animation-delay: 240ms">
        <div class="step-number">05</div>
        <div class="step-title">Track</div>
        <div class="step-desc">Review immersion history and repeat high-value patterns over time.</div>
      </div>
    </div>
  </section>

  <section class="demo-section">
    <h2>See it in action</h2>
    <p>Subtitles, lookup flow, and card enrichment from a real playback session.</p>
    <div class="demo-window">
      <div class="demo-window__bar">
        <span class="demo-window__dot"></span>
        <span class="demo-window__dot"></span>
        <span class="demo-window__dot"></span>
        <span class="demo-window__title">subminer -- playback</span>
      </div>
      <video controls playsinline preload="metadata" :poster="`/assets/minecard-poster.jpg?v=${demoAssetVersion}`">
        <source :src="`/assets/minecard.webm?v=${demoAssetVersion}`" type="video/webm" />
        <source :src="`/assets/minecard.mp4?v=${demoAssetVersion}`" type="video/mp4" />
        <a :href="`/assets/minecard.webm?v=${demoAssetVersion}`" target="_blank" rel="noreferrer">
          <img :src="`/assets/minecard.webp?v=${demoAssetVersion}`" alt="SubMiner demo Animated fallback" style="width: 100%; height: auto;" />
        </a>
      </video>
    </div>
  </section>
</div>

<style>
.landing-shell {
  max-width: 1120px;
  margin: 0 auto;
  padding: 0.5rem 1rem 4rem;
}

.landing-shell,
.landing-shell .step-title,
.landing-shell h1,
.landing-shell h2 {
  font-family: var(--tui-font-mono);
}

.VPHome :deep(.VPFeature),
.VPHome :deep(.VPButton),
.landing-shell .workflow-step,
.landing-shell .demo-window,
.landing-shell .demo-window__bar {
  border-radius: 8px;
}

.step-title,
.step-number {
  font-family: var(--tui-font-mono);
  letter-spacing: -0.01em;
}

/* === Workflow === */
.workflow-section {
  margin: 2.4rem auto 0;
  padding: 0;
}

.workflow-section h2,
.demo-section h2 {
  font-size: 1.45rem;
  font-weight: 600;
  letter-spacing: -0.01em;
  margin-bottom: 1rem;
  padding-bottom: 4px;
}

.workflow-section h2::after,
.demo-section h2::after {
  content: '';
  display: block;
  margin-top: 6px;
  height: 1px;
  background: repeating-linear-gradient(
    to right,
    var(--vp-c-divider) 0,
    var(--vp-c-divider) 1ch,
    transparent 1ch,
    transparent 1.5ch
  );
}

.workflow-steps {
  display: flex;
  align-items: stretch;
  gap: 0;
  border: 1px solid var(--vp-c-divider);
  border-radius: 8px;
  overflow: hidden;
}

.workflow-step {
  flex: 1;
  padding: 1.2rem 1.25rem;
  background: var(--vp-c-bg-soft);
  animation: step-enter 400ms ease-out both;
  position: relative;
  transition: background 180ms ease;
}

.workflow-step:hover {
  background: hsla(232, 23%, 18%, 0.6);
}

.workflow-step:hover .step-number {
  color: var(--vp-c-brand-1);
  text-shadow: 0 0 12px hsla(267, 83%, 80%, 0.3);
}

.workflow-connector {
  width: 1px;
  background: var(--vp-c-divider);
  flex-shrink: 0;
}

.workflow-step .step-number {
  display: inline-block;
  font-size: 0.7rem;
  font-weight: 700;
  letter-spacing: 0.05em;
  color: var(--vp-c-text-3);
  margin-bottom: 0.5rem;
  font-variant-numeric: tabular-nums;
  transition: color 180ms ease, text-shadow 180ms ease;
}

.workflow-step .step-number::before {
  content: '$ ';
  color: var(--vp-c-text-3);
}

.workflow-step .step-title {
  font-weight: 600;
  font-size: 1rem;
  margin-bottom: 0.35rem;
}

.workflow-step .step-desc {
  font-size: 0.85rem;
  color: var(--vp-c-text-2);
  line-height: 1.5;
}

@keyframes step-enter {
  from {
    opacity: 0;
    transform: translateY(10px);
  }
  to {
    opacity: 1;
    transform: translateY(0);
  }
}

@media (max-width: 960px) {
  .workflow-steps {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 1px;
    background: var(--vp-c-divider);
  }
  .workflow-step {
    min-width: 0;
  }
  .workflow-step:last-child {
    grid-column: 1 / -1;
  }
  .workflow-connector {
    display: none;
  }
}

@media (max-width: 640px) {
  .workflow-steps {
    grid-template-columns: 1fr;
  }
  .workflow-step:last-child {
    grid-column: auto;
  }
}

/* === Demo === */
.demo-section {
  max-width: 960px;
  margin: 3rem auto 0;
  padding: 0;
}

.demo-section p {
  color: var(--vp-c-text-2);
  margin: 0 0 1.2rem;
  line-height: 1.6;
}

.demo-window {
  border: 1px solid var(--vp-c-divider);
  border-radius: 8px;
  overflow: hidden;
  animation: step-enter 400ms ease-out 300ms both;
  box-shadow:
    0 4px 16px rgba(0, 0, 0, 0.18),
    0 20px 48px rgba(0, 0, 0, 0.14);
}

.demo-window__bar {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 8px 12px;
  background: var(--vp-c-bg-soft);
  border-bottom: 1px solid var(--vp-c-divider);
}

.demo-window__dot {
  width: 10px;
  height: 10px;
  border-radius: 50%;
}

.demo-window__dot:nth-child(1) { background: #ed8796; }
.demo-window__dot:nth-child(2) { background: #eed49f; }
.demo-window__dot:nth-child(3) { background: #a6da95; }

.demo-window__title {
  font-family: var(--tui-font-mono);
  font-size: 11px;
  color: var(--vp-c-text-3);
  margin-left: 6px;
}

.demo-window video {
  width: 100%;
  display: block;
  border: none;
  border-radius: 0;
  box-shadow: none;
  margin: 0;
}
</style>
