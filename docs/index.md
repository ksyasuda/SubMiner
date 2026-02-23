---
layout: home

title: SubMiner
titleTemplate: Immersion Mining Workflow for MPV

hero:
  name: SubMiner
  text: Immersion Mining for MPV
  tagline: Watch media, mine vocabulary, and build cards without leaving the scene.
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
    details: Tracks subtitles through mpv IPC in real time, with a single launch path and no external bridge services.
  - icon:
      src: /assets/yomitan-icon.svg
      alt: Yomitan logo
    title: Yomitan Integration
    details: Keep your flow moving with instant word lookups and context-aware card creation directly from subtitles.
  - icon:
      src: /assets/anki-card.svg
      alt: Anki card icon
    title: Anki Card Enrichment
    details: Auto-fills card fields with subtitle sentence, clipping, image, and translation so you can focus on learning.
  - icon:
      src: /assets/dual-layer.svg
      alt: Dual layer icon
    title: Three-Plane Overlay Stack
    details: Secondary context plane + visible interactive layer + invisible interaction plane, each with independent behavior and startup state.
  - icon:
      src: /assets/highlight.svg
      alt: Highlight icon
    title: N+1 Highlighting
    details: Surfaces known words from your deck so unknown targets stand out during immersion sessions.
  - icon:
      src: /assets/tokenization.svg
      alt: Tokenization icon
    title: Immersion Tracking
    details: Captures subtitle and mining telemetry to SQLite, with daily/monthly rollups for progress clarity.
  - icon:
      src: /assets/subtitle-download.svg
      alt: Subtitle download icon
    title: Subtitle Download & Sync
    details: Pull and synchronize subtitles with Jimaku plus alass/ffsubsync in one cohesive workflow.
  - icon:
      src: /assets/keyboard.svg
      alt: Keyboard icon
    title: Keyboard-Driven
    details: Run lookups, mining actions, clipping, and workflow toggles with one configurable shortcut surface.
  - icon:
      src: /assets/texthooker.svg
      alt: Texthooker icon
    title: Texthooker & WebSocket
    details: Stream subtitles in real time to browser tools via local WebSocket and keep your stack integrated.
---

<script setup>
const demoAssetVersion = '20260223-2';
</script>

<div class="landing-shell">
  <section class="workflow-section">
    <h2>How it fits together</h2>
    <div class="workflow-steps">
      <div class="workflow-step">
        <div class="step-number">01</div>
        <div class="step-title">Start</div>
        <div class="step-desc">Launch with the wrapper or existing mpv setup and keep subtitles in sync.</div>
      </div>
      <div class="workflow-step">
        <div class="step-number">02</div>
        <div class="step-title">Lookup</div>
        <div class="step-desc">Hover or click a token in the interactive overlay to open Yomitan context.</div>
      </div>
      <div class="workflow-step">
        <div class="step-number">03</div>
        <div class="step-title">Mine</div>
        <div class="step-desc">Create cards from Yomitan or mine sentence cards directly from subtitle lines.</div>
      </div>
      <div class="workflow-step">
        <div class="step-number">04</div>
        <div class="step-title">Enrich</div>
        <div class="step-desc">Automatically attach timing-accurate audio, sentence text, and visual evidence.</div>
      </div>
      <div class="workflow-step">
        <div class="step-number">05</div>
        <div class="step-title">Track</div>
        <div class="step-desc">Review immersion history and repeat high-value patterns over time.</div>
      </div>
    </div>
  </section>

  <section class="demo-section">
    <h2>See it in action</h2>
    <p>Subtitles, lookup flow, and card enrichment from a real playback session.</p>
    <video controls playsinline preload="metadata" :poster="`/assets/minecard-poster.jpg?v=${demoAssetVersion}`">
      <source :src="`/assets/minecard.webm?v=${demoAssetVersion}`" type="video/webm" />
      <source :src="`/assets/minecard.mp4?v=${demoAssetVersion}`" type="video/mp4" />
      <a :href="`/assets/minecard.webm?v=${demoAssetVersion}`" target="_blank" rel="noreferrer">
        <img :src="`/assets/minecard.gif?v=${demoAssetVersion}`" alt="SubMiner demo GIF fallback" style="width: 100%; height: auto;" />
      </a>
    </video>
  </section>
</div>

<style>
@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=Space+Grotesk:wght@500;600;700&display=swap');

.landing-shell {
  max-width: 1120px;
  margin: 0 auto;
  padding: 0.5rem 1rem 4rem;
}

.landing-shell,
.landing-shell .step-title,
.landing-shell h1,
.landing-shell h2 {
  font-family: 'Manrope', 'Arial', sans-serif;
}

.step-title,
.step-number {
  font-family: 'Space Grotesk', 'Manrope', 'Arial', sans-serif;
  letter-spacing: -0.01em;
}

.demo-section {
  max-width: 960px;
  margin: 0 auto;
  padding: 2rem 0 0;
}

.demo-section h2 {
  font-size: 1.45rem;
  font-weight: 600;
  margin-bottom: 0.5rem;
  letter-spacing: -0.01em;
}

.demo-section p {
  color: var(--vp-c-text-2);
  margin: 0 0 0.9rem;
  line-height: 1.6;
}

.demo-section video {
  width: 100%;
  border-radius: 12px;
  border: 1px solid var(--vp-c-divider);
  box-shadow: 0 18px 44px rgba(0, 0, 0, 0.28);
  animation: card-enter 380ms ease-out;
}

.workflow-section {
  margin: 2.4rem auto 0;
  padding: 0 0 2.5rem;
}

.workflow-section h2 {
  font-size: 1.45rem;
  font-weight: 600;
  margin-bottom: 1rem;
  letter-spacing: -0.01em;
}

.workflow-steps {
  display: grid;
  grid-template-columns: repeat(5, 1fr);
  gap: 1px;
  background: var(--vp-c-divider);
  border-radius: 12px;
  overflow: hidden;
}

@media (max-width: 960px) {
  .workflow-steps {
    grid-template-columns: repeat(3, 1fr);
  }
}

@media (max-width: 640px) {
  .workflow-steps {
    grid-template-columns: 1fr;
  }
}

.workflow-step {
  padding: 1.1rem 1.25rem;
  background: var(--vp-c-bg-soft);
  animation: card-enter 330ms ease-out;
}

.workflow-step .step-number {
  display: inline-block;
  font-size: 0.7rem;
  font-weight: 700;
  letter-spacing: 0.05em;
  color: var(--vp-c-brand-1);
  margin-bottom: 0.5rem;
  font-variant-numeric: tabular-nums;
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

@keyframes card-enter {
  from {
    opacity: 0.8;
    transform: translateY(8px);
  }

  to {
    opacity: 1;
    transform: translateY(0);
  }
}
</style>
