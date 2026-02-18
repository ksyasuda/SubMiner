---
layout: home

title: SubMiner
titleTemplate: Immersion Mining Workflow for MPV

hero:
  name: SubMiner
  text: Immersion Mining for MPV
  tagline: Look up words, mine to Anki, and enrich cards with context — all without leaving the video.
  image:
    src: /assets/SubMiner.png
    alt: SubMiner logo
  actions:
    - theme: brand
      text: Get Started
      link: /installation
    - theme: alt
      text: Mining Workflow
      link: /mining-workflow

features:
  - icon:
      src: /assets/mpv.svg
      alt: mpv icon
    title: Built for mpv
    details: Connects via IPC to track subtitles in real time and render a self-contained overlay — everything bundled in a single application.
  - icon:
      src: /assets/yomitan-icon.svg
      alt: Yomitan logo
    title: Yomitan Integration
    details: Hover over any word in the subtitle overlay to trigger dictionary lookups — instant definitions without leaving the player.
  - icon:
      src: /assets/anki-card.svg
      alt: Anki card icon
    title: Anki Card Enrichment
    details: Add a word from Yomitan and SubMiner fills in the sentence, audio clip, screenshot, and translation automatically.
  - icon:
      src: /assets/dual-layer.svg
      alt: Dual layer icon
    title: Dual-Layer Subtitles
    details: Interactive visible overlay plus an invisible layer aligned with mpv's own rendering for seamless click-through lookup.
  - icon:
      src: /assets/highlight.svg
      alt: Highlight icon
    title: N+1 Highlighting
    details: Marks words you already know from your Anki deck so you can spot new vocabulary and identify N+1 sentences at a glance.
  - icon:
      src: /assets/texthooker.svg
      alt: Texthooker icon
    title: Texthooker & WebSocket
    details: Built-in texthooker page that receives subtitles over WebSocket — use it as a clipboard inserter or connect external tools.
  - icon:
      src: /assets/subtitle-download.svg
      alt: Subtitle download icon
    title: Subtitle Download & Sync
    details: Search and download Japanese subtitles from Jimaku, then sync to audio with alass or ffsubsync — all from within the player.
  - icon:
      src: /assets/keyboard.svg
      alt: Keyboard icon
    title: Keyboard-Driven
    details: Mine sentences, copy subtitles, cycle display modes, and trigger field grouping — all from configurable shortcuts.
---

<style>
.demo-section {
  max-width: 960px;
  margin: 0 auto;
  padding: 0 24px;
}

.demo-section h2 {
  font-size: 1.5rem;
  font-weight: 600;
  margin-bottom: 0.5rem;
  letter-spacing: -0.01em;
}

.demo-section p {
  color: var(--vp-c-text-2);
  margin-bottom: 1.25rem;
  line-height: 1.6;
}

.demo-section video {
  width: 100%;
  border-radius: 12px;
  border: 1px solid var(--vp-c-divider);
  box-shadow: 0 4px 24px rgba(0, 0, 0, 0.15);
}

.workflow-section {
  max-width: 960px;
  margin: 4rem auto 0;
  padding: 0 24px 3rem;
}

.workflow-section h2 {
  font-size: 1.5rem;
  font-weight: 600;
  margin-bottom: 1.5rem;
  letter-spacing: -0.01em;
}

.workflow-steps {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 1px;
  background: var(--vp-c-divider);
  border-radius: 12px;
  overflow: hidden;
}

@media (max-width: 768px) {
  .workflow-steps {
    grid-template-columns: 1fr;
  }
}

.workflow-step {
  padding: 1.25rem 1.5rem;
  background: var(--vp-c-bg-soft);
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
</style>

<div class="workflow-section">

## How It Works

<div class="workflow-steps">
  <div class="workflow-step">
    <div class="step-number">01</div>
    <div class="step-title">Watch</div>
    <div class="step-desc">Play a video in mpv. SubMiner connects via IPC and captures subtitles in real time.</div>
  </div>
  <div class="workflow-step">
    <div class="step-number">02</div>
    <div class="step-title">Look Up</div>
    <div class="step-desc">Hover over a word in the subtitle overlay and hold Shift to trigger a Yomitan lookup.</div>
  </div>
  <div class="workflow-step">
    <div class="step-number">03</div>
    <div class="step-title">Mine</div>
    <div class="step-desc">Add the word to Anki from Yomitan. SubMiner detects the new card automatically.</div>
  </div>
  <div class="workflow-step">
    <div class="step-number">04</div>
    <div class="step-title">Enrich</div>
    <div class="step-desc">SubMiner fills in the sentence, audio clip, screenshot, and translation — no extra steps.</div>
  </div>
</div>

</div>

<div class="demo-section">

## CLI Quick Reference

```bash
subminer                    # Default picker + playback workflow
subminer jellyfin -d        # Jellyfin cast discovery mode (foreground)
subminer jellyfin -p        # Jellyfin play picker
subminer yt -o ~/subs URL   # YouTube subcommand with output dir shortcut
subminer doctor             # Dependency/config/socket health checks
subminer config path        # Active config file path
subminer config show        # Print active config
subminer mpv status         # MPV socket readiness
subminer texthooker         # Texthooker-only mode
```

See [Usage](/usage) for full command and option coverage.

## See It in Action

<video controls playsinline preload="metadata" poster="/assets/demo-poster.jpg">
  <source :src="'/assets/card-mine.webm'" type="video/webm" />
  Your browser does not support the video tag.
</video>

</div>
