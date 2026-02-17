---
layout: home

title: SubMiner
titleTemplate: Immersion Mining Workflow for MPV

hero:
  name: SubMiner
  text: Built for Immersion Mining
  tagline: A self-contained MPV overlay for Japanese study. Look up words, mine cards, and enrich Anki without breaking playback flow.
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
    - theme: alt
      text: Is This For Me?
      link: "#who-this-is-for"

features:
  - icon:
      src: /assets/mpv.svg
      alt: mpv icon
    title: Built for mpv
    details: Connects directly to mpv over IPC — tracks subtitles in real time, observes playback properties, and renders a self-contained overlay with everything bundled in a single application.
  - icon:
      src: /assets/yomitan-icon.svg
      alt: Yomitan logo
    title: Yomitan Integration
    details: Hover over any word in the subtitles to trigger Yomitan dictionary lookups — get instant definitions without leaving the video player.
  - icon:
      src: /assets/anki-card.svg
      alt: Anki card icon
    title: Anki Card Enrichment
    details: Add a word from Yomitan and SubMiner automatically updates the card with the sentence, audio clip, screenshot, and translation — no extra steps needed.
  - icon:
      src: /assets/dual-layer.svg
      alt: Dual layer icon
    title: Dual-Layer Subtitle System
    details: Visible overlay with styled, interactive subtitles — plus an invisible layer that aligns with mpv's own subtitle rendering for seamless click-through lookup.
  - icon:
      src: /assets/highlight.svg
      alt: Highlight icon
    title: N+1 Word Highlighting
    details: Highlights words you already know from your Anki deck, making it easy to spot new vocabulary and identify true N+1 sentences during immersion.
  - icon:
      src: /assets/texthooker.svg
      alt: Texthooker icon
    title: Texthooker & WebSocket
    details: Built-in texthooker page that receives subtitles over WebSocket — use it as a clipboard inserter for Yomitan or connect external tools for real-time subtitle streaming.
  - icon:
      src: /assets/subtitle-download.svg
      alt: Subtitle download icon
    title: Subtitle Download & Sync
    details: Search and download Japanese subtitles from Jimaku, then sync them to the audio with alass or ffsubsync — all from within the player.
  - icon:
      src: /assets/keyboard.svg
      alt: Keyboard icon
    title: Keyboard-Driven Workflow
    details: Mine sentences, copy subtitles, cycle display modes, and trigger field grouping — all from configurable keyboard shortcuts without touching the mouse.
---

<style>
.demo-section {
  max-width: 960px;
  margin: 2rem auto 0;
  padding: 0 24px;
}

.demo-section h2 {
  font-size: 1.5rem;
  font-weight: 600;
  margin-bottom: 0.75rem;
}

.demo-section p {
  color: var(--vp-c-text-2);
  margin-bottom: 1rem;
}

.demo-section video {
  width: 100%;
  border-radius: 8px;
  border: 1px solid var(--vp-c-divider);
}

.workflow-section {
  max-width: 960px;
  margin: 3rem auto 0;
  padding: 0 24px 3rem;
}

.workflow-section h2 {
  font-size: 1.5rem;
  font-weight: 600;
  margin-bottom: 1.5rem;
}

.workflow-steps {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
  gap: 1rem;
}

.workflow-step {
  padding: 1rem;
  border-radius: 8px;
  border: 1px solid var(--vp-c-divider);
  background: var(--vp-c-bg-soft);
}

.workflow-step .step-number {
  font-size: 0.8rem;
  font-weight: 700;
  color: var(--vp-c-brand-1);
  margin-bottom: 0.25rem;
}

.workflow-step .step-title {
  font-weight: 600;
  margin-bottom: 0.25rem;
}

.workflow-step .step-desc {
  font-size: 0.875rem;
  color: var(--vp-c-text-2);
}
</style>

<div class="demo-section">

## What SubMiner Is For

SubMiner is for people who learn Japanese by watching subtitled content in mpv and want a low-friction mining loop:

- stay inside the video while looking up words
- send mined content to Anki quickly
- keep media context (audio, screenshot, timestamp, subtitle context) attached to each card
- reduce tool switching between player, dictionary, and card workflow

</div>

<div class="workflow-section">

## Project Goals

<div class="workflow-steps">
  <div class="workflow-step">
    <div class="step-title">1. Keep Immersion Continuous</div>
    <div class="step-desc">Minimize context switching by making lookup and mining happen directly over mpv subtitles.</div>
  </div>
  <div class="workflow-step">
    <div class="step-title">2. Preserve Card Quality</div>
    <div class="step-desc">Attach sentence context, audio, image, and translation so mined cards stay reviewable and useful long-term.</div>
  </div>
  <div class="workflow-step">
    <div class="step-title">3. Support Real Workflows</div>
    <div class="step-desc">Handle day-to-day immersion needs: subtitle management, syncing, known-word awareness, and keyboard-first controls.</div>
  </div>
  <div class="workflow-step">
    <div class="step-title">4. Stay Configurable</div>
    <div class="step-desc">Offer defaults that work out of the box, while still letting advanced users shape behavior around their note type and setup.</div>
  </div>
  <div class="workflow-step">
    <div class="step-title">5. Evolve Safely</div>
    <div class="step-desc">Use a modular TypeScript codebase and automated tests so features can ship faster without breaking core mining behavior.</div>
  </div>
</div>

</div>

<div class="demo-section">

## See It in Action

SubMiner sits as a transparent overlay on top of mpv. Subtitles appear as interactive, clickable text — click a word to look it up with Yomitan, then add it to Anki with one click.

<video controls playsinline preload="metadata" poster="/assets/demo-poster.jpg">
  <source :src="'/assets/card-mine.webm'" type="video/webm" />
  Your browser does not support the video tag.
</video>

</div>

<div class="workflow-section">

## Who This Is For

- learners using mpv as their main immersion player
- users who already rely on Yomitan + AnkiConnect
- miners who care about preserving context on cards, not just raw words

SubMiner is likely overkill if you only want lightweight lookup without card enrichment, overlay controls, or integrated workflow tooling.

</div>

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
    <div class="step-desc">Hover over a word in the subtitle overlay and hold Shift to trigger a Yomitan dictionary lookup.</div>
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
