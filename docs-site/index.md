---
layout: home

title: SubMiner
titleTemplate: Immersion Mining Workflow for MPV

hero:
  name: SubMiner
  text: Immersion Mining for MPV
  tagline: Watch, look up a word, and get an Anki card with audio and a screenshot. Without pausing your show.
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
    details: Reads subtitle state over mpv's IPC socket. Launch with the wrapper script or the mpv plugin. There is no separate bridge process to run.
    link: /usage
    linkText: How it works
  - icon:
      src: /assets/yomitan-icon.svg
      alt: Yomitan logo
    title: Bundled Yomitan
    details: A Yomitan instance is bundled and preconfigured. Hover a word in the subtitle overlay to look it up and mine it.
    link: /mining-workflow
    linkText: Mining workflow
  - icon:
      src: /assets/anki-card.svg
      alt: Anki card icon
    title: Anki card enrichment
    details: New cards get the subtitle line, an audio clip cut to the line timing, and a screenshot from that moment.
    link: /anki-integration
    linkText: Anki integration
  - icon:
      src: /assets/highlight.svg
      alt: Highlight icon
    title: Reading annotations
    details: N+1 targeting, character-name matching, frequency highlighting, and JLPT tagging, drawn onto the subtitle line as it plays.
    link: /subtitle-annotations
    linkText: Annotation details
  - icon:
      src: /assets/video.svg
      alt: Video playback icon
    title: YouTube playback
    details: Pass a YouTube URL or a ytsearch target. SubMiner picks a subtitle track for the video and loads it.
    link: /usage#youtube-playback
    linkText: YouTube playback
  - icon:
      src: /assets/jellyfin.svg
      alt: Jellyfin icon
    title: Jellyfin integration
    details: Browse your Jellyfin library from the overlay and play a title through mpv. Subtitles and mining work the same as with local files.
    link: /jellyfin-integration
    linkText: Jellyfin setup
  - icon:
      src: /assets/subtitle-download.svg
      alt: Subtitle download icon
    title: Subtitle download and sync
    details: Search Jimaku or TsukiHime and download a track, then retime it with alass or ffsubsync. Both run from the overlay.
    link: /jimaku-integration
    linkText: Jimaku integration
  - icon:
      src: /assets/tokenization.svg
      alt: Tracking chart icon
    title: Stats dashboard
    details: A local dashboard with session history, streak calendars, word frequency, and per-series progress. You can mine cards from lines you already watched.
    link: /immersion-tracking
    linkText: Dashboard & tracking
  - icon:
      src: /assets/cross-platform.svg
      alt: Cross-platform icon
    title: Cross-platform
    details: Runs on Linux (Hyprland, Sway, X11), macOS, and Windows. Overlay positioning is handled per compositor rather than assuming one window manager.
    link: /installation
    linkText: Platform setup
---

<script setup>
import { withBase } from 'vitepress';

const demoAssetVersion = '20260819-1';
</script>

<div class="landing-shell">
  <section class="workflow-section">
    <h2>How it fits together</h2>
    <div class="workflow-steps">
      <div class="workflow-step" style="animation-delay: 0ms">
        <div class="step-number">01</div>
        <div class="step-title">Start</div>
        <div class="step-desc">Launch through the wrapper, or from an mpv setup you already have.</div>
      </div>
      <div class="workflow-connector" aria-hidden="true"></div>
      <div class="workflow-step" style="animation-delay: 60ms">
        <div class="step-number">02</div>
        <div class="step-title">Lookup</div>
        <div class="step-desc">Hover a token in the overlay to open the Yomitan popup for that word.</div>
      </div>
      <div class="workflow-connector" aria-hidden="true"></div>
      <div class="workflow-step" style="animation-delay: 120ms">
        <div class="step-number">03</div>
        <div class="step-title">Mine</div>
        <div class="step-desc">Add the word from Yomitan, or mine the whole line as a sentence card.</div>
      </div>
      <div class="workflow-connector" aria-hidden="true"></div>
      <div class="workflow-step" style="animation-delay: 180ms">
        <div class="step-number">04</div>
        <div class="step-title">Enrich</div>
        <div class="step-desc">SubMiner fills in the audio clip, the sentence, and a screenshot from that moment.</div>
      </div>
      <div class="workflow-connector" aria-hidden="true"></div>
      <div class="workflow-step" style="animation-delay: 240ms">
        <div class="step-number">05</div>
        <div class="step-title">Track</div>
        <div class="step-desc">Review past sessions and word trends, and mine anything you missed the first time.</div>
      </div>
    </div>
  </section>

  <section class="demo-section">
    <h2>See it in action</h2>
    <p>Recorded from an actual playback session: subtitle hover, lookup, and the card that comes out the other end.</p>
    <div class="demo-window">
      <div class="demo-window__bar">
        <span class="demo-window__dot"></span>
        <span class="demo-window__dot"></span>
        <span class="demo-window__dot"></span>
        <span class="demo-window__title">subminer -- playback</span>
      </div>
      <video controls playsinline preload="metadata" :poster="withBase(`/assets/minecard-poster.jpg?v=${demoAssetVersion}`)">
        <source :src="withBase(`/assets/minecard.webm?v=${demoAssetVersion}`)" type="video/webm" />
        <source :src="withBase(`/assets/minecard.mp4?v=${demoAssetVersion}`)" type="video/mp4" />
        <a :href="withBase(`/assets/minecard.webm?v=${demoAssetVersion}`)" target="_blank" rel="noreferrer">
          <img :src="withBase(`/assets/minecard.webp?v=${demoAssetVersion}`)" alt="SubMiner demo Animated fallback" style="width: 100%; height: auto;" />
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
  background: var(--tui-step-hover-bg);
}

.workflow-step:hover .step-number {
  color: var(--vp-c-brand-1);
  text-shadow: 0 0 12px var(--tui-step-hover-glow);
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
