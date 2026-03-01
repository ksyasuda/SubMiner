# Feature Demos

Short recordings of SubMiner's key features and integrations from real playback sessions.

<script setup>
const v = '20260301-1';
</script>

## Anki Card Mining & Enrichment

Mine vocabulary cards from Yomitan or directly from subtitle lines. SubMiner automatically attaches the sentence, a timing-accurate audio clip, a screenshot, and a translation.

<video controls playsinline preload="metadata" :poster="`/assets/minecard-poster.jpg?v=${demoAssetVersion}`">
  <source :src="`/assets/minecard.webm?v=${demoAssetVersion}`" type="video/webm" />
  <source :src="`/assets/minecard.mp4?v=${demoAssetVersion}`" type="video/mp4" />
  <a :href="`/assets/minecard.webm?v=${demoAssetVersion}`" target="_blank" rel="noreferrer">
    <img :src="`/assets/minecard.webp?v=${demoAssetVersion}`" alt="SubMiner demo Animated fallback" style="width: 100%; height: auto;" />
  </a>
</video>

::: info VIDEO COMING SOON
:::

## Subtitle Download & Sync

Search and download subtitles from Jimaku, then automatically synchronize them with alass or ffsubsync — all from within SubMiner.

<!-- <video controls playsinline preload="metadata" :poster="`/assets/demos/subtitle-sync-poster.jpg?v=${v}`">
  <source :src="`/assets/demos/subtitle-sync.webm?v=${v}`" type="video/webm" />
  <source :src="`/assets/demos/subtitle-sync.mp4?v=${v}`" type="video/mp4" />
</video> -->

::: info VIDEO COMING SOON
:::

## Jellyfin Integration

Browse your Jellyfin library, cast to devices, and launch playback directly from SubMiner. Watch progress syncs back to your Jellyfin server.

<!-- <video controls playsinline preload="metadata" :poster="`/assets/demos/jellyfin-poster.jpg?v=${v}`">
  <source :src="`/assets/demos/jellyfin.webm?v=${v}`" type="video/webm" />
  <source :src="`/assets/demos/jellyfin.mp4?v=${v}`" type="video/mp4" />
</video> -->

::: info VIDEO COMING SOON
:::

## Texthooker

Open subtitles in an external texthooker page for use with browser-based tools and extensions alongside the overlay.

<!-- <video controls playsinline preload="metadata" :poster="`/assets/demos/texthooker-poster.jpg?v=${v}`">
  <source :src="`/assets/demos/texthooker.webm?v=${v}`" type="video/webm" />
  <source :src="`/assets/demos/texthooker.mp4?v=${v}`" type="video/mp4" />
</video> -->

::: info VIDEO COMING SOON
:::

<style>
video {
  width: 100%;
  border-radius: 12px;
  border: 1px solid var(--vp-c-divider);
  box-shadow: 0 18px 44px rgba(0, 0, 0, 0.28);
  margin: 0.75rem 0 2.5rem;
}

h2 {
  margin-top: 2.5rem !important;
}
</style>
