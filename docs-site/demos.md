# Feature demos

Short recordings from real playback sessions.

Some vocabulary for what follows. _Yomitan_ is the pop-up dictionary. _Jimaku_ is a community subtitle database. _alass_ and _ffsubsync_ retime subtitles against the audio. _Jellyfin_ is a self-hosted media server. A _texthooker_ is a web page that mirrors the current subtitle as selectable text so browser tools can read it.

<script setup>
import { withBase } from 'vitepress';

const v = '20260819-1';
</script>

## Anki card mining and enrichment

Mine a card from Yomitan or straight from a subtitle line. SubMiner attaches the sentence, an audio clip cut to the line timing, and a screenshot.

<video controls playsinline preload="metadata" :poster="withBase(`/assets/minecard-poster.jpg?v=${v}`)">
  <source :src="withBase(`/assets/minecard.webm?v=${v}`)" type="video/webm" />
  <source :src="withBase(`/assets/minecard.mp4?v=${v}`)" type="video/mp4" />
  <a :href="withBase(`/assets/minecard.webm?v=${v}`)" target="_blank" rel="noreferrer">
    <img :src="withBase(`/assets/minecard.webp?v=${v}`)" alt="SubMiner demo Animated fallback" style="width: 100%; height: auto;" />
  </a>
</video>

## Subtitle download and sync

Search Jimaku, download a track, then retime it with alass or ffsubsync without leaving SubMiner.

<!-- <video controls playsinline preload="metadata" :poster="withBase(`/assets/demos/subtitle-sync-poster.jpg?v=${v}`)">
  <source :src="withBase(`/assets/demos/subtitle-sync.webm?v=${v}`)" type="video/webm" />
  <source :src="withBase(`/assets/demos/subtitle-sync.mp4?v=${v}`)" type="video/mp4" />
</video> -->

::: info VIDEO COMING SOON
:::

## Jellyfin integration

Browse your Jellyfin library, cast to a device, and start playback from SubMiner. Watch progress goes back to the Jellyfin server.

<!-- <video controls playsinline preload="metadata" :poster="withBase(`/assets/demos/jellyfin-poster.jpg?v=${v}`)">
  <source :src="withBase(`/assets/demos/jellyfin.webm?v=${v}`)" type="video/webm" />
  <source :src="withBase(`/assets/demos/jellyfin.mp4?v=${v}`)" type="video/mp4" />
</video> -->

::: info VIDEO COMING SOON
:::

## Texthooker

Mirror subtitles to an external texthooker page so browser extensions can read them while the overlay runs.

<!-- <video controls playsinline preload="metadata" :poster="withBase(`/assets/demos/texthooker-poster.jpg?v=${v}`)">
  <source :src="withBase(`/assets/demos/texthooker.webm?v=${v}`)" type="video/webm" />
  <source :src="withBase(`/assets/demos/texthooker.mp4?v=${v}`)" type="video/mp4" />
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
