<script setup>
import { useRoute, useData } from 'vitepress';
import { computed } from 'vue';

const route = useRoute();
const { page, frontmatter } = useData();

const mode = computed(() => {
  const layout = frontmatter.value.layout;
  if (layout === 'home') return 'HOME';
  return 'NORMAL';
});

const filePath = computed(() => {
  const path = route.path;
  return path === '/' ? 'index.md' : `${path.replace(/^\//, '')}.md`;
});

const section = computed(() => {
  const path = route.path;
  if (path === '/') return 'root';
  const parts = path.split('/').filter(Boolean);
  return parts[0] || 'root';
});

const lastUpdated = computed(() => {
  if (!page.value.lastUpdated) return '';
  const date = new Date(page.value.lastUpdated);
  return date.toISOString().slice(0, 10);
});
</script>

<template>
  <footer class="tui-statusline" aria-label="Page info">
    <div class="tui-statusline__left">
      <span class="tui-statusline__mode" :data-mode="mode">{{ mode }}</span>
      <span class="tui-statusline__sep"></span>
      <span class="tui-statusline__file">{{ filePath }}</span>
    </div>
    <div class="tui-statusline__right">
      <span class="tui-statusline__section">{{ section }}</span>
      <span class="tui-statusline__sep"></span>
      <span v-if="lastUpdated" class="tui-statusline__date">{{ lastUpdated }}</span>
      <span v-if="lastUpdated" class="tui-statusline__sep"></span>
      <span class="tui-statusline__branch">GPL-3.0</span>
    </div>
  </footer>
</template>
