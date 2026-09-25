<script setup>
import { useRoute, useData } from 'vitepress';
import { computed, onMounted, ref } from 'vue';
import { formatStatusLineDate, formatStatusLineFilePath } from '../status-line';

const route = useRoute();
const { frontmatter } = useData();

const mode = computed(() => {
  const layout = frontmatter.value.layout;
  if (layout === 'home') return 'HOME';
  return 'NORMAL';
});

const filePath = computed(() => {
  return formatStatusLineFilePath(route.path);
});

const section = computed(() => {
  const path = route.path;
  if (path === '/') return 'root';
  const parts = path.split('/').filter(Boolean);
  return parts[0] || 'root';
});

// Set on the client only, so the prerendered HTML never bakes in the build date.
const today = ref('');
onMounted(() => {
  today.value = formatStatusLineDate(new Date());
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
      <span v-if="today" class="tui-statusline__date">{{ today }}</span>
      <span v-if="today" class="tui-statusline__sep"></span>
      <span class="tui-statusline__branch">GPL-3.0</span>
    </div>
  </footer>
</template>
