<script setup>
import DefaultTheme from 'vitepress/theme';
import StatusLine from './components/StatusLine.vue';
import BlinkingCursor from './components/BlinkingCursor.vue';

const { Layout } = DefaultTheme;
</script>

<template>
  <Layout>
    <template #home-hero-info-after>
      <BlinkingCursor />
    </template>
    <template #layout-bottom>
      <StatusLine />
    </template>
  </Layout>
</template>
