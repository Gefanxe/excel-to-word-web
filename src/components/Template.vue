<script setup>
import { ref } from 'vue';

const msg = ref('Template');
</script>

<template>
  <div class="text-center">
    {{ msg }}
  </div>
</template>

<style scoped>
div {
  background-color: skyblue;
}
</style>
