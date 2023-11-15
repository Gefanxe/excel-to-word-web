<script setup>
import { ref } from 'vue'
const value1 = ref(true)
</script>

<template>
  <el-switch v-model="value1" />
</template>

<style scoped>
</style>
