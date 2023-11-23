<script setup>
import { reactive } from 'vue'

const formInline = reactive({
  ipt1: '',
  ipt2: '',
  ipt3: '',
  ipt4: '',
})

function handleInput(val) {
  formInline.ipt2 = val.toUpperCase().replace(/^([A-Z]{1}\d{,6})/g, '');
}

const onSubmit = () => {
  console.log('test: ', formInline.ipt2)
}
</script>

<template>
  <el-form :inline="true" :model="formInline" class="demo-form-inline">
    <el-form-item label="這裡輸入的英文會變成大寫">
      <el-input
        v-model="formInline.ipt1"
        placeholder=""
        :formatter="(val) => val.toUpperCase()"
        clearable 
      />
    </el-form-item>
    <el-form-item label="控制只能輸入A-Z開頭加數字">
      <el-input
        v-model="formInline.ipt2"
        placeholder=""
        :formatter="(val) => val.toUpperCase()"
        @input="handleInput"
        clearable 
      />
    </el-form-item>

    <el-form-item>
      <el-button type="primary" @click="onSubmit">Query</el-button>
    </el-form-item>
  </el-form>
</template>

<style scoped>
.demo-form-inline .el-input {
  --el-input-width: 220px;
}
</style>
