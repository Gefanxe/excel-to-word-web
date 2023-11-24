<script setup>
import { reactive } from 'vue';
import { vMaska } from 'maska';

const formInline = reactive({
  ipt1: '',
  ipt2: '',
  ipt3: '',
  ipt4: '',
})

function handleInput(val) {
  const upperVal = val.toUpperCase();
  console.log('1', upperVal);
  if (upperVal.length < 2 && !/^[A-Z]/.test(upperVal)) {
    formInline.ipt2 = '';
  }
  if (upperVal.length > 1 && !/^[A-Z]\d{1,5}/.test(upperVal)) {
    console.log('2', upperVal);
    formInline.ipt2 = upperVal.slice(0, -1);
  }
}

const onSubmit = () => {
  console.log('test: ', formInline.ipt1)
}

/** @type { import('maska').MaskInputOptions } */
const maskOpts = {
  mask: 'A#####',
  tokens: 'A:[A-Z]',
  preProcess: v => v.toUpperCase()
};

</script>

<template>
  <el-form :inline="true" :model="formInline" class="demo-form-inline">
    <el-form-item label="控制只能輸入A-Z開頭加數字">
      <el-input
        v-model="formInline.ipt1"
        v-maska:[maskOpts]
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
