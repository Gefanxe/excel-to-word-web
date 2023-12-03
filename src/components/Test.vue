<script setup>
import { ref, reactive } from 'vue';
import { liveQuery } from 'dexie';
import { useObservable } from '@vueuse/rxjs';
import xlsx, { read } from 'xlsx';
import { vMaska } from 'maska';
import { db } from '../utils/db';

const formInline = reactive({
  ipt1: '',
})

const uploadSource = ref(null);
const sourceExcel = ref(null);

const handleSourceExceed = (files) => {
  console.log('on exceed!', files[0]);
  const fileExtension = files[0].name.replace(/.+\.(.+)/, '$1');
  if (/xlsx/i.test(fileExtension)) {
    uploadSource.value.clearFiles();
    const file = files[0];
    file.uid = genFileId();
    uploadSource.value.handleStart(file);
  } else {
    alert('請上傳Excel檔, 副檔名必須為 xlsx!');
  }
}

const handleSourceChanged = (file) => {
  console.log('on Changed!', file);
  // 剖析副檔名(file.raw.type分不出來excel與word)
  const fileExtension = file.name.replace(/.+\.(.+)/, '$1');
  if (!/xlsx/i.test(fileExtension)) {
    uploadSource.value.handleRemove(file);
    alert('請上傳Excel檔, 副檔名必須為 xlsx!');
    return;
  }
  sourceExcel.value = file.raw;
  handleReadSourceData();
};

// 讀取來源
let arrayData;
function handleReadSourceData() {
  if (!sourceExcel.value) return alert('沒有來源資料!');

  const reader = new FileReader();
  reader.readAsArrayBuffer(sourceExcel.value);
  reader.onload = function (e) {
    arrayData = new Uint8Array(reader.result);
  }
}

const onSubmit = () => {
  // console.log('test: ', formInline.ipt1)

  const book = xlsx.read(arrayData, { type: 'array' });
  const sheets = book.SheetNames[0];
  const worksheet = book.Sheets[sheets];

  const colValue = worksheet[formInline.ipt1].v;
  console.log('read value: ', colValue);
}

/** @type { import('maska').MaskInputOptions } */
const maskOpts = {
  mask: 'A#####',
  tokens: {
    A: {
      pattern: /[A-Z]/,
      transform: str => str.toUpperCase()
    }
  }
};

const dataList = useObservable(
  liveQuery(() => db.mailMergeTool.toArray())
);

const fileName = ref(null);
const errMsg = ref(null);

const showErr = (msg) => {
  errMsg.value.innerHTML = msg;
  const id = setTimeout(() => {
    errMsg.value.innerHTML = '';
    clearTimeout(id);
  }, 3000);

};
const saveData = async () => {
  const _fName = fileName.value.value;
  if (_fName === '') return alert('no name!');
  try {
    const id = await db.mailMergeTool.add({
      name: _fName
    });
  } catch (error) {
    showErr(error);
  }
};

const delData = async (id) => {
  // const d = await db.mailMergeTool.toArray();
  // console.log('test: ', d);
  const dCount = await db.mailMergeTool.where('id').anyOf(id).delete();

  showErr(`刪除了 ${dCount}筆`);
};

</script>

<template>
  <el-form :inline="true" :model="formInline" class="demo-form-inline">
    <el-form-item label="上傳檔案">
      <el-upload ref="uploadSource" drag :limit="1" :auto-upload="false" :on-change="handleSourceChanged"
          :on-exceed="handleSourceExceed">
          <el-icon>
            <Plus />
          </el-icon>
          <div class="el-upload__text">
            將資料來源拖放到此<br>
            或<em>點擊</em>上傳
          </div>
          <template #tip>
            <div class="el-upload__tip">
              僅接受副檔名為xlsx的Excel檔案
            </div>
          </template>
        </el-upload>
    </el-form-item>
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
  <hr>
  <!-- 作業操作區 -->
  <div>
    <input type="text" ref="fileName" placeholder="存檔名"> <br>
    <button @click="saveData">存檔</button>
  </div>
  <div>已存檔列表:</div>
  <div>
    <div v-for="item in dataList" :key="item.id">
      <a href="javascript:void(0);">讀檔</a>
      &nbsp;
      <span>{{ item.name }}</span>
      &nbsp;
      <a href="javascript:void(0);" @click="delData(item.id)">刪除</a>
    </div>
  </div>
  <div>
    <div>目前讀入的檔:</div>
    <div></div>
  </div>
  <div>
    <div>操作結果:</div>
    <div ref="errMsg" class="error"></div>
  </div>
</template>

<style scoped>
.demo-form-inline .el-input {
  --el-input-width: 220px;
}

.error {
  color: red;
}
</style>
