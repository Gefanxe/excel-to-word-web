<script setup>
import { ref } from 'vue';
import { genFileId } from 'element-plus';
import xlsx, { read } from 'xlsx';
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import saveAs from 'save-as';


// #region 資料來源
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
};

// #endregion


// #region 讀取資料按鈕

function handleReadSourceData() {
  if (!sourceExcel.value) return alert('沒有來源資料!');
  // console.log(sourceExcel.value);
  readData(sourceExcel.value);
}

function readData(file) {
  const reader = new FileReader();
  reader.readAsArrayBuffer(file);
  reader.onload = function (e) {
    const data = new Uint8Array(reader.result);
    const book = xlsx.read(data, { type: 'array' });
    const sheets = book.SheetNames[0];
    const worksheet = book.Sheets[sheets];

    const opts = {
      header: 'A',    // 沒有標題, 使用A,B,C....
      range: 4,       // 跳過幾行才開始解析
      // range: 'A5:E6',     //限定範圍
      defval: ''      // 使用指定的值替代null或者undefined
    };

    // 整個工作表輸出 json
    const xlsxData = xlsx.utils.sheet_to_json(worksheet, opts);
    console.log(xlsxData);
  };
}
// #endregion


// #region Word模版

const uploadWord = ref(null);
const sourceWord = ref(null);
const handleWordChanged = (file, files) => {
  // TODO: 先判斷是否有多個檔案 files.length > 0
  // console.log('on Changed!', file);
  if (files[1]) {

    console.log('on Changed! s', files[1].name);
  }
};

// #endregion


</script>

<template>
  <div class="flex flex-justify-evenly flex-items-center">

    <!-- 來源 Excel 檔案 -->
    <div>
      <h3>來源Excel檔案</h3>
      <div class="float-left">
        <el-upload ref="uploadSource" drag :limit="1" :auto-upload="false" :on-change="handleSourceChanged" :on-exceed="handleSourceExceed">
          <el-icon><Plus /></el-icon>
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
      </div>
    </div>

    <!-- TODO: 資料讀取模式 -->

    <!-- 讀取資料按鈕 -->
    <div>
      <el-button type="primary" @click="handleReadSourceData">get data</el-button>
    </div>

    <!-- Excel 模版資料 -->
    
    
    <!-- Word 模版資料 -->
    <div>
      <h3>Word 模版檔案</h3>
      <div class="float-left">
        <el-upload ref="uploadWord" drag :auto-upload="false" :multiple="true" :on-change="handleWordChanged">
          <el-icon><Plus /></el-icon>
          <div class="el-upload__text">
            將資料來源拖放到此<br>
            或<em>點擊</em>上傳
          </div>
          <template #tip>
            <div class="el-upload__tip">
              僅接受副檔名為docx的Word檔案
            </div>
          </template>
        </el-upload>
      </div>
    </div>

  </div>
</template>

<style scoped></style>
