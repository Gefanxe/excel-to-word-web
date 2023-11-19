<script setup>
import { ref, reactive } from 'vue';
import { genFileId, switchProps, uploadBaseProps } from 'element-plus';
import xlsx, { read } from 'xlsx';
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import saveAs from 'save-as';
import { Plus, Minus } from '@element-plus/icons-vue';


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

const modeSwitch = ref(false);

const singleFields = reactive([
  {
    xlsxCol: 'A1',
    value: ''
  }
]);

function handleColAdd() {
  singleFields.push({
    xlsxCol: '',
    value: ''
  });
}

function handleColSub() {
  singleFields.pop();
}

function readDataRange(worksheet) {
  const opts = {
    header: 'A',    // 沒有標題, 使用A,B,C....
    range: 4,       // 跳過幾行才開始解析
    // range: 'A5:E6',     //限定範圍
    defval: ''      // 使用指定的值替代null或者undefined
  };

  // 整個工作表輸出 json
  const xlsxData = xlsx.utils.sheet_to_json(worksheet, opts);
  console.log(xlsxData);
}

function readSingle(worksheet) {
  for (let i = 0; i < singleFields.length; i++) {
    const item = singleFields[i];
    item.value = worksheet[item.xlsxCol].v;
  }
}

// 讀取來源
function handleReadSourceData() {
  if (!sourceExcel.value) return alert('沒有來源資料!');

  const reader = new FileReader();
  reader.readAsArrayBuffer(sourceExcel.value);
  reader.onload = function (e) {
    const data = new Uint8Array(reader.result);
    const book = xlsx.read(data, { type: 'array' });
    const sheets = book.SheetNames[0];
    const worksheet = book.Sheets[sheets];

    if (modeSwitch.value) {
      // 範圍資料讀取
      readDataRange(worksheet);
    } else {
      // 單一欄位讀取
      readSingle(worksheet);
    }
  }
}

// #endregion


// #region Word模版

/** @type { import('vue').Ref<import('element-plus').UploadInstance> } */
const uploadWord = ref(null);

/** @type { import('vue').Ref<import('element-plus').UploadRawFile[]> } */
const sourceWords = ref([]);

/** @type { import('vue').Ref<Docxtemplater<PizZip>[]> } */
const docxTemps = ref([]);



/**
 * 將檔案讀取至 buffer
 * @param { import('element-plus').UploadRawFile } f 
 */
function readWordToBuffer(f) {
  const reader = new FileReader();
  reader.readAsArrayBuffer(f);
  reader.onload = function (e) {
    const data = new Uint8Array(reader.result);

    const zip = new PizZip(data);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });
    
    docxTemps.value.push(doc);
  };
}

/**
 * 上傳 Word 發生變化時
 * @param { import('element-plus').UploadFile } file 
 * @param { import('element-plus').UploadFiles } files 
 */
function handleWordChanged (file, files) {
  files.forEach((f) => {
    const fileExt = f.name.replace(/.+\.(.+)/, '$1');
    if (!/docx/i.test(fileExt)) {
      uploadWord.value.handleRemove(f);
      console.log('請上傳 Word 檔, 副檔名必須為 docx! ', f.name);
    } else {
      sourceWords.value.push(f.raw);
    }
  });

  sourceWords.value.forEach((wordTemp) => {
    readWordToBuffer(wordTemp);
  });

};

// #endregion


// #region 生成按鈕
function handleGenerateData () {
  docxTemps.value.forEach((docx) => {

    docx.render({
      xxxx: '',
      yyy: ''
    });

    const buf = docx.getZip().generate({
      type: 'blob'
    })

    // TODO: 名字要怎麼取??
    saveAs(buf, 'newDoc.docx');
  });
}
// #endregion

// TODO: 一鍵清除 uploadWord / sourceWords

</script>

<template>
  <div class="flex flex-justify-evenly flex-items-center py-6">

    <!-- 來源 Excel 檔案 -->
    <div>
      <h3>來源Excel檔案</h3>
      <div class="float-left">
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
      </div>
    </div>

    <!-- TODO: 資料讀取模式 -->
    <div class="flex flex-col flex-self-stretch">
      <el-switch
        v-model="modeSwitch"
        inline-prompt
        style="--el-switch-on-color: #13ce66; --el-switch-off-color: #ff4949"
        active-text="範圍資料"
        inactive-text="單一欄位"
      />
      <div v-if="modeSwitch">
        範圍資料
      </div>
      <div v-else>
        <h4>單一欄位</h4>
        <div class="mb-2">
          <el-button type="primary" :icon="Plus" @click="handleColAdd"/>
          <el-button type="primary" :icon="Minus" @click="handleColSub"/>
        </div>
        <div class="flex flex-items-center mb-2" v-for="(item, index) in singleFields">
          <el-input
            v-model="item.xlsxCol"
            placeholder="ex:A1"
            maxlength="2"
            minlength="2"
            class="mr-4"
            style="width: 60px"
            clearable
          />
          <div>{{ item.value }}</div>
        </div>
      </div>
    </div>

    <!-- 讀取資料按鈕 -->
    <div class="flex flex-col flex-self-stretch">
      <el-button type="primary" @click="handleReadSourceData">讀取資料</el-button>
    </div>

    <!-- Excel 模版資料 -->


    <!-- Word 模版資料 -->
    <div>
      <h3>Word 模版檔案</h3>
      <div class="float-left">
        <el-upload ref="uploadWord" drag :auto-upload="false" :multiple="true" :on-change="handleWordChanged">
          <el-icon>
            <Plus />
          </el-icon>
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

    <!-- 生成資料按鈕 -->
    <div>
      <el-button type="primary" @click="handleGenerateData">生成資料</el-button>
    </div>

  </div>
</template>

<style scoped></style>
