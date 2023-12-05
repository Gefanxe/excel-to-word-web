<script setup>
import { ref, reactive, watch, toRaw } from 'vue';
import { ElMessage } from 'element-plus'
import { genFileId, switchProps, uploadBaseProps } from 'element-plus';
import { vMaska } from 'maska';
import xlsx, { read } from 'xlsx';
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import saveAs from 'save-as';
import { Plus, Minus, Delete } from '@element-plus/icons-vue';
import { liveQuery } from 'dexie';
import { useObservable } from '@vueuse/rxjs';
import { db } from '../utils/db';

/** @type { import('maska').MaskInputOptions } */
const maskOpts = {
  mask: 'A#####',
  tokens: {
    A: {
      pattern: /[A-Z]/,
      transform: str => str.toUpperCase()
    }
  }
}; // preProcess 可以做到的事, 也可以在tokens.transform裡做

// #region 資料來源
const uploadSource = ref(null); // <el-upload />
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


// #region 讀取資料區

const modeSwitch = ref(true); // false: 單一, true: 範圍

// test
function test() {
  
  console.log('test: ');
  // console.log('test: ', $Event.target);
  
  // console.log('test: ', xlsxData);
}

// 單一
const singleFields = reactive([
  {
    id: Date.now(),
    xlsxCol: '',
    tempStr: '',
    value: ''
  }
]);

const singleFieldRefs = ref([]);

watch(singleFieldRefs.value, (n, o) => {
  /** @type { HTMLInputElement } */
  let elem;
  if (n.length > 0) {
    elem = n.slice(-1)[0].input;
    elem.focus();
  }
});

function handleColAdd() {
  singleFields.push({
    id: Date.now(),
    xlsxCol: '',
    tempStr: '',
    value: ''
  });
}

function handleColSub() {
  singleFields.pop();
}

function handleDelOne(idx) {
  singleFields.splice(idx, 1);
}

function readSingle(worksheet) {
  for (let i = 0; i < singleFields.length; i++) {
    const item = singleFields[i];
    item.value = worksheet[item.xlsxCol].v;
  }
}

// 範圍

const rangeFields = reactive([]);
let xlsxData = [];

/** @type { import('vue').Ref<import('xlsx').Sheet2JSONOpts> } */
const rangeFieldSetting = ref();

rangeFieldSetting.value = {
  header: 'A',  // 設定A:代表沒有標題, 使用A,B,C....
  range: 0,    // 跳過幾行才開始解析, 或限定範圍, 例: 'A5:E6'
  defval: ''   // 使用指定的值替代null或者undefined
};

// 跳過 or 限定範圍
const isRangeFlag = ref(true);

// 限定範圍(起/迄)
const rangeStartFlag = ref('');
const rangeEndFlag = ref('');

/** @type { import('vue').Ref<import('element-plus').RowInstance> } */
const rangeDataList = ref(null);

function readDataRange(worksheet) {
  // 整個工作表輸出 json
  xlsxData = xlsx.utils.sheet_to_json(worksheet, rangeFieldSetting.value);
  const jsonSheet = xlsx.utils.json_to_sheet(xlsxData);
  const xlsxDataShow = xlsx.utils.sheet_to_html(jsonSheet, {
    id: 'sourceTable'
  });
  rangeDataList.value.$el.innerHTML = xlsxDataShow;

  rangeFields.length = 0;
  Object.keys(xlsxData[0]).forEach((item, idx) => {
    rangeFields.push({
      id: Date.now() + idx,
      rangeColumn: item,
      tempStr: ''
    });
  });
}

// 讀取來源
function handleReadSourceData() {
  if (!sourceExcel.value) return alert('沒有來源資料!');
  return new Promise((resolve, reject) => {
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
      resolve('ok');
    }
  });
}

// #endregion

// #region Excel模版

/** @type { import('vue').Ref<import('element-plus').UploadInstance> } */
const uploadExcel = ref(null);
/** @type { import('vue').UnwrapNestedRefs<import('element-plus').UploadRawFile[]> } */
const sourceExcels = reactive([]);

/**
 * 上傳 Excel 發生變化時
 * @param { import('element-plus').UploadFile } file 
 */
async function handleExcelChanged(file) {
  let isUploadOk = true;

  const fileExt = file.name.replace(/.+\.(.+)/, '$1');
  if (!/docx/i.test(fileExt)) {
    uploadExcel.value.handleRemove(file);
    isUploadOk = false;
    alert('請上傳 Word 檔, 副檔名必須為 docx! ', file.name);
  }

  if (isUploadOk) {
    sourceExcels.push(file.raw);
  }
};



// #endregion

// #region Word模版

/** @type { import('vue').Ref<import('element-plus').UploadInstance> } */
const uploadWord = ref(null);

/** @type { import('vue').UnwrapNestedRefs<import('element-plus').UploadRawFile[]> } */
const sourceWords = reactive([]);

/**
 * 將檔案讀取至 buffer
 * @param { import('element-plus').UploadRawFile } f 
 * @returns { Promise<Docxtemplater<PizZip>> }
 */
function readWordToDocTmp(f) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsArrayBuffer(f);
    reader.onload = function (e) {
      const data = new Uint8Array(reader.result);
      const zip = new PizZip(data);
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });
      resolve(doc);
    };
  });
}

/**
 * 上傳 Word 發生變化時
 * @param { import('element-plus').UploadFile } file 
 */
async function handleWordChanged(file) {
  let isUploadOk = true;

  const fileExt = file.name.replace(/.+\.(.+)/, '$1');
  if (!/docx/i.test(fileExt)) {
    uploadWord.value.handleRemove(file);
    isUploadOk = false;
    alert('請上傳 Word 檔, 副檔名必須為 docx! ', file.name);
  }

  if (isUploadOk) {
    sourceWords.push(file.raw);
  }
};

// #endregion


// #region 生成按鈕
async function handleGenerateWord() {
  // TODO: for Excel
  if (modeSwitch.value) {
    // 範圍
    const renderDatas = [];
    if (xlsxData.length > 0) {
      // 檢查有沒有要設定
      xlsxData.forEach((data) => {
        const renderData = {};
        for (const key in data) {
          if (Object.hasOwnProperty.call(data, key)) {
            const fieldValue = data[key];
            rangeFields.forEach((item) => {
              if (key === item.rangeColumn && item.tempStr !== '') {
                renderData[item.tempStr] = fieldValue;
              }
            });
          }
        }
        renderDatas.push(renderData);
      });
    } else {
      alert('沒有來源資料!');
    }

    if (renderDatas.length > 0) {

      for (let idx = 0; idx < sourceWords.length; idx++) {
        const sourceWord = sourceWords[idx];

        for (let i = 0; i < renderDatas.length; i++) {
          const renderData = renderDatas[i];
          const docx = await readWordToDocTmp(sourceWord);
          await docx.renderAsync(renderData);
          const buf = docx.getZip().generate({ type: 'blob' });
          // TODO: 名字要怎麼取??
          saveAs(buf, `Print_${idx+1}_${i+1}_${sourceWord.name}`);
        }
      }

    } else {
      alert('沒有設定資料!');
    }

  } else {
    // 單一欄位
    const haveSetFlag = false;
    const renderData = {};
    singleFields.forEach((item) => {
      if (item.xlsxCol !== '' && item.tempStr !== '') {
        renderData[item.tempStr] = item.value;
        haveSetFlag = true;
      }
    });

    if (!haveSetFlag) return alert('沒有設定資料!');

    for (let idx = 0; idx < sourceWords.length; idx++) {
      const sourceWord = sourceWords[idx];

      const docx = await readWordToDocTmp(sourceWord);
      await docx.renderAsync(renderData);
      const buf = docx.getZip().generate({ type: 'blob' });
      // TODO: 名字要怎麼取??
      saveAs(buf, `Print_${idx+1}_${sourceWord.name}`);
    }

  }
}
// #endregion

// TODO: 一鍵清除 uploadWord / sourceWords

// #region 檔案作業

// 目前作業檔案名稱
const fileName = ref(''); // 存檔名

// 檔案清單
const dataList = useObservable(
  liveQuery(() => db.mailMergeTool.toArray())
);

// const errMsg = ref('');   // 錯誤訊息顯示
const currentLoadFile = ref('');

const saveData = async ($event) => {
  const _fName = fileName.value;
  const _dataSet = {};
  
  _dataSet.sourceExcel = toRaw(sourceExcel.value);
  _dataSet.modeSwitch = toRaw(modeSwitch.value);

  if (modeSwitch.value) {
    _dataSet.isRangeFlag = toRaw(isRangeFlag.value);
    if (isRangeFlag.value) {
      _dataSet.rangeFieldSetting = toRaw(rangeFieldSetting.value);
      _dataSet.rangeFields = toRaw(rangeFields);
    } else {
      _dataSet.rangeStartFlag = toRaw(rangeStartFlag.value);
      _dataSet.rangeEndFlag = toRaw(rangeEndFlag.value);
    }
  } else {
    _dataSet.singleFields = toRaw(singleFields);
  }

  _dataSet.sourceWords = toRaw(sourceWords);

  try {
    const id = await db.mailMergeTool.add({
      name: _fName,
      dataSet: _dataSet
    });
    fileName.value = '';
    ElMessage.success({ message: '已儲存', duration: 1100 });
  } catch (error) {
    ElMessage.error(error);
  }
};

const delData = async (id) => {
  // const d = await db.mailMergeTool.toArray();
  // console.log('test: ', d);
  const dCount = await db.mailMergeTool.where('id').anyOf(id).delete();
  ElMessage.error({ message: `刪除了 ${dCount}筆`, duration: 1100 });
};

const loadData = async (item) => {
  // TODO: must clear current ref date
  currentLoadFile.value = item.name;
  
  sourceExcel.value = item.dataSet.sourceExcel;
  modeSwitch.value = item.dataSet.modeSwitch;
  if (modeSwitch.value) {
    isRangeFlag.value = item.dataSet.isRangeFlag;
    if (isRangeFlag.value) {
      rangeFieldSetting.value = item.dataSet.rangeFieldSetting;
      await handleReadSourceData();
      rangeFields.length = 0;
      item.dataSet.rangeFields.forEach((field) => {
        rangeFields.push(field);
      });
    } else {
      rangeStartFlag.value = item.dataSet.rangeStartFlag;
      rangeEndFlag.value = item.dataSet.rangeEndFlag;
    }
  } else {
    singleFields.length = 0;
    item.dataSet.singleFields.forEach((singleField) => {
      singleFields.push(singleField);
    });
    await handleReadSourceData();
  }
  item.dataSet.sourceWords.forEach((wordTmpSource) => {
    sourceWords.push(wordTmpSource);
  });

  ElMessage.success({ message: `已讀取: ${item.name}`, duration: 1100 });
};

// #endregion

</script>

<template>
  <div class="flex">
    <div class="flex-1">
      <!-- 作業操作區 -->
      <div>目前讀入的檔: <el-text class="mx-1" size="large" tag="b" type="success">{{ currentLoadFile }}</el-text></div>
      <hr>
      <div>
        <el-input v-model="fileName" placeholder="未儲存作業" style="width: 300px" /> &nbsp;
        <el-button type="primary" v-blur @click="saveData" :disabled="fileName === ''">存檔</el-button>
      </div>
  
      <div>已存檔列表:</div>
      <el-table :data="dataList" height="250" style="width: 400px" size="small" empty-text="無資料">
        <el-table-column prop="id" label="id" width="34" align="center"/>
        <el-table-column prop="name" label="名稱" />
        <el-table-column label="操作" width="100" align="center">
          <template #default="scope">
            <el-button link type="primary" size="small" v-blur @click="loadData(scope.row)">讀取</el-button>
            <el-button link type="danger" size="small" v-blur @click="delData(scope.row.id)">刪除</el-button>
          </template>
        </el-table-column>
      </el-table>
  
    </div>
    <div class="flex-1 flex flex-col">
      <div class="text-center text-xl">來源Excel檔案</div>
      <div class="flex-1 text-center">
        <div class=" color-green">
          {{ sourceExcel?.name }}
        </div>
      </div>
      <div class="text-center text-xl">設定</div>
      <div class="flex-1 text-center">
        <div v-if="fileName !== ''">{{ (modeSwitch) ? '範圍資料' : '單一欄位' }}</div>
      </div>
    </div>
    <div class="flex-1 flex flex-col">
      <div class="text-center text-xl">Excel模版檔案</div>
      <div class="flex-1 text-center">
        <div class="color-green" v-for="item in sourceExcels" :key="item.uid">{{ item.name }}</div>
      </div>
      <div class="text-center text-xl">Word模版檔案</div>
      <div class="flex-1 text-center">
        <div class="color-blue" v-for="item in sourceWords" :key="item.uid">{{ item.name }}</div>
      </div>
    </div>
  </div>

  <hr>

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

    <!-- 資料讀取模式 -->
    <div class="flex flex-col flex-items-center flex-self-stretch">
      <el-switch v-model="modeSwitch" inline-prompt style="--el-switch-on-color: #13ce66; --el-switch-off-color: #ff4949"
        active-text="範圍資料" inactive-text="單一欄位" />
      <div v-if="modeSwitch">
        <h4 class="text-center">範圍資料</h4>
        <el-row class="flex-col flex-items-center mb-4">
          <el-col>
            <el-checkbox v-model="isRangeFlag" label="跳過 or 限定範圍" size="large" />
          </el-col>
          <el-col v-if="isRangeFlag">
            跳過 <el-input-number class="mx-2" v-model="rangeFieldSetting.range" :min="0" size="small"></el-input-number> 行
          </el-col>
          <el-row v-else class="justify-center flex-nowrap">
            <span>限定範圍</span>
            <el-input v-model="rangeStartFlag" v-maska:[maskOpts] class="range-input mx-2" size="small"
              placeholder="A1" />～
            <el-input v-model="rangeEndFlag" v-maska:[maskOpts] class="range-input mx-2" size="small" placeholder="E5" />
          </el-row>
        </el-row>
        <el-row class="mb-2" ref="rangeDataList"></el-row>
        <el-row class="rangeColumnSetting flex-col">
          <div class="mb-2" v-for="item in rangeFields" :key="item.id">
            <el-input v-model="item.tempStr">
              <template #prepend>{{ item.rangeColumn }}</template>
            </el-input>
          </div>
        </el-row>
      </div>
      <div v-else>
        <h4 class="text-center">單一欄位</h4>
        <el-row class="mb-2 flex-justify-center">
          <el-button type="primary" :icon="Plus" v-blur @click="handleColAdd" />
          <el-button type="primary" :icon="Minus" v-blur @click="handleColSub" />
        </el-row>
        <div class="flex flex-items-center mb-2" v-for="(item, index) in singleFields" :key="item.id">
          <el-button type="danger" class="mr-2" :icon="Delete" circle v-blur @click="handleDelOne(index)" />
          <el-input ref="singleFieldRefs" spellcheck="false" v-model="item.xlsxCol" v-maska:[maskOpts] placeholder="ex:A1" class="mr-2"
            style="width: 60px" />
          <el-input v-model="item.tempStr" spellcheck="false" placeholder="tempStr" class="mr-4" style="width: 100px" />
          <div>{{ item.value }}</div>
        </div>
      </div>
    </div>

    <!-- 讀取資料按鈕 -->
    <div class="flex flex-col flex-self-stretch">
      <el-button type="primary" v-blur @click="handleReadSourceData">讀取資料</el-button>
      <hr>


    </div>

    <!-- Excel 模版資料 -->
    <div>
      <h3>Excel 模版檔案</h3>
      <div class="float-left">
        <el-upload ref="uploadExcel" drag :auto-upload="false" :multiple="true" :on-change="handleExcelChanged">
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
      <el-button type="primary" v-blur @click="handleGenerateWord">生成資料</el-button>
      <hr>
      <el-button type="primary" v-blur @click="test">TEST</el-button>
    </div>

  </div>

</template>

<style scoped>
:deep(.range-input.el-input) {
  width: 20%;
  display: inline;
}

:deep(.el-input__inner) {
  text-align: center;
}

.rangeColumnSetting .el-input {
  width: 220px;
}

</style>

<style>
#sourceTable,
th,
td {
  border: 1px solid;
}
</style>