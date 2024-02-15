<script setup>
// #region import
import { ref, reactive, watch, toRaw } from 'vue';
import { ElMessage } from 'element-plus'
import { genFileId } from 'element-plus';
import { vMaska } from 'maska';
import xlsx from 'xlsx';
import PizZip from "pizzip";
import { Renderer } from 'xlsx-renderer'
import Docxtemplater from "docxtemplater";
import nzhhk from 'nzh/hk';
import saveAs from 'save-as';
import { Plus, Minus, EditPen } from '@element-plus/icons-vue';
import { liveQuery } from 'dexie';
import { useObservable } from '@vueuse/rxjs';
import { db } from '../utils/db';
// #endregion

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

// 判斷字串是否為數字
function isNumeric(str) {
  if (typeof str != "string") return false;
  return !isNaN(str) && !isNaN(parseFloat(str));
}

// #region 資料來源
/** @type { import('vue').Ref<import('element-plus').UploadInstance> } */
const uploadSource = ref(null);
const sourceExcel = ref(null);
const sourceSheetName = ref(null);

const handleSourceExceed = (files) => {
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
  uploadSource.value.handleRemove(file);
};

// #endregion

// #region 讀取資料區

const modeSwitch = ref(true); // false: 單一, true: 範圍

// test
function test() {
}

// 單一
const singleFields = reactive([
  {
    id: Date.now(),
    xlsxCol: '',
    tempStr: '',
    value: '',
    isNum: false,
    nzhhk: false
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
    value: '',
    isNum: false,
    nzhhk: false
  });
}

function handleColSub() {
  singleFields.pop();
}

// 單一: 設定輸出的部份檔名
function handleSetPartOfName(item) {
  if (item.xlsxCol === '' || item.tempStr === '' || item.value === '') {
    alert('請填妥資料');
  } else {
    partOfFileName.value = item.value;
  }
}


function handleSetPartOfNameForRange(item) {
  if (item.tempStr === '') {
    alert('請填妥資料');
  } else {
    partOfFileName.value = item.tempStr;
  }
}

function readSingle(worksheet) {
  for (let i = 0; i < singleFields.length; i++) {
    const item = singleFields[i];
    if (worksheet[item.xlsxCol]) {
      const val = worksheet[item.xlsxCol].v;
      console.log('val: ', val);
      if (isNumeric(val)) item.isNum = true;
      // if (item.nzhhk) item.value = nzhhk.encodeB(val);
      item.value = (item.nzhhk) ? nzhhk.encodeB(val) : val;
    } else {
      ElMessage.error({ message: `欄位${item.xlsxCol}沒有資料`, duration: 1100 });
    }
  }
}

// 範圍

const rangeFields = reactive([]);
const xlsxData = ref([]);

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

/** @type { import('vue').Ref<HTMLDivElement> } */
// const rangeDataList = ref(null);

function readDataRange(worksheet, rangeFieldsFromLoad) {
  xlsxData.value.length = 0;
  // 整個工作表輸出 json
  const sheetJson = xlsx.utils.sheet_to_json(worksheet, rangeFieldSetting.value);
  sheetJson.forEach((item) => {
    xlsxData.value.push(item);
  });

  console.log('xlsxData: ', xlsxData.value)
  rangeFields.length = 0;
  if (rangeFieldsFromLoad) {
    rangeFieldsFromLoad.forEach(item => {
      rangeFields.push(item);
    });
    rangeFields.forEach(item => {
      if (item.nzhhk) {
        item.nzhhk = false;
        handleTransNumberR(item);
      }
    });
  } else {
    Object.keys(xlsxData.value[0]).forEach((item, idx) => {
      rangeFields.push({
        id: Date.now() + idx,
        rangeColumn: item,
        tempStr: '',
        isNum: isNumeric(xlsxData.value[0][item]),
        nzhhk: false
      });
    });
  }
}

// 讀取來源
function handleReadSourceData(loadData) {
  if (!sourceExcel.value) return alert('沒有來源資料!');
  if (!sourceSheetName.value) return alert('請輸入來源工作表名稱!');
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsArrayBuffer(sourceExcel.value);
    reader.onload = function (e) {
      const data = new Uint8Array(reader.result);
      const book = xlsx.read(data, { type: 'array' });
      // const sheets = book.SheetNames[0]; // get sheetname
      const worksheet = book.Sheets[sourceSheetName.value];

      if (modeSwitch.value) {
        // 範圍資料讀取
        readDataRange(worksheet, loadData?.rangeFields || null);
      } else {
        // 單一欄位讀取
        readSingle(worksheet);
      }
      resolve('ok');
    }
  });
}

// (單一) 阿拉伯數字 <=> 國字 轉換
function handleTransNumberS(item) {
  item.value = (!item.nzhhk) ? nzhhk.encodeB(item.value) : nzhhk.decodeB(item.value).toString();
  item.nzhhk = !item.nzhhk;
}

// (範圍) 阿拉伯數字 <=> 國字 轉換
function handleTransNumberR(item) {
  const col = item.rangeColumn;
  xlsxData.value.forEach((data) => {
    data[col] = (!item.nzhhk) ? nzhhk.encodeB(data[col]) : nzhhk.decodeB(data[col]).toString();
  });
  item.nzhhk = !item.nzhhk;
}

// #endregion

// #region Excel模版

/** @type { import('vue').Ref<import('element-plus').UploadInstance> } */
const uploadExcel = ref(null);
/** @type { import('vue').UnwrapNestedRefs<import('element-plus').UploadRawFile[]> } */
const sourceExcels = reactive([]);

/**
 * 將Excel檔案讀取至 buffer
 * @param { import('element-plus').UploadRawFile } f 
 * @returns { Promise<Docxtemplater<PizZip>> }
 */
function readExcelToXlsTmp(f) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsArrayBuffer(f);
    reader.onload = function (e) {
      const data = new Uint8Array(reader.result);
      resolve(data);
    };
  });
}

/**
 * 上傳 Excel 發生變化時
 * @param { import('element-plus').UploadFile } file 
 */
async function handleExcelChanged(file) {
  let isUploadOk = true;

  const fileExt = file.name.replace(/.+\.(.+)/, '$1');
  if (!/xlsx/i.test(fileExt)) {
    uploadExcel.value.handleRemove(file);
    isUploadOk = false;
    alert('請上傳 Excel 檔, 副檔名必須為 xlsx! ', file.name);
  }

  if (isUploadOk) {
    sourceExcels.push(file.raw);
    uploadExcel.value.handleRemove(file);
  }
};

// #endregion

// #region Word模版

/** @type { import('vue').Ref<import('element-plus').UploadInstance> } */
const uploadWord = ref(null);

/** @type { import('vue').UnwrapNestedRefs<import('element-plus').UploadRawFile[]> } */
const sourceWords = reactive([]);

/**
 * 將Word檔案讀取至 buffer
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
    uploadWord.value.handleRemove(file);
  }
};

// #endregion

// #region 生成

const partOfFileName = ref('');

// 生成excel
async function generateExcel(renderDatas) {
  if (modeSwitch.value) {
    for (let idx = 0; idx < sourceExcels.length; idx++) {
      const sourceExcel = sourceExcels[idx];

      for (let i = 0; i < renderDatas.length; i++) {
        const renderData = renderDatas[i];
        const buffer = await readExcelToXlsTmp(sourceExcel);
        const report = await new Renderer().renderFromArrayBuffer(buffer, renderData);
        const buf = await report.xlsx.writeBuffer();

        const partOfName = (partOfFileName.value !== '') ? renderData[partOfFileName.value] : `${(idx + 1)}`
        saveAs(new Blob([buf]), `RESULT_${partOfName}_${i + 1}_${sourceExcel.name}`);
      }
    }
  } else {
    const renderData = renderDatas[0];
    for (let idx = 0; idx < sourceExcels.length; idx++) {
      const sourceExcel = sourceExcels[idx];

      const buffer = await readExcelToXlsTmp(sourceExcel);
      const report = await new Renderer().renderFromArrayBuffer(buffer, renderData);
      const buf = await report.xlsx.writeBuffer();

      const partOfName = (partOfFileName.value !== '') ? partOfFileName.value : `${(idx + 1)}`
      saveAs(new Blob([buf]), `RESULT_${partOfName}_${idx + 1}_${sourceExcel.name}`);
    }
  }
}

// 生成word
async function generateWord(renderDatas) {

  if (modeSwitch.value) {
    for (let idx = 0; idx < sourceWords.length; idx++) {
      const sourceWord = sourceWords[idx];

      for (let i = 0; i < renderDatas.length; i++) {
        const renderData = renderDatas[i];
        const docx = await readWordToDocTmp(sourceWord);
        await docx.renderAsync(renderData);
        const buf = docx.getZip().generate({ type: 'blob' });

        const partOfName = (partOfFileName.value !== '') ? renderData[partOfFileName.value] : `${(idx + 1)}`
        saveAs(buf, `RESULT_${partOfName}_${sourceWord.name}`);
      }
    }
  } else {
    const renderData = renderDatas[0];
    for (let idx = 0; idx < sourceWords.length; idx++) {
      const sourceWord = sourceWords[idx];

      const docx = await readWordToDocTmp(sourceWord);
      await docx.renderAsync(renderData);
      const buf = docx.getZip().generate({ type: 'blob' });

      const partOfName = (partOfFileName.value !== '') ? partOfFileName.value : `${(idx + 1)}`
      saveAs(buf, `RESULT_${partOfName}_${sourceWord.name}`);
    }
  }
}

// 生成按鈕函數
async function handleGenerate() {
  if (partOfFileName.value === '') return alert('"部份檔名"沒有設定');
  if (sourceExcels.length === 0 && sourceWords.length === 0) return alert('沒有載入任何模版');

  const wordRenderDatas = [];
  const excelRenderDatas = [];

  if (modeSwitch.value) {
    // 範圍
    if (xlsxData.value.length === 0) return alert('沒有來源資料!');

    // for excel
    if (sourceExcels.length > 0) {
      xlsxData.value.forEach((data) => {
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
        excelRenderDatas.push(renderData);
      });
      if (excelRenderDatas.length === 0) return alert('資料沒有設定完全!');
    }

    // for word
    if (sourceWords.length > 0) {
      xlsxData.value.forEach((data) => {
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
        wordRenderDatas.push(renderData);
      });
      if (wordRenderDatas.length === 0) return alert('資料沒有設定完全!');
    }

  } else {
    // 單一欄位

    // for excel
    if (sourceExcels.length > 0) {
      let haveSetFlag = true;
      const renderData = {};
      singleFields.forEach((item) => {
        if (item.xlsxCol !== '' && item.tempStr !== '' && item.value !== '') {
          renderData[item.tempStr] = item.value;
        } else {
          haveSetFlag = false;
        }
      });
      if (!haveSetFlag) return alert('資料沒有設定完全!');
      excelRenderDatas.push(renderData);
    }

    // for word
    if (sourceWords.length > 0) {
      let haveSetFlag = true;
      const renderData = {};

      singleFields.forEach((item) => {
        if (item.xlsxCol !== '' && item.tempStr !== '' && item.value !== '') {
          renderData[item.tempStr] = item.value;
        } else {
          haveSetFlag = false;
        }
      });
      if (!haveSetFlag) return alert('資料沒有設定完全!');
      wordRenderDatas.push(renderData);
    }
  }
  // console.log('excelRenderDatas: ', excelRenderDatas);
  // console.log('wordRenderDatas: ', wordRenderDatas);
  if (sourceExcels.length > 0) await generateExcel(excelRenderDatas);
  if (sourceWords.length > 0) await generateWord(wordRenderDatas);

}

// #endregion

// #region 清除

function handleClear() {
  // upload 元件
  uploadSource.value.clearFiles();
  uploadExcel.value.clearFiles();
  uploadWord.value.clearFiles();

  // 資料
  sourceExcel.value = null;
  sourceExcels.length = 0;
  sourceWords.length = 0;

  // 設定
  currentLoadFile.value = '';
  partOfFileName.value = '';

  if (modeSwitch.value) {
    rangeFields.length = 0;
    xlsxData.value.length = 0;
  } else {
    while (singleFields.length > 0) {
      singleFields.pop();
    }
  }

}


// #endregion

// #region 檔案作業

// 目前作業檔案名稱
const fileName = ref(''); // 存檔名

// 檔案清單
const dataList = useObservable(
  liveQuery(() => db.mailMergeTool.toArray())
);

const currentLoadFile = ref('');

// 存檔
const saveData = async () => {
  const _fName = fileName.value;
  const _dataSet = {};

  _dataSet.sourceExcel = toRaw(sourceExcel.value);
  _dataSet.modeSwitch = toRaw(modeSwitch.value);

  _dataSet.partOfFileName = toRaw(partOfFileName.value);

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
  _dataSet.sourceExcels = toRaw(sourceExcels);

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

// 刪檔
const delData = async (id) => {
  const dCount = await db.mailMergeTool.where('id').anyOf(id).delete();
  ElMessage.error({ message: `刪除了 ${dCount}筆`, duration: 1100 });
};

// 讀檔
const loadData = async (item) => {
  // clear current setting
  handleClear();

  currentLoadFile.value = item.name;

  sourceExcel.value = item.dataSet.sourceExcel;
  modeSwitch.value = item.dataSet.modeSwitch;
  partOfFileName.value = item.dataSet.partOfFileName || '';
  if (modeSwitch.value) {
    isRangeFlag.value = item.dataSet.isRangeFlag;
    if (isRangeFlag.value) {
      rangeFieldSetting.value = item.dataSet.rangeFieldSetting;
      await handleReadSourceData({rangeFields: item.dataSet.rangeFields});
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

  item.dataSet.sourceExcels.forEach((excelTmpSource) => {
    sourceExcels.push(excelTmpSource);
  });

  ElMessage.success({ message: `已讀取: ${item.name}`, duration: 1100 });
  console.log('loaded');
};

// #endregion

</script>

<template>
  <div class="flex">

    <!-- 檔案操作 -->
    <div class="mr-6">
      <!-- 作業操作區 -->
      <div>目前讀入的檔: <el-text class="mx-1" size="large" tag="b" type="success">{{ currentLoadFile }}</el-text></div>
      <hr>
      <div>
        <el-input v-model="fileName" placeholder="未儲存作業" style="width: 300px" /> &nbsp;
        <el-button type="primary" v-blur @click="saveData" :disabled="fileName === ''">存檔</el-button>
      </div>

      <div>已存檔列表:</div>
      <el-table :data="dataList" height="300" style="width: 400px" size="small" empty-text="無資料">
        <el-table-column prop="id" label="id" width="34" align="center" />
        <el-table-column prop="name" label="名稱" />
        <el-table-column label="操作" width="100" align="center">
          <template #default="scope">
            <el-button link type="primary" size="small" v-blur @click="loadData(scope.row)">讀取</el-button>
            <el-button link type="danger" size="small" v-blur @click="delData(scope.row.id)">刪除</el-button>
          </template>
        </el-table-column>
      </el-table>

    </div>

    <!-- 目前載入資訊1 -->
    <div class="mr-6 flex flex-col">
      <div class="text-center text-xl color-gray-400">來源Excel檔案</div>
      <div class="flex-1 flex flex-col justify-center text-center">
        <div class=" color-green-600">
          {{ sourceExcel?.name }}
        </div>
        <div v-if="!!sourceExcel">
          <el-input v-model="sourceSheetName" placeholder="工作表名稱" />
        </div>
      </div>

      <div class="text-center text-xl">設定</div>
      <div class="flex-1 text-center">
        <div>輸入的部份檔名:</div>
        <div>{{ partOfFileName }}</div>
      </div>
    </div>

    <!-- 目前載入資訊2 -->
    <div class="mr-6 flex flex-col">

      <div class="text-center text-xl color-gray-400">Excel模版檔案</div>
      <div class="flex-1 flex flex-col justify-center text-center">
        <div class="color-green" v-for="item in sourceExcels" :key="item.uid">{{ item.name }}</div>
      </div>
      <div class="text-center text-xl color-gray-400">Word模版檔案</div>
      <div class="flex-1 flex flex-col justify-center text-center">
        <div class="color-blue" v-for="item in sourceWords" :key="item.uid">{{ item.name }}</div>
      </div>

    </div>

    <!-- 來源 / 模版資料 -->
    <div class="mr-6 flex flex-col flex-justify-between">

      <!-- 讀取來源 -->
      <div>
        <el-upload ref="uploadSource" drag :limit="1" :auto-upload="false" :on-change="handleSourceChanged"
          :on-exceed="handleSourceExceed">
          <div class="el-upload__text">
            <div class="text-xl font-extrabold">來源Excel</div>
            將資料來源拖放到此<br>
            或<em>點擊</em>上傳
          </div>
        </el-upload>
      </div>

      <!-- Excel 模版檔案 -->
      <div>
        <el-upload ref="uploadExcel" drag :auto-upload="false" :multiple="true" :on-change="handleExcelChanged">
          <div class="el-upload__text">
            <div class="text-xl font-extrabold">Excel模版</div>
            將資料來源拖放到此<br>
            或<em>點擊</em>上傳
          </div>
        </el-upload>
      </div>

      <!-- Word 模版檔案 -->
      <div>
        <el-upload ref="uploadWord" drag :auto-upload="false" :multiple="true" :on-change="handleWordChanged">
          <div class="el-upload__text">
            <div class="text-xl font-extrabold">Word模版</div>
            將資料來源拖放到此<br>
            或<em>點擊</em>上傳
          </div>
        </el-upload>
      </div>
    </div>

    <!-- 資料讀取模式 -->
    <div class="flex flex-col flex-items-center flex-self-stretch">

      <el-switch v-model="modeSwitch" size="large" inline-prompt
        style="--el-switch-on-color: #13ce66; --el-switch-off-color: #ff4949" active-text="範圍資料" inactive-text="單一欄位" />
      <div v-if="modeSwitch">
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
      </div>
      <div v-else>
        <el-row class="mb-2 flex-justify-center">
          <el-button type="primary" :icon="Plus" v-blur @click="handleColAdd" />
          <el-button type="primary" :icon="Minus" v-blur @click="handleColSub" />
        </el-row>
      </div>

      <!-- 按鈕 -->
      <div>
        <el-button type="primary" v-blur @click="handleReadSourceData">讀取資料</el-button>
        <el-button type="success" v-blur @click="handleGenerate">生成資料</el-button>
        <hr>
        <el-button type="danger" v-blur @click="handleClear">清除資料</el-button>
        <el-button type="primary" v-blur @click="test">TEST</el-button>
        <hr>
      </div>
    </div>

  </div>

  <hr>

  <div class="flex flex-items-start">
    <!-- 載入資訊後設定 -->
    <div class="flex-1" v-if="modeSwitch">
      <el-row v-if="xlsxData.length > 0">
        <table border="1">
          <thead>
            <tr>
              <th v-for="item in rangeFields" :key="item.id">
                <el-button :type="(partOfFileName !== '' && item.tempStr === partOfFileName) ? 'primary' : 'info'"
                  class="mr-2" v-blur @click="handleSetPartOfNameForRange(item)">{{ item.rangeColumn }}</el-button>
              </th>
            </tr>
            <tr>
              <th class="text-center" v-for="(item, i) in rangeFields" :key="`tool1_${i}`">
                <el-input v-model="item.tempStr" style="width: 100px;" /> <br>
              </th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="(tr, idx) in xlsxData" :key="`tr_${idx}`">
              <td class="text-center" v-for="(item, i) in rangeFields" :key="`td_${i}`">{{ tr[item.rangeColumn] }}</td>
            </tr>
            <tr>
              <td class="text-center" v-for="(item, i) in rangeFields" :key="`tool2_${i}`">
                <el-button
                  v-if="item.isNum"
                  :type="(!item.nzhhk) ? 'primary' : 'success'"
                  v-blur
                  @click="handleTransNumberR(item)">
                  {{ (!item.nzhhk) ? '轉國字' : '轉數字' }}
                </el-button>
              </td>
            </tr>
          </tbody>
        </table>

      </el-row>
    </div>
    <div class="flex-1" v-else>
      <table class="text-center">
        <tr>
          <td>SET</td>
          <td>功能</td>
          <td>欄</td>
          <td>變數</td>
          <td>資料</td>
        </tr>
        <tr v-for="(item, index) in singleFields" :key="item.id">
          <td>
            <el-button type="info" :icon="EditPen" v-blur @click="handleSetPartOfName(item)" />
          </td>
          <td>
            <el-button
              v-if="item.isNum"
              :type="(!item.nzhhk) ? 'primary' : 'success'"
              v-blur
              @click="handleTransNumberS(item)">
              {{ (!item.nzhhk) ? '轉國字' : '轉數字' }}
            </el-button>
          </td>
          <td>
            <el-input ref="singleFieldRefs" spellcheck="false" v-model="item.xlsxCol" v-maska:[maskOpts]
              placeholder="ex:A1" style="width: 60px" />
          </td>
          <td>
            <el-input v-model="item.tempStr" spellcheck="false" placeholder="tempStr" style="width: 100px" />
          </td>
          <td>
            <div class="mx-2">{{ item.value }}</div>
          </td>
        </tr>
      </table>
    </div>

  </div>
</template>

<style scoped>
:deep(.el-upload-dragger) {
  padding-top: 30px;
  padding-bottom: 30px;
}

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