<script setup>
import { ref, reactive, watch, toRaw } from 'vue';
import { ElMessage } from 'element-plus'
import { genFileId, switchProps, uploadBaseProps } from 'element-plus';
import { vMaska } from 'maska';
import xlsx, { read } from 'xlsx';
import PizZip from "pizzip";
import { Renderer } from 'xlsx-renderer'
import Docxtemplater from "docxtemplater";
import saveAs from 'save-as';
import { Plus, Minus, Delete, EditPen } from '@element-plus/icons-vue';
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
/** @type { import('vue').Ref<import('element-plus').UploadInstance> } */
const uploadSource = ref(null); // <el-upload />
const sourceExcel = ref(null);

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
          // TEST: 名字要怎麼取??
          const partOfName = (partOfFileName.value !== '') ? renderData.data[partOfFileName.value] : `${(idx + 1)}`
          saveAs(new Blob([buf]), `RESULT_${partOfName}_${i+1}_${sourceExcel.name}`);
        }
      }
  } else {
    const renderData = renderDatas[0];
    for (let idx = 0; idx < sourceExcels.length; idx++) {
      const sourceExcel = sourceExcels[idx];

      const buffer = await readExcelToXlsTmp(sourceExcel);
      const report = await new Renderer().renderFromArrayBuffer(buffer, renderData);
      const buf = await report.xlsx.writeBuffer();
      // TODO: 名字要怎麼取??
      const partOfName = (partOfFileName.value !== '') ? partOfFileName.value : `${(idx + 1)}`
      saveAs(new Blob([buf]), `RESULT_${partOfName}_${i+1}_${sourceExcel.name}`);
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
          // TODO: 名字要怎麼取??
          saveAs(buf, `Print_${idx+1}_${i+1}_${sourceWord.name}`);
        }
      }
  } else {
    const renderData = renderDatas[0];
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

// 生成按鈕函數
async function handleGenerate() {
  
  if (sourceExcels.length === 0 && sourceWords.length === 0) return alert('沒有載入任何模版');

  const wordRenderDatas = [];
  const excelRenderDatas = [];

  if (modeSwitch.value) {
    // 範圍
    if (xlsxData.length === 0) return alert('沒有來源資料!');

    // for excel
    if (sourceExcels.length > 0) {
      xlsxData.forEach((data) => {
        const renderData = {
          data: {}
        };
        for (const key in data) {
          if (Object.hasOwnProperty.call(data, key)) {
            const fieldValue = data[key];
            rangeFields.forEach((item) => {
              if (key === item.rangeColumn && item.tempStr !== '') {
                renderData.data[item.tempStr] = fieldValue;
              }
            });
          }
        }
        excelRenderDatas.push(renderData);
      });
      if (excelRenderDatas.length === 0) return alert('沒有設定資料!');
    }

    // for word
    if (sourceWords.length > 0) {
      xlsxData.forEach((data) => {
        const renderData = {
          data: {}
        };
        for (const key in data) {
          if (Object.hasOwnProperty.call(data, key)) {
            const fieldValue = data[key];
            rangeFields.forEach((item) => {
              if (key === item.rangeColumn && item.tempStr !== '') {
                renderData.data[item.tempStr] = fieldValue;
              }
            });
          }
        }
        wordRenderDatas.push(renderData);
      });
      if (wordRenderDatas.length === 0) return alert('沒有設定資料!');
    }

  } else {
    // 單一欄位

    // for excel
    if (sourceExcels.length > 0) {}

    // for word
    if (sourceWords.length > 0) {
      const haveSetFlag = false;
      const renderData = {};
      singleFields.forEach((item) => {
        if (item.xlsxCol !== '' && item.tempStr !== '') {
          renderData[item.tempStr] = item.value;
          haveSetFlag = true;
        }
      });
      if (!haveSetFlag) return alert('沒有設定資料!');
      wordRenderDatas.push(renderData);
    }
  }

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
    xlsxData.length = 0;
    rangeDataList.value.$el.innerHTML = '';
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
  // const d = await db.mailMergeTool.toArray();
  // console.log('test: ', d);
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
  partOfFileName.value = item.dataSet.partOfFileName;
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

  item.dataSet.sourceExcels.forEach((excelTmpSource) => {
    sourceExcels.push(excelTmpSource);
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
        <div>輸入的部份檔名: <span>{{ partOfFileName }}</span></div>
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

  <div class="flex flex-justify-evenly flex-items-start py-6">

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
            <el-button type="info" class="mr-2" :icon="EditPen" circle v-blur @click="handleSetPartOfNameForRange(item)" />
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
          <el-button type="info" class="mr-2" :icon="EditPen" circle v-blur @click="handleSetPartOfName(item)" />
          <el-input ref="singleFieldRefs" spellcheck="false" v-model="item.xlsxCol" v-maska:[maskOpts] placeholder="ex:A1" class="mr-2"
            style="width: 60px" />
          <el-input v-model="item.tempStr" spellcheck="false" placeholder="tempStr" class="mr-4" style="width: 100px" />
          <div>{{ item.value }}</div>
        </div>
      </div>
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

    <!-- 按鈕 -->
    <div>
      <el-button type="primary" v-blur @click="handleReadSourceData">讀取資料</el-button>
      <hr>
      <el-button type="success" v-blur @click="handleGenerate">生成資料</el-button>
      <hr>
      <el-button type="danger" v-blur @click="handleClear">清除資料</el-button>
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