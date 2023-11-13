<script setup>
import { onMounted, ref } from 'vue'
import xlsx, { read } from 'xlsx';
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import saveAs from 'save-as';

// defineProps({
//   msg: String,
// })

const msg = ref('Hello');

const eFile = ref(null);
const wFile = ref(null);

onMounted(() => {
  eFile.value.onchange = function (evt) {
    const reader = new FileReader();
    reader.readAsArrayBuffer(evt.target.files[0]);
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
  };

  wFile.value.onchange = function (evt) {
    const reader = new FileReader();
    reader.readAsArrayBuffer(evt.target.files[0]);
    reader.onload = function (e) {
      const data = new Uint8Array(reader.result);
      
      const zip = new PizZip(data);
      const doc = new Docxtemplater(zip, {
          paragraphLoop: true,
          linebreaks: true,
      });
      doc.render({
          my_name: '黃大大',
          my_tel: '0933567634'
      });
      const buf = doc.getZip().generate({
        type: 'blob'
          // type: "nodebuffer",
          // compression: "DEFLATE",
      });

      saveAs(buf, 'newDoc.docx');
    };
  };
});


</script>

<template>
  <h1>{{ msg }}</h1>

  <div>
    Excel:
    <input ref="eFile" type="file" name="eFile" id="eFile">
    <hr>
    Word:
    <input ref="wFile" type="file" name="wFile" id="wFile">
  </div>
</template>

<style scoped></style>
