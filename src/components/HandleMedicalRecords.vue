<script setup lang="ts">
import {
  NSpace,
  NUpload,
  NUploadDragger,
  NIcon,
  NText,
  NForm,
  NFormItem,
  NH1,
  UploadFileInfo,
  NButton,
} from 'naive-ui';
import { ArchiveOutline } from '@vicons/ionicons5';
import { read, writeFileXLSX, utils, stream } from 'xlsx';
import { ref } from 'vue';
import { useMessage, useDialog } from 'naive-ui';

const message = useMessage();
const dialog = useDialog();

const sheetA = ref<any>();
const sheetB = ref<any>();
const sheetNew = ref<any>();

const onUpdateFileA = (fileList: UploadFileInfo[]) => {
  if (fileList[0] && fileList[0].file) {
    fileList[0].file.arrayBuffer().then((arrayBuffer) => {
      sheetA.value = read(arrayBuffer, { type: 'buffer' });
      console.log('sheetA.value', sheetA.value);
    });
  }
};
const onUpdateFileB = (fileList: UploadFileInfo[]) => {
  if (fileList[0] && fileList[0].file) {
    fileList[0].file.arrayBuffer().then((arrayBuffer) => {
      sheetB.value = read(arrayBuffer, { type: 'buffer' });
      console.log('sheetB.value', sheetB.value);
    });
  }
};
const handleNewSheet = () => {
  if (!sheetA.value) {
    dialog.error({
      title: '错误',
      content: '请选择「正畸老顾客」文件',
    });
    return;
  }
  const 正畸老顾客_总合计人数 = sheetA.value.Sheets['总合计人数'];
  console.log('正畸老顾客_总合计人数', 正畸老顾客_总合计人数);
  if (!正畸老顾客_总合计人数) {
    dialog.error({
      title: '错误',
      content: '在「正畸老顾客」文件中没有找到「总合计人数」表格',
    });
    return;
  }

  const 正畸老顾客_总合计人数_json = utils.sheet_to_json(
    正畸老顾客_总合计人数,
    { header: 1 }
  );
  console.log('sheet_to_json', 正畸老顾客_总合计人数_json);

  sheetNew.value = utils.book_new();
  const { A2, B2, C2, D2, E2, F2, G2, H2, I2 } = 正畸老顾客_总合计人数;
  console.log('sheetNew.value', sheetNew.value);
  // utils.book_append_sheet(sheetNew.value,{ });

  // writeFileXLSX(sheetNew.value, '处理后的数据.xlsx');

  // const newSheet = {};
  // utils.book_append_sheet(sheetNew.value);
  // console.log('sheetNew.value', sheetNew.value);

  // if (sheetNew.value) {
  //   writeFileXLSX(sheetNew.value,'处理后的数据.xlsx');
  // }
};
</script>

<template>
  <n-space vertical style="padding: 20px; max-width: 800px; margin: auto">
    <n-h1>正畸老顾客表格自动化</n-h1>
    <n-text type="success"
      >所有文件操作仅存在客户端本地，不会上传任何文件、信息。</n-text
    >
    <n-form ref="formRef" :show-feedback="false">
      <n-form-item label="第一步 选择正畸老顾客excel">
        <!-- <n-input v-model:value="model.age" @keydown.enter.prevent /> -->
        <n-upload
          multiple
          directory-dnd
          :max="1"
          :default-upload="false"
          :on-update:file-list="onUpdateFileA"
        >
          <n-upload-dragger>
            <div style="margin-bottom: 12px">
              <n-icon size="48" :depth="3">
                <ArchiveOutline />
              </n-icon>
            </div>
            <n-text style="font-size: 16px"> 点击或者拖动文件到该区域 </n-text>
          </n-upload-dragger>
        </n-upload>
      </n-form-item>
      <n-form-item label="第二步 选择每月统计的患者查询excel">
        <!-- <n-input v-model:value="model.age" @keydown.enter.prevent /> -->
        <n-upload
          multiple
          directory-dnd
          :max="1"
          :default-upload="false"
          :on-update:file-list="onUpdateFileB"
        >
          <n-upload-dragger>
            <div style="margin-bottom: 12px">
              <n-icon size="48" :depth="3">
                <ArchiveOutline />
              </n-icon>
            </div>
            <n-text style="font-size: 16px"> 点击或者拖动文件到该区域 </n-text>
          </n-upload-dragger>
        </n-upload>
      </n-form-item>
      <n-form-item label="第三步 处理文件并下载">
        <n-button type="primary" block @click="handleNewSheet"
          >处理文件</n-button
        >
      </n-form-item>
    </n-form>
  </n-space>
</template>

<style scoped></style>
