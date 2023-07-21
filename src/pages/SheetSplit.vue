<script setup lang="ts">
import {
  NSpace,
  NUpload,
  NUploadDragger,
  NIcon,
  NText,
  NH1,
  NH2,
  UploadFileInfo,
  NButton,
  NAlert,
  NInputNumber,
} from 'naive-ui';
import { ArchiveOutline } from '@vicons/ionicons5';
import { read, utils, WorkSheet, writeFileXLSX } from 'xlsx';
import { computed, ref } from 'vue';
// import { useMessage, useDialog } from 'naive-ui';
import { useDialog } from 'naive-ui';
import Dayjs from 'dayjs'


const splitNumber = ref(5)

// const message = useMessage();
const dialog = useDialog();

const sheetA = ref<any>();
const sheetNew = ref<any>();

// 去掉所有的空格、换行、首尾空格
const removeSpaceNewline = (str: string): string => {
  if (typeof str === 'string') {
    return str.replace(/\s/g, "").trim()
  }
  return ''
}

const onUpdateFileA = (fileList: UploadFileInfo[]) => {
  if (fileList[0] && fileList[0].file) {
    fileList[0].file.arrayBuffer().then((arrayBuffer) => {
      sheetA.value = read(arrayBuffer, { type: 'buffer' });
      console.log('sheetA.value', sheetA.value);
    });
  }
};

// 处理文件
const handleNewSheet = () => {
  // ============== 处理 正畸老顾客excel ===========
  if (!sheetA.value) {
    dialog.error({
      title: '错误',
      content: '请选择一个Excel文件',
    });
    return;
  }

  // const 总合计人数sheetName = '总合计人数'
  // const 正畸老顾客_总合计人数 = sheetA.value.Sheets[总合计人数sheetName];
  // // console.log('正畸老顾客_总合计人数', 正畸老顾客_总合计人数);
  // if (!正畸老顾客_总合计人数) {
  //   dialog.error({
  //     title: '错误',
  //     content: '在「正畸老顾客」文件中没有找到「总合计人数」表格',
  //   });
  //   return;
  // }

  // console.log('targetArr', targetArr);

  // 创建并下载新的表格
  sheetNew.value = utils.book_new();
  const targetSheet = utils.json_to_sheet(targetArr)
  // console.log('targetSheet', targetSheet)
  const targetSheetName = Dayjs().format('YYYY-MM-DD')
  utils.book_append_sheet(sheetNew.value, targetSheet, targetSheetName);
  writeFileXLSX(sheetNew.value, targetSheetName + '.xlsx');

};

const disabledBtn = computed(() => {
  return !Boolean(sheetA.value)
})
</script>

<template>
  <n-space vertical style="padding: 20px; max-width: 800px; margin: auto">
    <n-h1>表格分割</n-h1>

    <n-alert type="info">
      所有文件操作仅存在客户端本地，不会上传任何文件、信息。
    </n-alert>

    <n-h2 prefix="bar"> <n-text type="primary"> 第一步 选择需要分割的Excel文件 </n-text> </n-h2>
    <n-upload multiple directory-dnd :max="1" :default-upload="false" :on-update:file-list="onUpdateFileA">
      <n-upload-dragger>
        <div style="margin-bottom: 12px">
          <n-icon size="48" :depth="3">
            <ArchiveOutline />
          </n-icon>
        </div>
        <n-text style="font-size: 16px"> 点击或者拖动文件到该区域 </n-text>
      </n-upload-dragger>
    </n-upload>

    <n-h2 prefix="bar"> <n-text type="primary"> 第二步 分割份数 </n-text> </n-h2>
    <n-input-number v-model:value="splitNumber" :min='0' :max='100' style='width: 120px;' />

    <n-h2 prefix="bar"> <n-text type="primary"> 第三步 下载 </n-text> </n-h2>
    <n-button type="primary" :disabled='disabledBtn' block @click="handleNewSheet">处理文件并下载</n-button>
  </n-space>
</template>

<style scoped>
:deep(.n-h2) {
  margin-top: 20px;
  margin-bottom: 0;
}
</style>
