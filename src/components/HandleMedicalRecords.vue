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
import { read, utils, WorkSheet, writeFileXLSX } from 'xlsx';
import { ref } from 'vue';
// import { useMessage, useDialog } from 'naive-ui';
import { useDialog } from 'naive-ui';
import Dayjs from 'dayjs'

// const message = useMessage();
const dialog = useDialog();

const sheetA = ref<any>();
const sheetB = ref<any>();
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
const onUpdateFileB = (fileList: UploadFileInfo[]) => {
  if (fileList[0] && fileList[0].file) {
    fileList[0].file.arrayBuffer().then((arrayBuffer) => {
      sheetB.value = read(arrayBuffer, { type: 'buffer' });
      console.log('sheetB.value', sheetB.value);
    });
  }
};

// 处理文件
const handleNewSheet = () => {
  // ============== 处理 正畸老顾客excel ===========
  if (!sheetA.value) {
    dialog.error({
      title: '错误',
      content: '请选择「正畸老顾客」文件',
    });
    return;
  }
  const 总合计人数sheetName = '总合计人数'
  const 正畸老顾客_总合计人数 = sheetA.value.Sheets[总合计人数sheetName];
  // console.log('正畸老顾客_总合计人数', 正畸老顾客_总合计人数);
  if (!正畸老顾客_总合计人数) {
    dialog.error({
      title: '错误',
      content: '在「正畸老顾客」文件中没有找到「总合计人数」表格',
    });
    return;
  }

  // 表头
  let tableHeaderA = utils.sheet_to_json(正畸老顾客_总合计人数, { header: 1 })[1] as string[]
  tableHeaderA = tableHeaderA.map((item: string) => removeSpaceNewline(item))
  console.log('tableHeaderA', tableHeaderA);
  const 客户姓名key = tableHeaderA[1]
  const 初诊日期key = tableHeaderA[4]
  const 复诊日期key = tableHeaderA[5]
  const 剩余治疗次数key = tableHeaderA[6]

  let tableHeaderB = utils.sheet_to_json(sheetB.value.Sheets[sheetB.value.SheetNames[0]], { header: 1 })[0] as any[]
  tableHeaderB = tableHeaderB.map((item: string) => removeSpaceNewline(item))
  console.log('tableHeaderB', tableHeaderB);
  const 患者姓名key = tableHeaderB[2]
  const 复诊日期key2 = tableHeaderB[26]
  const 累计就诊次数key = tableHeaderB[30]

  // 表格数据转为数组对象 WorkSheet to obj[]
  const 正畸老顾客sheet_to_arr = (list: WorkSheet, sheetHeader: any[]) => {
    let sheetJson = utils.sheet_to_json(list, { header: 1 }) as any[]
    const sheetArr = sheetJson
      // 过滤表头
      .filter(item => item && Array.isArray(item) && item.length > 0)
      .map((item: any) => {
        const obj = {} as any
        sheetHeader.forEach((key: string, index: number) => {
          obj[key] = item[index]
        })
        return obj
      })
    // console.log('sheet_to_json', sheetArr);


    return sheetArr
      // .filter(item => item[客户姓名key] && item[客户姓名key] !== 客户姓名key)
      // 过滤空值
      .filter(item => item[客户姓名key] || item[患者姓名key])
      // 过滤表头
      .filter(item => item[客户姓名key] !== 客户姓名key && item[患者姓名key] !== 患者姓名key)
      .map(item => {
        // if (item[客户姓名key] === '朱昌蔚' || item[患者姓名key] === '朱昌蔚') {
        //   console.log('朱昌蔚', item, item[复诊日期key], item[复诊日期key2])
        // }
        // 剩余治疗次数为空时赋值0
        item[剩余治疗次数key] = item[剩余治疗次数key] || 0
        // 处理患者姓名的格式问题
        if (item[客户姓名key]) item[客户姓名key] = removeSpaceNewline(item[客户姓名key])
        if (item[患者姓名key]) item[患者姓名key] = removeSpaceNewline(item[患者姓名key])
        // 处理就诊号补零
        // item[就诊号key] = item[就诊号key] && String(item[就诊号key]).padStart(6,'0')
        // 处理日期
        if (item[初诊日期key]) item[初诊日期key] = formatDate(item[初诊日期key])
        if (item[复诊日期key]) item[复诊日期key] = formatDate(item[复诊日期key])
        if (item[复诊日期key2]) item[复诊日期key2] = formatDate(item[复诊日期key2])
        return item
      })

  }

  const formatDate = (str: any) => {
    // const formatTemp = 'YYYY-MM-DD'
    const invalidDate = '无效的日期数据，请复查'
    if (str) {
      if (typeof str === 'number') {
        // return Dayjs('1900/1/1').add(str - 2, 'day').format(formatTemp)
        return str
      } else if (Dayjs(str).isValid()) {
        // return Dayjs(str).format(formatTemp)
        return Dayjs(str).diff('1900', 'day') + 2
      } else {
        return invalidDate
      }
    } else {
      return invalidDate
    }
  }


  // 以【正畸老顾客_总合计人数】表生成基准数据
  const targetArr = 正畸老顾客sheet_to_arr(正畸老顾客_总合计人数, tableHeaderA)
  // console.log('targetArr', targetArr);
  // 按表格顺序合并【正畸老顾客】中的其他表数据
  sheetA.value.SheetNames.forEach((sheetName: string) => {
    if (sheetName !== 总合计人数sheetName) {
      正畸老顾客sheet_to_arr(sheetA.value.Sheets[sheetName], tableHeaderA)
        // 合并入总合计人数表格
        .forEach(item => {
          const index = targetArr.findIndex(subItem => subItem[客户姓名key] === item[客户姓名key])
          if (index > -1) {
            // 合并数据
            Object.assign(targetArr[index], {
              [复诊日期key]: item[复诊日期key] || targetArr[index][复诊日期key],
              // [剩余治疗次数key]: item[剩余治疗次数key] ? (targetArr[index][剩余治疗次数key] - item[剩余治疗次数key]) : targetArr[index][剩余治疗次数key],
              [剩余治疗次数key]: item[剩余治疗次数key] ? (Math.min(item[剩余治疗次数key], targetArr[index][剩余治疗次数key])) : targetArr[index][剩余治疗次数key]
            })

          } else {
            // 追加数据
            targetArr.push(item)
          }
        })
    }
  });
  // console.log('targetArr', targetArr);

  // ============== 处理 每月统计的患者查询excel ===========
  if (!sheetB.value) {
    dialog.error({
      title: '错误',
      content: '请选择「患者查询」文件',
    });
    return;
  }

  // 合并【患者查询】中的第一张表格
  const targetArr2 = 正畸老顾客sheet_to_arr(sheetB.value.Sheets[sheetB.value.SheetNames[0]], tableHeaderB)
  // console.log('targetArr2', targetArr2);
  targetArr2
    // 合并入总合计人数表格
    .forEach(item => {
      const index = targetArr.findIndex(subItem => subItem[客户姓名key] === item[患者姓名key])
      if (index > -1) {
        // 合并数据
        Object.assign(targetArr[index], {
          [复诊日期key]: item[复诊日期key2] || targetArr[index][复诊日期key],
          [剩余治疗次数key]: item[累计就诊次数key] ? (targetArr[index][剩余治疗次数key] - item[累计就诊次数key]) : targetArr[index][剩余治疗次数key]
        })
      } else {
        // 追加数据
        // ['序号', '客户姓名', '就诊号', '开单医生', '初诊日期', '复诊日期', '剩余治疗次数', '待收费金额（元）', '备注']
        targetArr.push({
          '序号': '-',
          '客户姓名': item[tableHeaderB[2]],
          '就诊号': item[tableHeaderB[1]],
          '开单医生': item[tableHeaderB[12]],
          '初诊日期': item[tableHeaderB[24]],
          '复诊日期': item[tableHeaderB[26]],
          '剩余治疗次数': 0 - item[tableHeaderB[30]],
          '待收费金额（元）': '',
          '备注': item[tableHeaderB[33]],

        })
      }
    })
  // console.log('targetArr', targetArr);

  // 创建并下载新的表格
  sheetNew.value = utils.book_new();
  const targetSheet = utils.json_to_sheet(targetArr)
  // console.log('targetSheet', targetSheet)
  const targetSheetName = Dayjs().format('YYYY-MM-DD')
  utils.book_append_sheet(sheetNew.value, targetSheet, targetSheetName);
  writeFileXLSX(sheetNew.value, targetSheetName + '.xlsx');

};
</script>

<template>
  <n-space vertical style="padding: 20px; max-width: 800px; margin: auto">
    <n-h1>正畸老顾客表格自动化表格处理</n-h1>
    <n-text type="success">所有文件操作仅存在客户端本地，不会上传任何文件、信息。</n-text>
    <n-form ref="formRef" :show-feedback="false">
      <n-form-item label="第一步 选择正畸老顾客excel">
        <!-- <n-input v-model:value="model.age" @keydown.enter.prevent /> -->
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
      </n-form-item>
      <n-form-item label="第二步 选择每月统计的患者查询excel">
        <!-- <n-input v-model:value="model.age" @keydown.enter.prevent /> -->
        <n-upload multiple directory-dnd :max="1" :default-upload="false" :on-update:file-list="onUpdateFileB">
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
        <n-button type="primary" block @click="handleNewSheet">处理文件</n-button>
      </n-form-item>
    </n-form>
  </n-space>
</template>

<style scoped></style>
