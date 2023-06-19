<script setup lang="ts">
import { reactive, ref } from 'vue'
import {
  FormInstance,
  FormRules,
  TableColumnCtx,
  UploadFile,
  genFileId,
} from 'element-plus'
import type { UploadInstance, UploadProps, UploadRawFile } from 'element-plus'
import { utils, read } from 'xlsx'
import * as dayjs from 'dayjs'

const excelAccept =
  '.csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel'
const timeFormat = 'HH:mm'

const formRef = ref<FormInstance>()
const loading = ref(false)
const form = reactive({
  dateField: '日期',
  outTimeField: '签退时间',
  nameField: '姓名',
  name: localStorage.getItem('name') || '',
  startTime: '19:30',
})

const upload = ref<UploadInstance>()
const rows = ref([] as any[])
const filterRows = ref([] as any[])

interface SummaryMethodProps<T = any> {
  columns: TableColumnCtx<T>[]
  data: T[]
}

const getSummaries = (param: SummaryMethodProps) => {
  const { columns, data } = param
  const sums: string[] = []
  columns.forEach((column, index) => {
    if (index === 0) {
      sums[index] = '时长合计(h)'
      return
    }
    const values = data.map((item) => Number(item[column.property]))
    if (!values.every((value) => Number.isNaN(value))) {
      sums[index] = values
        .reduce((prev, curr) => {
          const value = Number(curr)
          if (!Number.isNaN(value)) {
            return prev + curr
          } else {
            return prev
          }
        }, 0)
        .toFixed(2)
    } else {
      sums[index] = ''
    }
  })

  return sums
}

const rules: FormRules = {
  name: [{ required: true, message: '必填' }],
  nameField: [{ required: true, message: '必填' }],
  outTimeField: [{ required: true, message: '必填' }],
  startTime: [{ required: true, message: '请选择' }],
  excel: [
    {
      required: true,
      validator(_, __, callback) {
        if (!rows.value.length) {
          callback('请先上传Excel / 解析的数据为空')
        } else {
          callback()
        }
      },
    },
  ],
}

const handleExceed: UploadProps['onExceed'] = (files) => {
  upload.value!.clearFiles()
  const file = files[0] as UploadRawFile
  file.uid = genFileId()
  upload.value!.handleStart(file)
}

const parseExcel = () => {
  formRef.value?.validate((isValid) => {
    if (!isValid) {
      return
    }

    filterRows.value = rows.value
      .filter((item) => {
        if (
          item[form.nameField] === form.name &&
          dayjs(item[form.outTimeField], timeFormat).isAfter(
            dayjs(form.startTime, timeFormat),
          )
        ) {
          return true
        }
        return false
      })
      .map((item) => {
        return {
          ...item,
          overtime: dayjs(item[form.outTimeField], timeFormat).diff(
            dayjs(form.startTime, timeFormat),
            'hour',
            true,
          ),
        }
      })

    localStorage.setItem('name', form.name)
  })
}

const onChange = (uploadFile: UploadFile) => {
  loading.value = true

  const fileReader = new FileReader()
  fileReader.onload = (ev) => {
    try {
      const data = ev.target!.result
      const workbook = read(data, {
        type: 'binary',
        dense: true,
      })
      rows.value = utils.sheet_to_json(
        workbook.Sheets[workbook.SheetNames[0]],
        { raw: false },
      )
    } catch (err) {
      console.error(err)
    } finally {
      loading.value = false
    }
  }

  // 以二进制方式打开文件
  fileReader.readAsBinaryString(uploadFile.raw!)
}
</script>

<template>
  <el-form
    :model="form"
    :rules="rules"
    ref="formRef"
    label-width="auto"
    v-loading="loading"
  >
    <el-form-item prop="nameField" label="姓名字段">
      <el-input v-model="form.nameField"></el-input>
    </el-form-item>

    <el-form-item prop="outTimeField" label="签退时间字段">
      <el-input v-model="form.outTimeField"></el-input>
    </el-form-item>

    <el-form-item prop="startTime" label="开始加班时间">
      <el-time-picker
        value-format="HH:mm"
        v-model="form.startTime"
      ></el-time-picker>
    </el-form-item>

    <el-form-item prop="name" label="姓名">
      <el-input v-model="form.name"></el-input>
    </el-form-item>

    <el-form-item prop="excel" label="考勤表">
      <div class="upload-wrap">
        <el-upload
          ref="upload"
          :accept="excelAccept"
          :limit="1"
          :on-exceed="handleExceed"
          :auto-upload="false"
          :on-change="onChange"
        >
          <template #trigger>
            <el-button type="primary">请上传</el-button>
          </template>
        </el-upload>
      </div>
    </el-form-item>

    <el-form-item label=" ">
      <el-button type="success" @click="parseExcel"> 开始解析 </el-button>
    </el-form-item>
  </el-form>

  <el-table :data="filterRows" :summary-method="getSummaries" show-summary>
    <el-table-column label="加班日期" :prop="form.dateField"> </el-table-column>
    <el-table-column label="加班时长(h)" prop="overtime">
      <template #default="{ row }">{{ row.overtime.toFixed(2) }}</template>
    </el-table-column>
  </el-table>
</template>

<style scoped>
.upload-wrap {
  text-align: left;
  width: 100%;
}
</style>
