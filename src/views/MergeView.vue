<script setup lang="ts">
import { ref } from 'vue'
import { genFileId } from 'element-plus'
import type { UploadUserFile, UploadInstance, UploadProps, UploadRawFile } from 'element-plus'
import { compareSheet,  writeToFile } from './marge'

const fileList = ref<UploadUserFile[]>([])
const upload = ref<UploadInstance>()

const newItems = ref<string[]>([])
const onlineItems = ref<string[]>([])

const handleExceed: UploadProps['onExceed'] = (files) => {
  upload.value!.clearFiles()
  const file = files[0] as UploadRawFile
  file.uid = genFileId()
  upload.value!.handleStart(file)
}

const submitUpload = () => {
  compareSheet(fileList.value[0].raw).then((res) => {
    newItems.value = res.newItems
    onlineItems.value = res.onlineItems

    writeToFile(res.mergedRowsReason)
  })
}
</script>

<template>
  <div>
    <el-upload
      v-model:file-list="fileList"
      ref="upload"
      class="upload-demo"
      action="https://run.mocky.io/v3/9d059bf9-4660-45f2-925d-ce80ad6c4d15"
      :limit="1"
      :on-exceed="handleExceed"
      :auto-upload="false"
    >
      <template #trigger>
        <el-button type="primary">select file</el-button>
      </template>
      <el-button class="ml-3" type="success" @click="submitUpload"> OK </el-button>
      <template #tip>
        <div class="el-upload__tip text-red">limit 1 file, new file will cover the old file</div>
      </template>
    </el-upload>

    <section>
      新增离线监控
      <ul class="newItems">
        <li v-for="item in newItems" :key="item">{{ item }}</li>
      </ul>
    </section>
    <section>
      上线监控
      <ul class="onlineItems">
        <li v-for="item in onlineItems" :key="item">{{ item }}</li>
      </ul>
    </section>
  </div>
</template>

<style scoped></style>

//
