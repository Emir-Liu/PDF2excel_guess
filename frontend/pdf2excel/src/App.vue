<template>
  PDF2Excel
  <br>
  <br>
  
  <div>
    <input type="file" v-on:change="onFileChange"/>
  </div>
</template>

<script lang="ts" setup>
  import axios from 'axios';

  async function onFileChange(event: Event) {
    console.log('event:',event)
    const element = event.target as HTMLInputElement;
    console.log('element:',element)
    const formData = new FormData();
    formData.append('file', element.files[0]);
    try {
      const response = await axios.post('http://172.19.3.206:12309/uploadpdf/', formData, {
        headers: {
          'Content-Type': 'multipart/form-data'
        },
        responseType: 'blob' // 确保响应类型为 blob
      });

      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      const contentDisposition = response.headers['content-disposition']
      const filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
      let fileName = 'sheet.xlsx'
      const matches = filenameRegex.exec(contentDisposition);
      if (matches != null && matches[1]) {
        fileName = matches[1].replace(/['"]/g, ''); // 移除引号
      }
      // return 'default_filename';
      // link.setAttribute('download', element.name.replace(" ", "_") + '.xlsx');
      link.setAttribute('download', fileName);

      document.body.appendChild(link);
      link.click();
      link.parentNode.removeChild(link);
      window.URL.revokeObjectURL(url);
    } catch (error) {
      console.error('上传失败:', error);
    }
  }
</script>