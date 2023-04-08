<template>
  <div class="test">
    <input type="file" ref="fileRef" />
    <button @click="read">读取excel数据</button>
  </div>
</template>

<script>
import Excel from 'exceljs'
import dayjs from 'dayjs'
/**
 *  需求:读取excel表中的每行数据
 */
export default {
  methods: {
    read() {
      const workbook = new Excel.Workbook()

      // 这种方式不行   会用到 Node.js 的专属 API，所以在浏览器上不适用,
      // https://github.com/Dream4ever/Knowledge-Base/issues/142
      // workbook.xlsx.readFile('../a.xlsx').then(() => {
      //   console.log(1)
      // })

      // 浏览器解析excel需要使用到xlsx.load方式
      const file = this.$refs.fileRef.files[0]
      // 1.将File类型转换成buffer
      const reader = new FileReader()
      reader.readAsArrayBuffer(file)
      reader.onloadend = (e) => {
        const buffer = e.target.result

        // 2.解析buffer
        workbook.xlsx.load(buffer).then((res) => {
          // console.log(res)
          const worksheet = res.getWorksheet(1)
          // sheet名称
          const sheetName = worksheet.name
          worksheet.eachRow((row, rowNumber) => {
            row.eachCell((cell, colNumber) => {
              let value = ''
              // 判断单元格的类型
              //    6-公式 ;2-数值；3-字符串；4-时间
              if (cell.type == 6) {
                value = cell.result
              } else if (cell.type == 4) {
                // 对时间类型的数据要单独格式化处理，如果需要修改返回的格则需要修改此处
                value = dayjs(cell.value).format('YYYY-MM-DD')
              } else {
                value = cell.value
              }

              console.log(
                `当前为第${rowNumber}行,第${colNumber}列,单元格类型:${cell.type},值为：${value}`
              )
            })
          })
        })
      }
    },
  },
}
</script>

<style scoped></style>
