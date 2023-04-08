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
 *  需求:读取excel表指定列的数据
 */
export default {
  methods: {
    read() {
      const workbook = new Excel.Workbook()

      // 1.设置excel的表头
      const excelColumns = [
        { header: '姓名', key: 'name', width: 10 },
        { header: '年龄', key: 'age', width: 32 },
        { header: '生日', key: 'birth', width: 10 },
        { header: '住址', key: 'address', width: 10 },
      ]

      // 2. 指定需要读取哪些列
      const needColumns = ['年龄', '姓名', '住址', '国家']

      const result = []

      function getKey(columnName) {
        const obj = excelColumns.find((item) => item.header == columnName)
        if (obj) {
          return obj.key
        } else {
          return ''
        }
      }

      const file = this.$refs.fileRef.files[0]
      const reader = new FileReader()
      reader.readAsArrayBuffer(file)
      reader.onloadend = (e) => {
        const buffer = e.target.result

        workbook.xlsx.load(buffer).then((res) => {
          const worksheet = res.getWorksheet(1)
          // 默认是没有列名的，可以手动设置列名
          worksheet.columns = excelColumns

          worksheet.eachRow(function (row, rowNumber) {
            if (rowNumber != 1) {
              let obj = {}
              for (let i = 0; i < needColumns.length; i++) {
                const columnName = needColumns[i]
                const columnKey = getKey(columnName)
                try {
                  const cell = row.getCell(columnKey)
                  let value = ''
                  if (cell.type == 6) {
                    value = cell.result
                  } else if (cell.type == 4) {
                    value = dayjs(cell.value).format('YYYY-MM-DD')
                  } else {
                    value = cell.value
                  }
                  obj[columnKey] = value
                } catch (err) {
                  console.error(
                    `excel中不存在列名为:${columnName},请在"excelColumns"中添加`
                  )
                }
              }
              result.push(obj)
            }
          })
          console.log(result)

          // console.log(worksheet.columns)
          // const sheetName = worksheet.name
          // worksheet.eachRow((row, rowNumber) => {
          //   console.log(row)
          // })
        })
      }
    },
  },
}
</script>

<style scoped></style>
