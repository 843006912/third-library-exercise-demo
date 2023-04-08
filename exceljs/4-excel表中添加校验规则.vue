<template>
  <div class="test">
    <button @click="download">点击下载excel</button>
  </div>
</template>

<script>
import Excel from 'exceljs'
import FileSaver from 'file-saver'

/**
 * 需求：为指定的列添加校验以及注释，比如下拉选择内容
 *
 */
export default {
  methods: {
    async download() {
      const workbook = new Excel.Workbook()

      workbook.creator = 'test123'
      workbook.created = new Date()
      const worksheet = workbook.addWorksheet('汇总')

      // 添加数据
      // worksheet.addRow(['姓名', '年龄', '地址'])
      // worksheet.addRow(['zs', 30, 'dg'])
      // worksheet.addRow(['ls', 17, 'gz'])

      // 指定列名以及列属性
      worksheet.columns = [
        {
          header: '序号',
          key: 'index',
        },
        {
          header: '姓名',
          key: 'name',
        },
        { header: '年龄', key: 'age' },
        { header: '国家', key: 'country' },
        { header: '生日', key: 'birth' },
      ]

      // 为某个单元格设置校验规则
      // worksheet.getCell('E2').dataValidation = {
      //   type: 'date',
      //   operator: 'greaterThan',
      //   showErrorMessage: true,
      //   formulae: [new Date('2023-03-03')],
      //   errorStyle: 'error',
      //   errorTitle: '提示',
      //   error: '输入的值必须大于2023-02-03',
      // }

      // 添加注释
      // worksheet.getCell('E1').note = '输入时间大于2023-02-03!'
      //   有样式的注释
      worksheet.getCell('E1').note = {
        texts: [
          {
            font: {
              size: 12,
              color: { theme: 0 },
              name: 'Calibri',
              family: 2,
              scheme: 'minor',
            },
            text: 'This is ',
          },
          {
            font: {
              italic: true,
              size: 12,
              color: { theme: 0 },
              name: 'Calibri',
              scheme: 'minor',
            },
            text: 'a',
          },
          {
            font: {
              size: 12,
              color: { theme: 1 },
              name: 'Calibri',
              family: 2,
              scheme: 'minor',
            },
            text: ' ',
          },
          {
            font: {
              size: 12,
              color: { argb: 'FFFF6600' },
              name: 'Calibri',
              scheme: 'minor',
            },
            text: 'colorful',
          },
        ],
      }

      // 对第2-100行的第5列添加校验
      //     备注：没有找到对表格中第5列的所有数据校验方法
      for (let row = 2; row <= 100; row++) {
        worksheet.getCell(row, 5).dataValidation = {
          type: 'date',
          operator: 'greaterThan',
          showErrorMessage: true,
          formulae: [new Date('2023-03-03')],
          errorStyle: 'error',
          errorTitle: '提示',
          error: '输入的值必须大于2023-02-03',
        }
      }

      const EXCEL_TYPE =
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8'

      workbook.xlsx.writeBuffer().then((data) => {
        const blob = new Blob([data], { type: EXCEL_TYPE })

        FileSaver.saveAs(blob, 'download.xlsx')
      })
    },
  },
}
</script>

<style scoped></style>
