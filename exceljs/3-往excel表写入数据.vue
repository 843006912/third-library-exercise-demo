<template>
  <div class="test">
    <button @click="download">点击下载excel</button>
  </div>
</template>

<script>
import Excel from 'exceljs'
import FileSaver from 'file-saver'

/**
 *
 * 写入数据，并为列设置样式
 */
export default {
  methods: {
    async download() {
      const workbook = new Excel.Workbook()

      // 不适用
      // workbook.xlsx.writeFile('aaa.xlsx')

      workbook.creator = 'test123'
      workbook.created = new Date()
      // 设置sheet的名称
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
          // width: 30,
          style: {
            alignment: {
              horizontal: 'center',
              vertical: 'middle',
            },
            font: {
              name: '黑体',
              size: 20,
            },
            // 设置边框样式
            // border: {
            //   top: {
            //     style: 'double',
            //     color: {
            //       argb: 'FF00FF00',
            //     },
            //   },
            // },
          },
        },
        {
          header: '姓名',
          key: 'name',
          // style: {
          //   fill: {
          //     // 设置渐变色
          //     type: 'gradient',
          //     pattern: 'solid',
          //     stops: [
          //       { position: 0, color: { argb: 'FF0000' } },
          //       { position: 1, color: { argb: '00FF00' } },
          //     ],
          //   },
          // },
        },
        { header: '年龄', key: 'age' },
        { header: '国家', key: 'country' },
        { header: '生日', key: 'birth' },
      ]

      // 添加数据
      const data = [
        { index: 1, name: 'zs', age: 30, country: '中国', birth: '2023-04-08' },
        { index: 2, name: 'll', age: 25, country: '日本', birth: '2020-04-08' },
      ]
      worksheet.addRows(data)

      const EXCEL_TYPE =
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8'

      // excelJs 官方文档上面的 写入文件方法只能通过node服务端使用，浏览器却不能使用writeFile()但是可以使用writeBuffer()；
      //    需要通过 workbook.xlsx.writeBuffer() 将 buffer转为blob 配合 file-saver导出文件
      workbook.xlsx.writeBuffer().then((data) => {
        const blob = new Blob([data], { type: EXCEL_TYPE })

        FileSaver.saveAs(blob, 'download.xlsx')
      })
    },
  },
}
</script>

<style scoped></style>
