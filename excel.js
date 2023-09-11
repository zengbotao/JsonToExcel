const XLSX = require('xlsx');

// 读取 Excel 文件
const workbook = XLSX.readFile('data.xlsx');

// 获取第一个工作表
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

// 将数据存入工作表的两列
let data = require('./num6').data2
// console.log(data)
// 将键和值分别存入两列
const keys = Object.keys(data);
const values = Object.values(data);

// 寻找工作表的最后一行
let lastRow = 0;
for (const cell in worksheet) {
  if (cell.startsWith('A') && !isNaN(parseInt(cell.slice(1)))) {
    const row = parseInt(cell.slice(1));
    if (row > lastRow) {
      lastRow = row;
    }
  }
}


// 追加数据到下一行
const newRow = lastRow + 1;

// 将键存入第一列
const firstColumn = 'A';
for (let i = 0; i < keys.length; i++) {
  const cellAddress = firstColumn + (newRow + i);
  worksheet[cellAddress] = { t: 's', v: keys[i] };
}

// 将值存入第二列
const secondColumn = 'B';
for (let i = 0; i < values.length; i++) {
  const cellAddress = secondColumn + (newRow + i);
  worksheet[cellAddress] = { t: 's', v: values[i] };
}

// 更新工作表的范围
const range = XLSX.utils.decode_range(worksheet['!ref']);
range.e.r = Math.max(range.e.r, newRow + keys.length - 1);
worksheet['!ref'] = XLSX.utils.encode_range(range);

// 保存修改后的工作簿为 Excel 文件
XLSX.writeFile(workbook, 'data.xlsx');