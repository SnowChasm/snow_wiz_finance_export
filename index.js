// 从 origin 文件夹下读取所有的 excel 文件
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const processData = require('./processData');

const originFolder = path.join(__dirname, 'origin');
const resultFolder = path.join(__dirname, 'result');

fs.readdir(originFolder, (err, files) => {
  if (err) {
    console.error('无法读取 origin 文件夹:', err);
    return;
  }

  // 清空 result 文件夹
  fs.readdir(resultFolder, (err, resultFiles) => {
    if (err) {
      console.error('无法读取 result 文件夹:', err);
      return;
    }
    let deleteCount = 0;
    resultFiles.forEach((resultFile) => {
      const resultFilePath = path.join(resultFolder, resultFile);
      fs.unlink(resultFilePath, (err) => {
        if (err) {
          console.error('无法删除文件:', resultFilePath, err);
        } else {
          deleteCount++;
          if (deleteCount === resultFiles.length) {
            processFiles(files);
          }
        }
      });
    });
    if (resultFiles.length === 0) {
      processFiles(files);
    }
  });
});

function processFiles(files) {
  files.forEach((file) => {
    const filePath = path.join(originFolder, file);
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const excelData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    const processedWorkbook = processData(excelData);

    // 设置单元格格式为数字格式
    Object.keys(processedWorkbook.Sheets).forEach((sheetName) => {
      const sheet = processedWorkbook.Sheets[sheetName];
      Object.keys(sheet).forEach((cell) => {
        if (typeof sheet[cell].v === 'number') {
          sheet[cell].t = 'n'; // 设置单元格类型为数字
        }
      });
    });

    const resultFileName = `result_${file}`;
    const resultFilePath = path.join(resultFolder, resultFileName);
    xlsx.writeFile(processedWorkbook, resultFilePath);
  });

  console.log('所有 Excel 文件处理完成');
}
