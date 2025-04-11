const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const dayjs = require("dayjs");

const originFolder = path.join(__dirname, "origin");
const resultFolder = path.join(__dirname, "result");

const taxRate = 0.01;

// 检查 origin 文件夹是否存在，如果不存在则创建
if (!fs.existsSync(originFolder)) {
  fs.mkdirSync(originFolder);
}

// 检查 result 文件夹是否存在，如果不存在则创建
if (!fs.existsSync(resultFolder)) {
  fs.mkdirSync(resultFolder);
}

fs.readdir(originFolder, (err, files) => {
  if (err) {
    console.error("无法读取 origin 文件夹:", err);
    return;
  }

  fs.readdir(resultFolder, (err, resultFiles) => {
    if (err) {
      console.error("无法读取 result 文件夹:", err);
      return;
    }
    let deleteCount = 0;
    resultFiles.forEach((resultFile) => {
      const resultFilePath = path.join(resultFolder, resultFile);
      fs.unlink(resultFilePath, (err) => {
        if (err) {
          console.error("无法删除文件:", resultFilePath, err);
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

function sanitizeSheetName(name) {
  return name.replace(/[:\/\\?*\[\]]/g, "_");
}

function processFiles(files) {
  files.forEach((file) => {
    const filePath = path.join(originFolder, file);
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const excelData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    const groupedData = {};

    excelData.forEach((row) => {
      const category = row["类别"] || "未知类别";
      const channel = row["支付渠道"] || "未知渠道";
      const sheetName = sanitizeSheetName(`${category}_${channel}`);
      if (!groupedData[sheetName]) {
        groupedData[sheetName] = [];
      }
      groupedData[sheetName].push(row);
    });

    const result = {};

    Object.keys(groupedData).forEach((sheetName) => {
      const sheetData = groupedData[sheetName];

      sheetData.forEach((row) => {
        const startDate = dayjs(row["有效起始时间"]).startOf("day");
        const endDate = dayjs(row["有效到期时间"]).endOf("day");
        const daysDiff = endDate.diff(startDate, "day") + 1;

        row["总服务器天数"] = daysDiff;
        row["应交税费"] = (row["实收金额"] / (1 + taxRate)) * taxRate;
        row["税后金额"] = row["实收金额"] - row["应交税费"];
        row["DRR"] = row["税后金额"] / daysDiff;

        row["实收金额"] = parseFloat(row["实收金额"]);
        row["应交税费"] = parseFloat(row["应交税费"]);
        row["税后金额"] = parseFloat(row["税后金额"]);
        row["总服务器天数"] = parseFloat(row["总服务器天数"]);
        row["DRR"] = parseFloat(row["DRR"]);
      });

      const allDates = sheetData
        .map((row) => [
          dayjs(row["有效起始时间"]).startOf("day"),
          dayjs(row["有效到期时间"]).endOf("day"),
        ])
        .flat();
      const minDate = allDates.reduce(
        (min, date) => (date.isBefore(min) ? date : min),
        allDates[0]
      );
      const maxDate = allDates.reduce(
        (max, date) => (date.isAfter(max) ? date : max),
        allDates[0]
      );

      console.log(`${sheetName} 开始处理, 数据量 ${sheetData.length}`);
      for (let year = minDate.year(); year <= maxDate.year(); year++) {
        for (let month = 0; month < 12; month++) {
          const monthStart = dayjs(new Date(year, month, 1)).startOf("day");
          const monthEnd = monthStart.endOf("month").endOf("day");
          let daysInMonth = monthEnd.diff(monthStart, "day") + 1;

          sheetData.forEach((row) => {
            const startDate = dayjs(row["有效起始时间"]).startOf("day");
            const endDate = dayjs(row["有效到期时间"]).endOf("day");
            const incomeKey = `${year}年${month + 1}月收入`;
            if (!row[incomeKey]) {
              row[incomeKey] = 0;
            }
            if (startDate.isBefore(monthEnd) && endDate.isAfter(monthStart)) {
              let actualDaysInMonth = daysInMonth;
              if (
                startDate.isAfter(monthStart) &&
                startDate.isBefore(monthEnd)
              ) {
                actualDaysInMonth = endDate.isBefore(monthEnd)
                  ? endDate.diff(startDate, "day") + 1
                  : monthEnd.diff(startDate, "day") + 1;
              } else if (
                endDate.isAfter(monthStart) &&
                endDate.isBefore(monthEnd)
              ) {
                actualDaysInMonth = endDate.diff(monthStart, "day") + 1;
              }
              if (startDate.isSame(monthEnd, "day")) {
                actualDaysInMonth = 1;
              }
              row[incomeKey] += parseFloat(row["DRR"] * actualDaysInMonth);
            }
          });
        }
      }
      console.log(`${sheetName} 处理完成`);
      result[sheetName] = sheetData;
    });

    try {
      console.log("开始导出");
      Object.keys(result).forEach((sheetName, index) => {
        const newWorkbook = xlsx.utils.book_new();
        const chunkSize = 1000; // 每批次数据量
        const data = result[sheetName];
        for (let i = 0; i < data.length; i += chunkSize) {
          console.log(`正在处理 ${sheetName} 的第 ${i / chunkSize + 1} 批次`);
          const chunk = data.slice(i, i + chunkSize);
          const newSheet = xlsx.utils.json_to_sheet(chunk);
          xlsx.utils.book_append_sheet(
            newWorkbook,
            newSheet,
            `${sheetName}_part${i / chunkSize + 1}`
          );
        }
        const originalFileName = path.basename(file, path.extname(file));
        const resultFileName = `${originalFileName}_${sheetName}_${index}.xlsx`;
        const resultFilePath = path.join(resultFolder, resultFileName);
        xlsx.writeFile(newWorkbook, resultFilePath);
      });
    } catch (error) {
      console.log("导出失败", error);
    }
  });

  console.log("所有 Excel 文件处理完成");
}
