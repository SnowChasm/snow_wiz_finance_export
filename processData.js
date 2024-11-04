// 处理读取的 excel 数据
const xlsx = require("xlsx");
const dayjs = require("dayjs");

function sanitizeSheetName(name) {
  // 替换所有非法字符为下划线
  return name.replace(/[:\/\\?*\[\]]/g, "_");
}

function processData(excelData) {
  const groupedData = {};

  // 1. 根据 类别 + 支付渠道 分组，每个分组生成单独的 sheet, sheet 的名称叫做 类别 + 支付渠道
  excelData.forEach((row, index) => {
    const category = row["类别"] || "未知类别";
    const channel = row["支付渠道"] || "未知渠道";
    const sheetName = sanitizeSheetName(`${category}_${channel}`);
    if (!groupedData[sheetName]) {
      groupedData[sheetName] = [];
    }
    groupedData[sheetName].push(row);
  });

  const result = {};

  Object.keys(groupedData).forEach((sheetName, sheetIndex) => {
    const sheetData = groupedData[sheetName];

    sheetData.forEach((row, rowIndex) => {
      // 修正 有效起始时间 和 有效到期时间 的数据，只保留日期，去掉时间
      const startDate = dayjs(row["有效起始时间"]).startOf("day");
      const endDate = dayjs(row["有效到期时间"]).endOf("day");
      const daysDiff = endDate.diff(startDate, "day") + 1;

      // 3. 新增 总服务器天数 列，计算公式为：有效到期时间 - 有效起始时间
      row["总服务器天数"] = daysDiff;

      // 4. 新增 应交税费 列，计算公式为：实收金额 / 1.06 * 0.06
      row["应交税费"] = (row["实收金额"] / 1.06) * 0.06;

      // 4. 新增 税后金额 列，计算公式为：实收金额 - 应交税费
      row["税后金额"] = row["实收金额"] - row["应交税费"];

      // 5. 新增 DRR 列，计算公式为：税后金额 / 之前算出的天数
      row["DRR"] = row["税后金额"] / daysDiff;

      // 设置为数字格式
      row["实收金额"] = parseFloat(row["实收金额"]);
      row["应交税费"] = parseFloat(row["应交税费"]);
      row["税后金额"] = parseFloat(row["税后金额"]);
      row["总服务器天数"] = parseFloat(row["总服务器天数"]);
      row["DRR"] = parseFloat(row["DRR"]);
    });

    // 6. 根据所有数据的 有效起始时间 和 有效到期时间，算出最早和最晚的时间差，以年为单位
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

    // 7. 根据前面算出的年数及当前行的数据，新增每月确认收入，计算公式为：DRR * 当月的天数
    for (let year = minDate.year(); year <= maxDate.year(); year++) {
      for (let month = 0; month < 12; month++) {
        const monthStart = dayjs(new Date(year, month, 1)).startOf("day");
        const monthEnd = monthStart.endOf("month").endOf("day");
        let daysInMonth = monthEnd.diff(monthStart, "day") + 1;

        sheetData.forEach((row, index) => {
          const startDate = dayjs(row["有效起始时间"]).startOf("day");
          const endDate = dayjs(row["有效到期时间"]).endOf("day");
          const incomeKey = `${year}年${month + 1}月收入`;
          if (!row[incomeKey]) {
            row[incomeKey] = 0;
          }
          if (startDate.isBefore(monthEnd) && endDate.isAfter(monthStart)) {
            // 修正为真实的天数
            let actualDaysInMonth = daysInMonth;
            if (startDate.isAfter(monthStart) && startDate.isBefore(monthEnd)) {
              actualDaysInMonth = endDate.isBefore(monthEnd)
                ? endDate.diff(startDate, "day") + 1
                : monthEnd.diff(startDate, "day") + 1;
            } else if (
              endDate.isAfter(monthStart) &&
              endDate.isBefore(monthEnd)
            ) {
              actualDaysInMonth = endDate.diff(monthStart, "day") + 1;
            }
            // 修复起始日期是当月最后一天的情况
            if (startDate.isSame(monthEnd, "day")) {
              actualDaysInMonth = 1;
            }
            row[incomeKey] += parseFloat(row["DRR"] * actualDaysInMonth);
          }
        });
      }
    }

    // 新增一列，将每月收入相加后与税后金额进行比较，输出差值
    // sheetData.forEach((row) => {
    //   let totalMonthlyIncome = 0;
    //   Object.keys(row).forEach((key) => {
    //     if (key.endsWith("月收入")) {
    //       totalMonthlyIncome += row[key];
    //     }
    //   });
    //   row["收入差值"] = row["税后金额"] - totalMonthlyIncome;
    // });

    result[sheetName] = sheetData;

    // 输出处理进度
    console.log(
      `已处理 ${sheetIndex + 1} / ${Object.keys(groupedData).length} 个 sheet`
    );
  });

  // 使用 xlsx.js 导出数据
  const newWorkbook = xlsx.utils.book_new();
  Object.keys(result).forEach((sheetName) => {
    const newSheet = xlsx.utils.json_to_sheet(result[sheetName]);
    xlsx.utils.book_append_sheet(newWorkbook, newSheet, sheetName);
  });

  return newWorkbook;
}

module.exports = processData;
