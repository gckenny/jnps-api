const express = require("express");
const cors = require("cors");
const { serialToDate } = require("./utils");
const { getSheet } = require("./spreadsheet");
const bodyParser = require("body-parser");
const moment = require("moment-timezone");

// 引入dotenv以讀取.env文件中的環境變數
require("dotenv").config();

const app = express();

// 設定CORS
app.use(cors());
app.use(bodyParser.json()); // to support JSON-encoded bodies

// 定義路由
app.get("/getMembers", async (req, res) => {
  try {
    const sheet = await getSheet("選手名單");
    await sheet.loadCells("A1:R" + sheet.rowCount);

    const result = [];

    for (let rowIndex = 1; rowIndex < sheet.rowCount; rowIndex++) {
      const nameCell = sheet.getCell(rowIndex, 1); // B欄的單元格
      if (nameCell.value) {
        const rowData = {};
        for (let colIndex = 0; colIndex <= 17; colIndex++) {
          // A-R 欄
          const cell = sheet.getCell(rowIndex, colIndex);
          let cellValue = cell.value;
          if (colIndex === 7 && rowIndex > 0) {
            cellValue = serialToDate(sheet.getCell(rowIndex, colIndex).value);
          }
          rowData[String.fromCharCode(65 + colIndex)] = cellValue; // 將列索引轉換為字母 (A-R)
        }
        result.push(rowData);
      }
    }

    // 輸出結果並結束請求
    res.json({ data: result });
  } catch (error) {
    console.error("Error fetching data: ", error);
    res.status(500).json({ error: "Internal Server Error", message: error.message });
  }
});

// 定義新的路由 /getMember/{guid}
app.get("/getMember/:guid", async (req, res) => {
  try {
    const guid = req.params.guid;
    const sheet = await getSheet("選手名單");

    // 加載 A-R 欄的單元格
    await sheet.loadCells("A1:R" + sheet.rowCount);

    for (let rowIndex = 0; rowIndex < sheet.rowCount; rowIndex++) {
      const guidCell = sheet.getCell(rowIndex, 17); // 比對R欄
      if (guidCell.value === guid) {
        const rowData = {};
        for (let colIndex = 0; colIndex <= 17; colIndex++) {
          // A-R 欄
          const cell = sheet.getCell(rowIndex, colIndex);
          let cellValue = cell.value;
          if (colIndex === 7) {
            cellValue = serialToDate(sheet.getCell(rowIndex, colIndex).value);
          }
          rowData[String.fromCharCode(65 + colIndex)] = cellValue; // 將列索引轉換為字母 (A-R)
        }
        return res.json({ data: rowData }); // 如果找到了匹配的 guid，則返回該行的資料並結束請求
      }
    }

    // 如果循環結束而未找到匹配的 guid，則返回404錯誤
    res.status(404).json({ error: "Member not found" });
  } catch (error) {
    console.error("Error fetching data: ", error);
    res.status(500).json({ error: "Internal Server Error", message: error.message });
  }
});

app.post("/setCheckin", async (req, res) => {
  // debug
  // req = {
  //   body: {
  //     date: "2023/11/01",
  //     name: "黃予慈",
  //     gender: "女",
  //     classNumber: "207",
  //     group: "A",
  //     checkinTime: "2023/11/01 06:15:05",
  //     status: 0,
  //     remarks: "這是備註",
  //     guid: "23e4a9ad-4bfe-4c38-9fb7-e4c6c01bb86b",
  //   },
  // };

  try {
    const { date, name, gender, classNumber, group, checkinTime, status, remarks, guid } = req.body;
    const sheet = await getSheet("DEV");
    const statusMap = { 0: "出席", 1: "缺席", 2: "遲到", 3: "請假" };
    const rows = await sheet.getRows(); // Load all rows from the sheet
    // Find a row with the same date and name
    const twTimestamp = moment().tz("Asia/Taipei").format();
    const existingRow = rows.find((row) => row._rawData[0] === date && row._rawData[8] === guid);
    if (existingRow) {
      // Update the existing row if found
      existingRow._rawData[5] = checkinTime;
      existingRow._rawData[6] = statusMap[status];
      existingRow._rawData[7] = remarks;
      existingRow._rawData[9] = twTimestamp;
      await existingRow.save(); // Save the updates to the sheet
      console.log(existingRow);
      res.json({ success: true, message: "Check-in data updated successfully" });
    } else {
      // Add a new row if no matching row was found
      const newRow = {
        日期: date,
        姓名: name,
        性別: gender,
        班級: classNumber,
        組別: group,
        簽到時間: checkinTime,
        狀態: statusMap[status],
        備註: remarks,
        GUID: guid,
        Timestamp: twTimestamp,
      };
      await sheet.addRow(newRow);
      res.json({ success: true, message: "Check-in data added successfully" });
    }
  } catch (error) {
    console.error("Error writing data: ", error);
    res.status(500).json({ error: "Internal Server Error", message: error.message });
  }
});

app.get("/getCheckins/:date", async (req, res) => {
  try {
    let { date } = req.params; // 從 URL 中獲取日期參數
    // 格式化日期
    date = date.slice(0, 4) + "/" + date.slice(4, 6) + "/" + date.slice(6, 8);

    const sheet = await getSheet("DEV"); // 使用之前定義的 getSheet 函數來獲取工作表
    const rows = await sheet.getRows(); // 使用 getRows 方法來獲取所有行

    // 使用 filter 方法來過濾符合指定日期的行
    const filteredRows = rows.filter((row) => row._rawData[0] === date);

    // 將過濾後的行數據格式化為 JSON 對象
    const result = filteredRows.map((row) => ({
      date: row._rawData[0],
      name: row._rawData[1],
      gender: row._rawData[2],
      classNumber: row._rawData[3],
      group: row._rawData[4],
      checkinTime: row._rawData[5],
      status: row._rawData[6],
      remarks: row._rawData[7],
      guid: row._rawData[8],
      writeTime: row._rawData[9],
    }));

    // 輸出結果並結束請求
    res.json({ data: result });
  } catch (error) {
    console.error("Error fetching data: ", error);
    res.status(500).json({ error: "Internal Server Error", message: error.message });
  }
});

// 監聽端口
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
