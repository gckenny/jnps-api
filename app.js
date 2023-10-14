const express = require("express");
const cors = require("cors");
const { GoogleSpreadsheet } = require('google-spreadsheet'); // 引入google-spreadsheet模塊

// 設定CORS
app.use(cors());

// 設置dotenv環境變數
require('dotenv').config();

const app = express();

app.get("/getMembers", async (req, res) => {
  try {
    // 使用google-spreadsheet模塊來訪問Google Sheets
    const doc = new GoogleSpreadsheet(process.env.SHEET_ID);

    // 使用Google Sheets API設置訪問令牌（如果需要的話）
    doc.useServiceAccountAuth({
      client_email: process.env.CLIENT_EMAIL,
      private_key: process.env.PRIVATE_KEY,
    });

    // 加載文檔信息並讀取表格
    await doc.loadInfo(); // 加載文檔信息
    const sheet = doc.sheetsByTitle['選手名單']; // 假設工作表名稱為選手名單

    // 讀取工作表的數據
    const rows = await sheet.getRows();

    // 將數據作為JSON返回
    res.json(rows);
  } catch (error) {
    console.error("Error fetching data: ", error);
    res.status(500).json({ error: "An error occurred while fetching data" });
  }
});

// 監聽端口
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
