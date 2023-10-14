const express = require("express");
const cors = require("cors");

const app = express();

// 設定CORS
app.use(cors());

// 定義您的API路由和邏輯

// 監聽端口
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log('process=', process);
  console.log('process.env=', process.env);
  console.log(`Server is running on port ${PORT}`);
});
