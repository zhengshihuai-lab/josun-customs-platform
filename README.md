# JOSUN 關務工作平台

喬山永續科技集團 — 關務作業管理平台 v3.0

12 個作業模組 × 54 個步驟 × 端到端自動化

## 快速啟動

```bash
python3 關務工作平台.py
# 開啟瀏覽器: http://localhost:8899
```

## 雲端部署 (Render.com)

1. 將此 repo 推送到 GitHub
2. 前往 [Render Dashboard](https://dashboard.render.com)
3. 點擊 "New +" → "Web Service"
4. 連接 GitHub repo
5. 設定:
   - **Runtime**: Docker
   - **Plan**: Free
6. 點擊 "Create Web Service"

部署完成後會自動產生公開 URL，關務同事直接打開即可操作。

## 模組清單

| # | 模組 | 步驟數 |
|---|------|--------|
| 1 | 📦 出貨準備 | 3 |
| 2 | 📐 包裝與標識管理 | 8 |
| 3 | 📋 報關資料準備 | 6 |
| 4 | 📄 單據製作與審核 | 4 |
| 5 | 🚢 訂艙與物流 | 5 |
| 6 | 💰 稅費核算 | 4 |
| 7 | 💻 ERP錄入 | 6 |
| 8 | 🇮🇩 印尼協同 | 5 |
| 9 | 🗃️ 資料歸檔 | 3 |
| 10 | ® 品牌授權 | 3 |
| 11 | 📡 政策監控 | 3 |
| 12 | 📊 每日運營 | 4 |

## API

| Method | URL | 說明 |
|--------|-----|------|
| GET | / | 首頁 |
| GET | /api/tickets | 列出工作票 |
| POST | /api/ticket/create | 建立工作票 |
| GET | /api/status | 系統狀態 |
| GET | /api/execute?tool=X&args=Y | 執行工具 |

## License

Internal use only - JOSUN Group
