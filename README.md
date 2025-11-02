# 妙華蓮華經 - 佛經閱讀系統

[![GitHub](https://img.shields.io/badge/GitHub-Repository-blue)](https://github.com/Rong-bit/dts-maitreya)
[![Live Demo](https://img.shields.io/badge/Live-Demo-green)](https://script.google.com/macros/s/AKfycbxzWiK-jRfWXS2M9AuIFG3rK0_GveXn3UvtEWhu3Yv2M3HD-ZvWSJQ4miEi5-WdOANsqQ/exec)

這是一個基於 Google Apps Script 開發的佛經閱讀與翻譯系統，提供經文閱讀、語音朗讀、同音字顯示、AI 翻譯等功能。

- **GitHub 倉庫**：https://github.com/Rong-bit/dts-maitreya
- **線上版本**：https://script.google.com/macros/s/AKfycbxzWiK-jRfWXS2M9AuIFG3rK0_GveXn3UvtEWhu3Yv2M3HD-ZvWSJQ4miEi5-WdOANsqQ/exec

## 專案結構

- `Code.gs` - Google Apps Script 後端代碼（服務器端邏輯）
- `index.html` - 前端 HTML 模板文件（用戶界面）

## 功能特色

### 核心功能
- 📖 經文閱讀與顯示
- 🔊 語音朗讀功能（支援多種語音）
- 🔤 同音字顯示功能
- 🤖 AI 翻譯功能（使用 Gemini 2.5 Flash）

### 進階功能
- 🔍 全文搜尋
- 📑 書籤管理
- 📊 閱讀進度追蹤
- 🎨 字體大小調整
- 🔁 重複播放功能

## 部署說明

### 方法一：使用 clasp（推薦）

1. 安裝 clasp：
```bash
npm install -g @google/clasp
```

2. 登入並創建項目：
```bash
clasp login
clasp create --type webapp --title "佛經閱讀系統"
```

3. 推送代碼：
```bash
clasp push
```

4. 部署：
```bash
clasp deploy
```

### 方法二：手動上傳

1. 前往 [Google Apps Script](https://script.google.com/)
2. 創建新專案
3. 複製 `Code.gs` 內容到 `Code.gs`
4. 創建 HTML 文件，命名為 `index.html`，複製 `index.html` 內容
5. 部署為網路應用程式

## 使用說明

1. 在 Google Sheets 中建立工作表，格式如下：
   - 第 1 行：品名列（奇數欄）和翻譯說明（偶數欄）
   - 第 2 行開始：經文內容

2. 設置 API 金鑰：
   - 在 GAS 編輯器中運行 `設置 API 金鑰`
   - 輸入您的 Gemini API 金鑰

3. 訪問應用：
   - **線上版本**：[直接訪問](https://script.google.com/macros/s/AKfycbxzWiK-jRfWXS2M9AuIFG3rK0_GveXn3UvtEWhu3Yv2M3HD-ZvWSJQ4miEi5-WdOANsqQ/exec)
   - 或部署後獲得 URL 並在瀏覽器中訪問

## 功能說明

### 全文搜尋
- 在搜尋框輸入關鍵字即可搜尋當前品名中的內容
- 支援即時搜尋（防抖處理）

### 書籤功能
- **電腦版**：`Shift + Ctrl + 點擊` 經文行
- **手機版**：長按經文行（0.8秒）或點擊右上角書籤按鈕
- 點擊「書籤」按鈕可查看所有書籤

### 閱讀進度
- 自動記錄：點擊任何經文行自動保存
- 恢復進度：點擊「進度」按鈕跳轉到上次閱讀位置

## 技術架構

- **前端**：HTML5 + CSS3 + JavaScript (ES6+)
- **後端**：Google Apps Script
- **AI 翻譯**：Google Gemini 2.5 Flash API
- **數據存儲**：Google Sheets + PropertiesService

## 授權

本專案僅供學習和研究使用。

## 更新日誌

### 最新版本
- ✅ 新增全文搜尋功能
- ✅ 新增書籤管理功能
- ✅ 新增閱讀進度追蹤
- ✅ 優化手機版體驗
- ✅ 改進代碼結構和性能

