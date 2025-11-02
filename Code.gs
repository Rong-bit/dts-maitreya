function doGet(e) {
  try {
    // 檢查用戶權限
    const user = Session.getActiveUser();
    if (!user) {
      throw new Error('無法獲取用戶信息，請確保您已登入');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error('無法訪問試算表，請確保您有權限');
    }

    // 使用統一的 getAllSheetNames 函數，避免重複載入
    const sheetNames = getAllSheetNames();
    
    // 生成 HTML 內容，只傳遞工作表名稱，不預載任何數據
    const template = HtmlService.createTemplateFromFile('index');
    template.sheetNames = sheetNames;
    // 完全移除數據預載入，改為真正的按需載入
    
    return template.evaluate()
      .setTitle('佛經閱讀系統')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setFaviconUrl('https://www.google.com/favicon.ico')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  } catch (error) {
    Logger.log('錯誤：' + error.toString());
    return HtmlService.createHtmlOutput(`
      <html>
        <head>
          <title>錯誤</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            .error { color: red; }
          </style>
        </head>
        <body>
          <h1 class="error">發生錯誤</h1>
          <p>請檢查以下事項：</p>
          <ol>
            <li>確保您已登入 Google 帳戶</li>
            <li>確保您有權限訪問此試算表</li>
            <li>如果問題持續，請重新整理頁面</li>
          </ol>
          <p>錯誤詳情：${error.toString()}</p>
          <button onclick="window.location.reload()">重新整理</button>
        </body>
      </html>
    `);
  }
}

// 新增：輕量級載入工作表標題行（僅第一行）
function loadSheetHeaders(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error('找不到工作表：' + sheetName);
    }
    
    // 只載入第一行
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    const headerValues = headerRange.getValues()[0];
    
    // 處理標題行數據
    const processedHeaders = [];
    for (let j = 0; j < headerValues.length; j++) {
      processedHeaders.push(headerValues[j] || ''); // 將空值轉換為空字符串
    }
    
    return [processedHeaders]; // 返回二維數組格式以保持一致性
  } catch (error) {
    Logger.log('載入工作表標題錯誤：' + error.toString());
    throw error;
  }
}

// 新增：按需載入特定工作表的數據
function loadSheetData(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error('找不到工作表：' + sheetName);
    }
    
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    // 檢查是否有資料
    if (values.length > 0 && values[0].length > 0) {
      // 確保數據是二維數組
      const processedData = [];
      for (let i = 0; i < values.length; i++) {
        const row = [];
        for (let j = 0; j < values[i].length; j++) {
          row.push(values[i][j] || ''); // 將空值轉換為空字符串
        }
        processedData.push(row);
      }
      
      return processedData;
    } else {
      Logger.log(`工作表 ${sheetName} 沒有數據`);
      return [];
    }
  } catch (error) {
    Logger.log('載入工作表數據錯誤：' + error.toString());
    throw error;
  }
}

// 新增：載入特定品名範圍的數據（優化版本）
function loadChapterData(sheetName, chapterIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error('找不到工作表：' + sheetName);
    }
    
    // 先載入標題行以確定品名位置
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    const headers = headerRange.getValues()[0];
    
    // 找到指定品名的列範圍
    const chapterColumns = [];
    let currentChapterIndex = 0;
    
    // 遍歷奇數欄位（品名欄位）找到目標品名
    for (let col = 0; col < headers.length; col += 2) {
      if (headers[col] && headers[col].toString().trim()) {
        if (currentChapterIndex === parseInt(chapterIndex)) {
          // 找到目標品名，記錄其列範圍（品名列和內容列）
          chapterColumns.push(col, col + 1);
          break;
        }
        currentChapterIndex++;
      }
    }
    
    // 如果沒有找到指定的品名，返回空數據
    if (chapterColumns.length === 0) {
      Logger.log(`找不到品名索引 ${chapterIndex} 在工作表 ${sheetName} 中`);
      return [];
    }
    
    // 找到該品名的數據範圍（從第一行到最後一行有數據的地方）
    const lastRow = sheet.getLastRow();
    
    // 創建結果數組，包含標題行和數據行
    const processedData = [];
    
    // 先添加標題行（只包含目標品名的兩列）
    const headerRow = [];
    for (let col of chapterColumns) {
      headerRow.push(headers[col] || '');
    }
    processedData.push(headerRow);
    
    // 載入數據行（只載入目標品名的兩列）
    if (lastRow > 1) {
      for (let row = 2; row <= lastRow; row++) {
        const dataRow = [];
        let hasData = false;
        
        for (let col of chapterColumns) {
          const cellValue = sheet.getRange(row, col + 1).getValue(); // +1 因為 getRange 是 1-based
          const cellString = (cellValue || '').toString();
          dataRow.push(cellString);
          if (cellString.trim()) {
            hasData = true;
          }
        }
        
        // 只添加有數據的行
        if (hasData) {
          processedData.push(dataRow);
        }
      }
    }
    
    return processedData;
    
  } catch (error) {
    Logger.log('載入品名數據錯誤：' + error.toString());
    throw error;
  }
}


// 新增：獲取所有工作表名稱
function getAllSheetNames() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log('getAllSheetNames: 無法獲取試算表');
      return [];
    }
    
    const sheets = ss.getSheets();
    if (!sheets || sheets.length === 0) {
      Logger.log('getAllSheetNames: 試算表中沒有工作表');
      return [];
    }
    
    const sheetNames = [];
    
    sheets.forEach((sheet, index) => {
      const sheetName = sheet.getName();
      
      // 跳過同音字字典等系統工作表
      if (sheetName !== '同音字字典') {
        sheetNames.push(sheetName);
      }
    });
    
    return sheetNames;
  } catch (error) {
    Logger.log('getAllSheetNames: 獲取工作表名稱錯誤：' + error.toString());
    Logger.log('getAllSheetNames: 錯誤堆棧：' + error.stack);
    return [];
  }
}

function generateHtmlContent(data) {
  const template = HtmlService.createTemplateFromFile('index');
  template.data = data;
  return template.evaluate().getContent();
}

// 檢查是否為管理員
function isAdmin() {
  try {
    const userEmail = Session.getEffectiveUser().getEmail().toLowerCase().trim();
    Logger.log('當前用戶郵箱：' + userEmail);
    
    // 定義允許的管理員郵箱列表
    const adminEmails = [
      'hjr640511@gmail.com'  // 管理員郵箱
    ].map(email => email.toLowerCase().trim());
    
    Logger.log('允許的管理員郵箱：' + adminEmails.join(', '));
    
    // 檢查用戶是否在管理員列表中
    const isAdminUser = adminEmails.includes(userEmail);
    Logger.log('是否為管理員：' + isAdminUser);
    
    if (!isAdminUser) {
      Logger.log('非管理員嘗試訪問：' + userEmail);
    }
    
    return isAdminUser;
  } catch (error) {
    Logger.log('檢查管理員權限錯誤：' + error.toString());
    return false;
  }
}

// 添加自定義菜單
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('佛經系統')
    .addItem('管理同音字', 'showHomophoneManager')
    .addSeparator()
    .addItem('測試 Gemini 2.5 Flash', 'testGemini25Flash')
    .addItem('檢查 API 配額狀態', 'showQuotaStatusDialog')
    .addItem('翻譯選取的欄位', 'translateSelectedRange')
    .addItem('翻譯當前工作表', 'translateCurrentSheet')
    .addItem('翻譯所有工作表', 'translateAllSheets')
    .addSeparator()
    .addItem('設置 API 金鑰', 'showApiKeyDialog')
    .addSeparator()
    .addItem('設置欄位寬度', 'setColumnWidth')
    .addToUi();
}

// 顯示 API 金鑰設置對話框
function showApiKeyDialog() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    '設置 API 金鑰',
    '請輸入您的 Gemini API 金鑰：',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const apiKey = result.getResponseText().trim();
    if (apiKey) {
      try {
        setApiKey(apiKey);
        ui.alert('API 金鑰設置成功！');
      } catch (error) {
        ui.alert('設置失敗：' + error.message);
      }
    } else {
      ui.alert('請輸入有效的 API 金鑰');
    }
  }
}

// 顯示配額狀態對話框
function showQuotaStatusDialog() {
  try {
    const quotaStatus = checkApiQuotaStatus();
    const ui = SpreadsheetApp.getUi();
    
    let message = 'API 配額狀態檢查結果：\n\n';
    message += quotaStatus.message + '\n\n';
    
    if (quotaStatus.quotaExceeded) {
      message += '建議解決方案：\n';
      message += '1. 等待配額重置（通常24小時後重置）\n';
      message += '2. 升級到 Google AI Studio 付費方案\n';
      message += '3. 分批處理翻譯任務\n';
      message += '4. 使用其他翻譯服務\n\n';
      message += '配額重置時間：每日 UTC 00:00';
    } else if (quotaStatus.success) {
      message += '您可以正常使用翻譯功能。\n\n';
      message += '免費層限制：\n';
      message += '- 每日請求限制：200次\n';
      message += '- 每分鐘請求限制：15次\n';
      message += '- 建議在翻譯間隔2秒以上';
    }
    
    ui.alert('API 配額狀態', message, ui.ButtonSet.OK);
  } catch (error) {
    Logger.log('顯示配額狀態對話框錯誤：' + error.toString());
    SpreadsheetApp.getUi().alert('檢查配額狀態時發生錯誤：' + error.message);
  }
}

// 設置 API 金鑰
function setApiKey(apiKey) {
  try {
    if (!isAdmin()) {
      throw new Error('只有管理員可以設置 API 金鑰');
    }
    
    const properties = PropertiesService.getDocumentProperties();
    properties.setProperty('API_KEY', apiKey);
    return 'API 金鑰設置成功！';
  } catch (error) {
    Logger.log('設置 API 金鑰錯誤：' + error.toString());
    throw error;
  }
}

// 檢查 API 金鑰是否已設置
function checkApiKey() {
  try {
    const properties = PropertiesService.getDocumentProperties();
    const apiKey = properties.getProperty('API_KEY');
    
    if (!apiKey) {
      throw new Error('請先設置 API 金鑰');
    }
    
    return apiKey;
  } catch (error) {
    Logger.log('檢查 API 金鑰錯誤：' + error.toString());
    throw error;
  }
}

// 獲取 API 金鑰狀態
function getApiKeyStatus() {
  try {
    const apiKey = checkApiKey();
    const isUserAdmin = isAdmin();
   // const userEmail = Session.getEffectiveUser().getEmail().toLowerCase().trim();
    
    //Logger.log('當前用戶郵箱：' + userEmail);
    Logger.log('API 金鑰狀態：' + (apiKey ? '已設置' : '未設置'));
    Logger.log('用戶管理員狀態：' + isUserAdmin);
    
    return {
      hasApiKey: !!apiKey,
      isAdmin: isUserAdmin,
      message: isUserAdmin ? (apiKey ? 'API 金鑰已設置' : '請先設置 API 金鑰') : '您沒有權限使用此功能'
    };
  } catch (error) {
    Logger.log('獲取 API 金鑰狀態錯誤：' + error.toString());
    return {
      hasApiKey: false,
      isAdmin: false,
      message: '獲取狀態時發生錯誤'
    };
  }
}

// 工具函數：清理翻譯文本（移除換行符和空白）
function cleanTranslationText(text) {
  if (!text) return '';
  return text.replace(/\n/g, '').replace(/\r/g, '').trim();
}

// 工具函數：處理 API 錯誤響應
function handleApiError(response, responseCode) {
  const errorText = response.getContentText();
  Logger.log('API 錯誤響應：' + errorText);
  
  if (responseCode === 429) {
    try {
      const errorData = JSON.parse(errorText);
      const retryAfter = errorData.error?.details?.find(d => d['@type'] === 'type.googleapis.com/google.rpc.RetryInfo')?.retryDelay;
      const waitTime = retryAfter ? parseInt(retryAfter.replace('s', '')) : 60;
      
      Logger.log(`配額限制，需要等待 ${waitTime} 秒`);
      return {
        isQuotaError: true,
        waitTime: waitTime,
        message: `API 配額已用完，請等待 ${waitTime} 秒後重試。建議：1) 等待重試 2) 升級到付費方案 3) 分批處理翻譯`,
        error: new Error(`API 配額已用完，請等待 ${waitTime} 秒後重試`)
      };
    } catch (parseError) {
      Logger.log('解析配額錯誤響應失敗：' + parseError.toString());
    }
  }
  
  return {
    isQuotaError: false,
    message: `API 請求失敗，響應代碼：${responseCode}`,
    error: new Error(`API 請求失敗，響應代碼：${responseCode}，錯誤信息：${errorText}`)
  };
}

// 工具函數：從 API 響應中提取翻譯文本
function extractTranslationFromResponse(result) {
  if (!result.candidates || !result.candidates[0] || !result.candidates[0].content || !result.candidates[0].content.parts[0]) {
    Logger.log('API 響應格式錯誤：' + JSON.stringify(result));
    throw new Error('API 響應格式不正確');
  }
  
  return result.candidates[0].content.parts[0].text;
}

// 工具函數：檢查錯誤是否為配額限制錯誤
function isQuotaError(error) {
  if (!error || !error.message) return false;
  const errorMessage = error.message.toString();
  return errorMessage.includes('配額已用完') || errorMessage.includes('quota') || errorMessage.includes('429');
}

// 檢查 API 配額狀態
function checkApiQuotaStatus() {
  try {
    const apiKey = checkApiKey();
    if (!apiKey) {
      return {
        success: false,
        message: '請先設置 API 金鑰',
        quotaExceeded: false
      };
    }
    
    // 發送一個簡單的測試請求來檢查配額狀態
    const response = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{
          parts: [{
            text: '測試繁體中文'
          }]
        }]
      }),
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      return {
        success: true,
        message: 'API 配額正常，可以進行翻譯',
        quotaExceeded: false
      };
    } else if (responseCode === 429) {
      const errorHandling = handleApiError(response, responseCode);
      return {
        success: false,
        message: errorHandling.message,
        quotaExceeded: true,
        waitTime: errorHandling.waitTime || 60
      };
    } else {
      const errorHandling = handleApiError(response, responseCode);
      return {
        success: false,
        message: errorHandling.message,
        quotaExceeded: false
      };
    }
  } catch (error) {
    Logger.log('檢查 API 配額狀態錯誤：' + error.toString());
    return {
      success: false,
      message: '檢查配額狀態時發生錯誤：' + error.message,
      quotaExceeded: false
    };
  }
}

// 測試 Gemini 2.5 Flash 模型（發送第一個 API 請求）
function testGemini25Flash() {
  try {
    const apiKey = checkApiKey();
    if (!apiKey) {
      throw new Error('請先設置 API 金鑰');
    }
    
    Logger.log('開始測試 Gemini 2.5 Flash 模型...');
    
    const response = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{
          parts: [{
            text: '請用簡潔的現代白話文翻譯這段佛經文言文，使用繁體中文，請將翻譯結果保持在一行內，不要換行：爾時佛告諸比丘：「汝等見是富樓那彌多羅尼子否。我常稱其於說法人中、最為第一。」'
          }]
        }]
      }),
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    Logger.log('API 響應代碼：' + responseCode);
    
    if (responseCode === 200) {
      const result = JSON.parse(response.getContentText());
      Logger.log('API 響應：' + JSON.stringify(result));
      
      const translation = extractTranslationFromResponse(result);
      Logger.log('Gemini 2.5 Flash 翻譯結果：' + translation);
      
      // 移除換行符，確保翻譯結果在一行內
      const cleanTranslation = cleanTranslationText(translation);
      
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        'Gemini 2.5 Flash 測試成功',
        `模型：gemini-2.5-flash\n\n原文：爾時佛告諸比丘：「汝等見是富樓那彌多羅尼子否。我常稱其於說法人中、最為第一。」\n\n翻譯：${cleanTranslation}\n\n✅ Gemini 2.5 Flash 模型運作正常！`,
        ui.ButtonSet.OK
      );
      
      return {
        success: true,
        model: 'gemini-2.5-flash',
        translation: cleanTranslation,
        message: 'Gemini 2.5 Flash 模型測試成功！'
      };
    } else {
      const errorHandling = handleApiError(response, responseCode);
      throw errorHandling.error;
    }
  } catch (error) {
    Logger.log('測試 Gemini 2.5 Flash 錯誤：' + error.toString());
    const ui = SpreadsheetApp.getUi();
    ui.alert('測試失敗', 'Gemini 2.5 Flash 模型測試失敗：' + error.message, ui.ButtonSet.OK);
    
    return {
      success: false,
      model: 'gemini-2.5-flash',
      error: error.message,
      message: 'Gemini 2.5 Flash 模型測試失敗'
    };
  }
}

// AI 翻譯函數（優化版本，支援配額限制處理）
function translateText(text) {
  try {
    const apiKey = checkApiKey();
    if (!apiKey) {
      Logger.log('API 金鑰未設置');
      throw new Error('請先設置 API 金鑰');
    }
    
    Logger.log('開始翻譯文本：' + text.substring(0, 50) + '...');
    
    const response = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{
          parts: [{
            text: `請將以下佛經文言文翻譯成簡潔的現代白話文，使用繁體中文，保持莊重優雅但避免過多解釋與"這段經文可以翻譯成"，請將翻譯結果保持在一行內，不要換行：\n${text}`
          }]
        }]
      }),
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    Logger.log('API 響應代碼：' + responseCode);
    
    if (responseCode !== 200) {
      const errorHandling = handleApiError(response, responseCode);
      throw errorHandling.error;
    }
    
    const result = JSON.parse(response.getContentText());
    Logger.log('API 響應：' + JSON.stringify(result));
    
    const translation = extractTranslationFromResponse(result);
    Logger.log('翻譯成功：' + translation.substring(0, 50) + '...');
    
    // 移除換行符，確保翻譯結果在一行內
    const cleanTranslation = cleanTranslationText(translation);
    
    return cleanTranslation;
  } catch (error) {
    Logger.log('翻譯錯誤詳情：' + error.toString());
    Logger.log('錯誤堆棧：' + error.stack);
    throw new Error('翻譯失敗：' + error.message);
  }
}

// 批量翻譯函數（優化版本，支援配額限制處理）
function batchTranslate() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      throw new Error('工作表為空或只有標題行');
    }
    
    Logger.log('開始批量翻譯，共 ' + (data.length - 1) + ' 行數據');
    
    let successCount = 0;
    let errorCount = 0;
    let quotaExceeded = false;
    
    // 跳過標題行
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0] || data[i][0] === '') {
        Logger.log('跳過空行：' + (i + 1));
        continue;
      }
      
      if (!data[i][1] || data[i][1] === '') { // 如果白話文為空
        Logger.log('正在翻譯第 ' + (i + 1) + ' 行：' + data[i][0]);
        
        try {
          const translation = translateText(data[i][0]);
          Logger.log('翻譯結果：' + translation);
          
          if (translation && !translation.startsWith('翻譯失敗：')) {
            sheet.getRange(i + 1, 2).setValue(translation);
            successCount++;
          } else {
            Logger.log('翻譯失敗或結果為空：' + translation);
            errorCount++;
          }
          
          // 添加延遲以避免觸發 API 限制
          Utilities.sleep(2000);
        } catch (error) {
          Logger.log('第 ' + (i + 1) + ' 行翻譯失敗：' + error.toString());
          
          // 檢查是否為配額限制錯誤
          if (isQuotaError(error)) {
            quotaExceeded = true;
            Logger.log('檢測到配額限制，停止批量翻譯');
            break;
          }
          
          errorCount++;
        }
      }
    }
    
    Logger.log(`批量翻譯完成：成功 ${successCount} 個，失敗 ${errorCount} 個`);
    
    if (quotaExceeded) {
      return `翻譯因配額限制中斷：成功 ${successCount} 個，失敗 ${errorCount} 個。請等待配額重置後繼續。`;
    } else {
      return `翻譯完成：成功 ${successCount} 個，失敗 ${errorCount} 個`;
    }
  } catch (error) {
    Logger.log('批量翻譯錯誤：' + error.toString());
    Logger.log('錯誤堆棧：' + error.stack);
    throw error;
  }
}

// 自動翻譯經文
function translateSutra() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  sheets.forEach(sheet => {
    const range = sheet.getDataRange();
    const values = range.getValues();
    const newValues = [...values];
    
    // 遍歷每一列
    for (let i = 1; i < values.length; i++) {
      // 遍歷所有欄位
      for (let j = 0; j < values[i].length; j++) {
        // 檢查是否為原文欄位（奇數欄）
        if (j % 2 === 0) {
          const originalText = values[i][j];
          const translationColumn = j + 1;
          
          // 只翻譯有原文且對應翻譯欄位為空的內容
          if (originalText && originalText.trim() !== '' && 
              (!values[i][translationColumn] || values[i][translationColumn].trim() === '')) {
            try {
              Logger.log(`正在翻譯：第 ${i + 1} 行，第 ${j + 1} 列 -> 第 ${translationColumn + 1} 列`);
              Logger.log(`原文：${originalText}`);
              
              // 使用 AI 翻譯
              const translation = translateText(originalText);
              
              // 將翻譯放入對應的偶數欄
              newValues[i][translationColumn] = translation;
              Logger.log(`翻譯成功：${translation}`);
            } catch (error) {
              Logger.log(`翻譯失敗：第 ${i + 1} 行，第 ${j + 1} 列 -> 第 ${translationColumn + 1} 列：${error.message}`);
            }
          }
        }
      }
    }
    
    // 更新工作表
    range.setValues(newValues);
  });
}

// 批量翻譯所有工作表
function batchTranslateAllSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    let totalRows = 0;
    
    Logger.log('開始批量翻譯所有工作表');
    
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const data = sheet.getDataRange().getValues();
      totalRows += data.length - 1; // 減去標題行
      
      Logger.log(`正在處理工作表：${sheetName}，共 ${data.length - 1} 行數據`);
      
      // 跳過標題行
      for (let i = 1; i < data.length; i++) {
        // 遍歷所有欄位
        for (let j = 0; j < data[i].length; j++) {
          // 檢查是否為原文欄位（奇數欄）
          if (j % 2 === 0) {
            const originalText = data[i][j];
            const translationColumn = j + 1;
            
            // 只翻譯有原文且對應翻譯欄位為空的內容
            if (originalText && originalText.trim() !== '' && 
                (!data[i][translationColumn] || data[i][translationColumn].trim() === '')) {
              Logger.log(`正在翻譯 ${sheetName} 第 ${i + 1} 行，第 ${j + 1} 列 -> 第 ${translationColumn + 1} 列：${originalText}`);
              
              try {
                const translation = translateText(originalText);
                Logger.log('翻譯結果：' + translation);
                
                if (translation && !translation.startsWith('翻譯失敗：')) {
                  sheet.getRange(i + 1, translationColumn + 1).setValue(translation);
                } else {
                  Logger.log(`第 ${i + 1} 行翻譯結果為空`);
                  throw new Error('翻譯失敗：' + translation);
                }
                
                // 添加延遲以避免觸發 API 限制
                Utilities.sleep(2000);
              } catch (error) {
                Logger.log(`翻譯失敗：第 ${i + 1} 行，第 ${j + 1} 列 -> 第 ${translationColumn + 1} 列：${error.message}`);
              }
            }
          }
        }
      }
    });
    
    Logger.log(`批量翻譯完成，共處理 ${totalRows} 行數據`);
    return '所有工作表翻譯完成';
  } catch (error) {
    Logger.log('批量翻譯錯誤：' + error.toString());
    Logger.log('錯誤堆棧：' + error.stack);
    throw error;
  }
}

// 獲取所有工作表數據
function getSheetsData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const data = {};
    
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      data[sheetName] = sheet.getDataRange().getValues();
    });
    
    Logger.log('獲取到工作表數據：' + Object.keys(data).join(', '));
    return data;
  } catch (error) {
    Logger.log('獲取工作表數據錯誤：' + error.toString());
    throw error;
  }
}

// 獲取指定工作表的數據
function getSheetData(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error('找不到工作表：' + sheetName);
    }
    
    const data = sheet.getDataRange().getValues();
    Logger.log('獲取到工作表 ' + sheetName + ' 的數據，共 ' + data.length + ' 行');
    return data;
  } catch (error) {
    Logger.log('獲取工作表 ' + sheetName + ' 數據錯誤：' + error.toString());
    throw error;
  }
}

// 獲取同音字字典
function getHomophoneDict(chapterTitle, lineNumber) {
  try {
    const fullDict = getFullHomophoneDict();
    chapterTitle = (chapterTitle || '').toString().trim();
    lineNumber = (lineNumber || '').toString().trim();
    Logger.log(`收到查詢參數：chapterTitle=[${chapterTitle}], lineNumber=[${lineNumber}]`);
    // 1. 品名+行號
    if (chapterTitle && lineNumber && fullDict[chapterTitle] && fullDict[chapterTitle][lineNumber]) {
      Logger.log(`優先命中：品名+行號 ${chapterTitle} ${lineNumber}`);
      return fullDict[chapterTitle][lineNumber];
    }
    // 2. 品名預設
    if (chapterTitle && fullDict[chapterTitle] && fullDict[chapterTitle][""]) {
      Logger.log(`命中：品名預設 ${chapterTitle}`);
      return fullDict[chapterTitle][""];
    }
    // 3. 全域預設
    if (fullDict[""] && fullDict[""][""]) {
      Logger.log('命中：全域預設');
      return fullDict[""][""];
    }
    Logger.log('未命中任何同音字層級，回傳空物件');
    return {};
  } catch (error) {
    Logger.log('獲取同音字字典錯誤：' + error.toString());
    return {};
  }
}

// 設置同音字字典
function setHomophoneDict(dict) {
  try {
    if (!isAdmin()) {
      throw new Error('只有管理員可以設置同音字字典');
    }
    
    const properties = PropertiesService.getDocumentProperties();
    properties.setProperty('HOMOPHONE_DICT', JSON.stringify(dict));
    return '同音字字典設置成功！';
  } catch (error) {
    Logger.log('設置同音字字典錯誤：' + error.toString());
    throw error;
  }
}

// 獲取完整同音字字典結構
function getFullHomophoneDict() {
  try {
    // 從 Properties 讀取
    const properties = PropertiesService.getDocumentProperties();
    const storedDict = properties.getProperty('FULL_HOMOPHONE_DICT');
    
    if (storedDict) {
      const parsedDict = JSON.parse(storedDict);
      Logger.log('讀取到的字典結構：' + JSON.stringify(parsedDict));
      return parsedDict;
    }
  } catch (error) {
    Logger.log('讀取存儲的同音字字典失敗：' + error.toString());
  }
  
  // 如果沒有存儲的字典，返回空的字典結構
  const emptyDict = { "": { } };
  Logger.log('返回空字典結構：' + JSON.stringify(emptyDict));
  return emptyDict;
}

// 設置完整同音字字典結構
function setFullHomophoneDict(fullDict) {
  try {
    if (!isAdmin()) {
      throw new Error('只有管理員可以設置同音字字典');
    }
    
    // 這裡可以選擇將完整字典存儲到 Properties 中
    // 或者保持硬編碼方式，只更新特定部分
    const properties = PropertiesService.getDocumentProperties();
    properties.setProperty('FULL_HOMOPHONE_DICT', JSON.stringify(fullDict));
    return '同音字字典設置成功！';
  } catch (error) {
    Logger.log('設置同音字字典錯誤：' + error.toString());
    throw error;
  }
}

// 添加同音字（支援品名和行號）
function addHomophone(char, homophone, chapterTitle = '', lineNumber = '') {
  try {
    if (!isAdmin()) {
      throw new Error('只有管理員可以添加同音字');
    }
    
    // 取得完整的字典結構
    const fullDict = getFullHomophoneDict();
    
    // 確定要修改的層級
    if (chapterTitle && lineNumber) {
      // 品名 + 行號
      if (!fullDict[chapterTitle]) {
        fullDict[chapterTitle] = {};
      }
      if (!fullDict[chapterTitle][lineNumber]) {
        fullDict[chapterTitle][lineNumber] = {};
      }
      fullDict[chapterTitle][lineNumber][char] = homophone;
    } else if (chapterTitle) {
      // 品名預設
      if (!fullDict[chapterTitle]) {
        fullDict[chapterTitle] = {};
      }
      if (!fullDict[chapterTitle][""]) {
        fullDict[chapterTitle][""] = {};
      }
      fullDict[chapterTitle][""][char] = homophone;
    } else {
      // 全域預設
      if (!fullDict[""]) {
        fullDict[""] = {};
      }
      if (!fullDict[""][""]) {
        fullDict[""][""] = {};
      }
      fullDict[""][""][char] = homophone;
    }
    
    return setFullHomophoneDict(fullDict);
  } catch (error) {
    Logger.log('添加同音字錯誤：' + error.toString());
    throw error;
  }
}

// 刪除同音字（支援品名和行號）
function removeHomophone(char, chapterTitle = '', lineNumber = '') {
  try {
    if (!isAdmin()) {
      throw new Error('只有管理員可以刪除同音字');
    }
    
    Logger.log(`開始刪除同音字：char=${char}, chapterTitle=${chapterTitle}, lineNumber=${lineNumber}`);
    
    // 取得完整的字典結構
    const fullDict = getFullHomophoneDict();
    Logger.log(`當前字典結構：${JSON.stringify(fullDict)}`);
    
    // 確定要刪除的層級
    if (chapterTitle && lineNumber) {
      // 品名 + 行號
      Logger.log(`嘗試刪除品名+行號層級：${chapterTitle} -> ${lineNumber} -> ${char}`);
      if (fullDict[chapterTitle] && fullDict[chapterTitle][lineNumber]) {
        if (fullDict[chapterTitle][lineNumber][char]) {
          delete fullDict[chapterTitle][lineNumber][char];
          Logger.log(`成功刪除品名+行號層級的同音字`);
        } else {
          Logger.log(`在品名+行號層級找不到字符：${char}`);
        }
      } else {
        Logger.log(`品名+行號層級不存在：${chapterTitle} -> ${lineNumber}`);
      }
    } else if (chapterTitle) {
      // 品名預設
      Logger.log(`嘗試刪除品名預設層級：${chapterTitle} -> ${char}`);
      if (fullDict[chapterTitle] && fullDict[chapterTitle][""]) {
        if (fullDict[chapterTitle][""][char]) {
          delete fullDict[chapterTitle][""][char];
          Logger.log(`成功刪除品名預設層級的同音字`);
        } else {
          Logger.log(`在品名預設層級找不到字符：${char}`);
        }
      } else {
        Logger.log(`品名預設層級不存在：${chapterTitle}`);
      }
    } else {
      // 全域預設
      Logger.log(`嘗試刪除全域預設層級：${char}`);
      if (fullDict[""] && fullDict[""][""]) {
        if (fullDict[""][""][char]) {
          delete fullDict[""][""][char];
          Logger.log(`成功刪除全域預設層級的同音字`);
        } else {
          Logger.log(`在全域預設層級找不到字符：${char}`);
        }
      } else {
        Logger.log(`全域預設層級不存在`);
      }
    }
    
    Logger.log(`刪除後的字典結構：${JSON.stringify(fullDict)}`);
    return setFullHomophoneDict(fullDict);
  } catch (error) {
    Logger.log('刪除同音字錯誤：' + error.toString());
    throw error;
  }
}

// 獲取完整同音字列表（支援品名和行號）
function getHomophoneList() {
  try {
    const fullDict = getFullHomophoneDict();
    const result = [];
    
    Logger.log('開始遍歷字典結構：' + JSON.stringify(fullDict));
    
    // 遍歷所有品名
    for (const [chapterTitle, chapterDict] of Object.entries(fullDict)) {
      Logger.log(`處理品名：${chapterTitle}, 類型：${typeof chapterDict}`);
      
      // 檢查 chapterDict 是否為對象
      if (typeof chapterDict === 'object' && chapterDict !== null) {
        // 遍歷所有行號
        for (const [lineNumber, lineDict] of Object.entries(chapterDict)) {
          Logger.log(`處理行號：${lineNumber}, 類型：${typeof lineDict}`);
          
          // 檢查 lineDict 是否為對象
          if (typeof lineDict === 'object' && lineDict !== null) {
            // 遍歷所有字符
            for (const [char, homophone] of Object.entries(lineDict)) {
              const item = {
                chapterTitle: chapterTitle || '全域',
                lineNumber: lineNumber === '' ? '預設' : (lineNumber === '0' ? '預設' : lineNumber),
                originalLineNumber: lineNumber, // 保存原始行號用於刪除
                char: char,
                homophone: homophone, // 保持向後相容
                example: `例：${char}（${homophone}）`
              };
              
              // 調試信息
              Logger.log(`添加同音字項目：品名=${item.chapterTitle}, 行號=${item.lineNumber}, 漢字=${item.char}, 同音字=${item.homophone}`);
              Logger.log(`原始數據：chapterTitle=${chapterTitle}, lineNumber=${lineNumber}, char=${char}, homophone=${homophone}`);
              
              result.push(item);
            }
          } else {
            Logger.log(`lineDict 不是對象，跳過：${lineNumber}`);
          }
        }
      } else {
        Logger.log(`chapterDict 不是對象，跳過：${chapterTitle}`);
      }
    }
    
    Logger.log(`最終結果數量：${result.length}`);
    return result;
  } catch (error) {
    Logger.log('獲取同音字列表錯誤：' + error.toString());
    return [];
  }
}

// 批次導入同音字（支援品名和行號，支援逗號、Tab、空格分隔）
function batchImportHomophones(data) {
  try {
    if (!isAdmin()) {
      throw new Error('只有管理員可以批次導入同音字');
    }
    
    const fullDict = getFullHomophoneDict();
    let importedCount = 0;
    let skippedCount = 0;
    
    // 處理批次資料
    if (typeof data === 'string') {
      const lines = data.split('\n');
      for (let line of lines) {
        line = line.trim();
        if (!line) continue;
        // 自動判斷分隔符：逗號、Tab、或多個空格
        let parts = [];
        if (line.includes(',')) {
          parts = line.split(',').map(item => item.trim());
        } else if (line.includes('\t')) {
          parts = line.split(/\t+/).map(item => item.trim());
        } else {
          // 多個空白分隔
          parts = line.split(/\s{2,}|\s+\s+/).map(item => item.trim());
          // 若只有一個空格，仍嘗試分割
          if (parts.length < 2) {
            parts = line.split(/\s+/).map(item => item.trim());
          }
        }
        if (parts.length >= 2) {
          const [char, homophone, chapterTitle = '', lineNumber = ''] = parts;
          if (char && homophone) {
            // 確定要修改的層級
            if (chapterTitle && lineNumber) {
              // 品名 + 行號
              if (!fullDict[chapterTitle]) {
                fullDict[chapterTitle] = {};
              }
              if (!fullDict[chapterTitle][lineNumber]) {
                fullDict[chapterTitle][lineNumber] = {};
              }
              fullDict[chapterTitle][lineNumber][char] = homophone;
            } else if (chapterTitle) {
              // 品名預設
              if (!fullDict[chapterTitle]) {
                fullDict[chapterTitle] = {};
              }
              if (!fullDict[chapterTitle][""]) {
                fullDict[chapterTitle][""] = {};
              }
              fullDict[chapterTitle][""][char] = homophone;
            } else {
              // 全域預設
              if (!fullDict[""]) {
                fullDict[""] = {};
              }
              if (!fullDict[""][""]) {
                fullDict[""][""] = {};
              }
              fullDict[""][""][char] = homophone;
            }
            importedCount++;
          } else {
            skippedCount++;
          }
        } else {
          skippedCount++;
        }
      }
    }
    // 處理 JSON 格式
    else if (typeof data === 'object') {
      for (const [char, homophone] of Object.entries(data)) {
        if (char && homophone) {
          // JSON 格式預設為全域
          if (!fullDict[""]) {
            fullDict[""] = {};
          }
          if (!fullDict[""][""]) {
            fullDict[""][""] = {};
          }
          fullDict[""][""][char] = homophone;
          importedCount++;
        } else {
          skippedCount++;
        }
      }
    }
    
    setFullHomophoneDict(fullDict);
    return {
      success: true,
      importedCount,
      skippedCount,
      message: `成功導入 ${importedCount} 個同音字，跳過 ${skippedCount} 個無效數據`
    };
  } catch (error) {
    Logger.log('批次導入同音字錯誤：' + error.toString());
    throw error;
  }
}

// 導出同音字字典
function exportHomophones(format = 'simple') {
  try {
    const fullDict = getFullHomophoneDict();
    let csvContent = '';
    
    if (format === 'simple') {
      // 簡化格式：只有漢字和同音字
      const simpleDict = {};
      
      // 收集所有層級的同音字，去重
      for (const [chapterTitle, chapterDict] of Object.entries(fullDict)) {
        for (const [lineNumber, lineDict] of Object.entries(chapterDict)) {
          for (const [char, homophone] of Object.entries(lineDict)) {
            // 如果同一個漢字有多個同音字，保留最後一個
            simpleDict[char] = homophone;
          }
        }
      }
      
      csvContent = Object.entries(simpleDict)
        .map(([char, homophone]) => `${char},${homophone}`)
        .join('\n');
    } else {
      // 完整格式：包含品名和行號
      const fullList = [];
      
      for (const [chapterTitle, chapterDict] of Object.entries(fullDict)) {
        for (const [lineNumber, lineDict] of Object.entries(chapterDict)) {
          if (lineNumber === '0') continue; // 忽略行號為 '0' 的條目
          for (const [char, homophone] of Object.entries(lineDict)) {
            fullList.push({
              char,
              homophone,
              chapterTitle: chapterTitle || '',
              lineNumber: lineNumber || ''
            });
          }
        }
      }
      
      csvContent = fullList
        .map(item => `${item.char},${item.homophone},${item.chapterTitle},${item.lineNumber}`)
        .join('\n');
    }
    
    // 創建 BOM 標記的 UTF-8 編碼
    const bom = '\uFEFF';
    const blob = Utilities.newBlob(bom + csvContent, 'text/csv;charset=utf-8', 'homophones.csv');
    
    return {
      success: true,
      data: fullDict,
      csv: csvContent,
      blob: blob,
      format: format
    };
  } catch (error) {
    Logger.log('導出同音字字典錯誤：' + error.toString());
    throw error;
  }
}

// 顯示同音字管理界面
function showHomophoneManager() {
  try {
    if (!isAdmin()) {
      throw new Error('只有管理員可以管理同音字');
    }
    
    const html = HtmlService.createHtmlOutput(`
      <html>
        <head>
          <base target="_top">
          <meta charset="UTF-8">
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            .container { max-width: 800px; margin: 0 auto; }
            .form-group { margin-bottom: 15px; }
            label { display: block; margin-bottom: 5px; }
            input { width: 100%; padding: 8px; margin-bottom: 10px; }
            button { padding: 8px 15px; margin-right: 10px; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; }
            th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
            th { background-color: #f5f5f5; }
            .delete-btn { color: red; cursor: pointer; }
            .edit-btn { color: #4CAF50; cursor: pointer; }
            .edit-btn:hover { text-decoration: underline; }
            .delete-btn:hover { text-decoration: underline; }
            .import-export { margin: 20px 0; padding: 15px; background: #f5f5f5; }
            textarea { width: 100%; height: 100px; margin: 10px 0; }
            .example { font-size: 0.9em; color: #666; margin-top: 5px; }
            .error { color: red; font-size: 0.9em; margin-top: 5px; }
            .success { color: green; font-size: 0.9em; margin-top: 5px; }
            .char-preview { font-size: 24px; margin: 10px 0; padding: 10px; border: 1px solid #ddd; background: #f9f9f9; text-align: center; }
          </style>
        </head>
        <body>
          <div class="container">
          
            
            <div class="import-export">
              <h3>批次導入/導出</h3>
              <div>
                <button onclick="exportHomophones('simple')">導出簡化格式（僅漢字+同音字）</button>
                <button onclick="exportHomophones('full')">導出完整格式（含品名+行號）</button>
                <button onclick="showImportDialog()">批次導入</button>
                <button onclick="clearAllHomophones()" style="background-color: #ff6b6b; color: white;">清除所有同音字</button>
              </div>
              <div id="importDialog" style="display: none; margin-top: 15px;">
                <p>請輸入同音字數據（CSV 格式，每行一個）：</p>
                <textarea id="importData" placeholder="例：\n行,航,序品,1\n行,衡,序品,5\n佛,福\n\n或\n行    航    序品    1\n行    衡    序品    5\n佛    福\n（可用逗號、Tab 或多空格分隔，也可直接從 Excel/試算表複製貼上）"></textarea>
                <div class="example">
                  格式說明：<br>
                  1. 每行一個同音字，可用逗號、Tab 或多個空格分隔：漢字,同音字,品名,行號<br>
                  2. 品名和行號可選，留空表示全域預設<br>
                  3. 每行一個同音字<br>
                  4. 支持批量導入，可直接從 Excel/試算表複製貼上<br>
                  5. 支持生僻字和特殊字符
                </div>
                <button onclick="importHomophones()">導入</button>
                <button onclick="hideImportDialog()">取消</button>
              </div>
            </div>
            
            <div class="form-group">
              <label for="chapterTitle">品名：</label>
              <select id="chapterTitle">
                <option value="">全域</option>
              </select>
              <div class="example">提示：留空表示全域設定，選擇品名表示該品名的預設</div>
            </div>
            <div class="form-group">
              <label for="lineNumber">行號：</label>
              <input type="number" id="lineNumber" placeholder="請輸入行號（留空為該品名預設）" min="1">
              <div class="example">提示：留空表示該品名的預設，輸入數字表示特定行</div>
            </div>
            <div class="form-group">
              <label for="char">漢字：</label>
              <input type="text" id="char" placeholder="請輸入漢字（支持生僻字）" oninput="previewChar()">
              <div id="charPreview" class="char-preview" style="display: none;"></div>
              <div class="example">提示：可以輸入單個漢字或多個字符，支持複製貼上生僻字</div>
            </div>
            <div class="form-group">
                      <label for="homophone">同音字：</label>
        <input type="text" id="homophone" placeholder="請輸入同音字">
            </div>
            <div id="message"></div>
            <button onclick="addHomophone()" id="addButton">添加</button>
            <button onclick="updateHomophone()" id="updateButton" style="display: none; background-color: #4CAF50; color: white;">更新</button>
            <button onclick="cancelEdit()" id="cancelButton" style="display: none;">取消</button>
            <button onclick="closeDialog()">關閉</button>
            
            <table id="homophoneTable">
              <thead>
                <tr>
                  <th>品名</th>
                  <th>行號</th>
                  <th>漢字</th>
                  <th>同音字</th>
                  <th>操作</th>
                </tr>
              </thead>
              <tbody></tbody>
            </table>
          </div>
          <script>
            function loadHomophones() {
              google.script.run
                .withSuccessHandler(updateTable)
                .withFailureHandler(showError)
                .getHomophoneList();
            }
            
            function updateTable(homophones) {
              const tbody = document.querySelector('#homophoneTable tbody');
              tbody.innerHTML = '';
              
              homophones.forEach(homophone => {
                const tr = document.createElement('tr');
                tr.innerHTML = \`
                  <td>\${homophone.chapterTitle}</td>
                  <td>\${homophone.lineNumber}</td>
                  <td>\${homophone.char}</td>
                  <td>\${homophone.homophone}</td>
                  <td>
                    <span class="edit-btn" onclick="editHomophone('\${homophone.char}', '\${homophone.homophone}', '\${homophone.chapterTitle}', '\${homophone.originalLineNumber}')" style="color: #4CAF50; cursor: pointer; margin-right: 10px;">編輯</span>
                    <span class="delete-btn" onclick="removeHomophone('\${homophone.char}', '\${homophone.chapterTitle}', '\${homophone.originalLineNumber}')" style="color: red; cursor: pointer;">刪除</span>
                  </td>
                \`;
                tbody.appendChild(tr);
              });
            }
            
            function previewChar() {
              const char = document.getElementById('char').value;
              const preview = document.getElementById('charPreview');
              
              if (char) {
                preview.textContent = char;
                preview.style.display = 'block';
              } else {
                preview.style.display = 'none';
              }
            }
            
            function showMessage(text, isError = false) {
              const messageDiv = document.getElementById('message');
              messageDiv.textContent = text;
              messageDiv.className = isError ? 'error' : 'success';
              messageDiv.style.display = 'block';
              
              setTimeout(() => {
                messageDiv.style.display = 'none';
              }, 3000);
            }
            
            function addHomophone() {
              const chapterTitle = document.getElementById('chapterTitle').value.trim();
              const lineNumber = document.getElementById('lineNumber').value.trim();
              const char = document.getElementById('char').value.trim();
              const homophone = document.getElementById('homophone').value.trim();
              
              if (!char) {
                showMessage('請輸入漢字', true);
                return;
              }
              
              if (!homophone) {
                showMessage('請輸入同音字', true);
                return;
              }
              
              // 檢查漢字是否為有效字符
              if (char.length === 0) {
                showMessage('漢字不能為空', true);
                return;
              }
              
              google.script.run
                .withSuccessHandler(() => {
                  // 清空表單
                  document.getElementById('chapterTitle').value = '';
                  document.getElementById('lineNumber').value = '';
                  document.getElementById('char').value = '';
                  document.getElementById('homophone').value = '';
                  document.getElementById('charPreview').style.display = 'none';
                  
                  // 確保按鈕狀態正確
                  document.getElementById('addButton').style.display = 'inline-block';
                  document.getElementById('updateButton').style.display = 'none';
                  document.getElementById('cancelButton').style.display = 'none';
                  
                  // 清除編輯數據
                  window.editingData = null;
                  
                  showMessage('同音字添加成功！');
                  loadHomophones();
                })
                .withFailureHandler(error => {
                  showMessage('添加失敗：' + error, true);
                })
                .addHomophone(char, homophone, chapterTitle, lineNumber);
            }
            
            function removeHomophone(char, chapterTitle, lineNumber) {
              if (confirm('確定要刪除這個同音字嗎？')) {
                // 處理顯示名稱到實際值的轉換
                const actualChapterTitle = chapterTitle === '全域' ? '' : chapterTitle;
                const actualLineNumber = lineNumber === '預設' ? '' : lineNumber;
                

                
                google.script.run
                  .withSuccessHandler(() => {
                    showMessage('同音字刪除成功！');
                    loadHomophones();
                  })
                  .withFailureHandler(error => {
                    showMessage('刪除失敗：' + error, true);
                  })
                  .removeHomophone(char, actualChapterTitle, actualLineNumber);
              }
            }
            
            function editHomophone(char, homophone, chapterTitle, lineNumber) {
              // 處理顯示名稱到實際值的轉換
              const actualChapterTitle = chapterTitle === '全域' ? '' : chapterTitle;
              const actualLineNumber = lineNumber === '預設' ? '' : lineNumber;
              
              // 填充表單
              document.getElementById('chapterTitle').value = actualChapterTitle;
              document.getElementById('lineNumber').value = actualLineNumber;
              document.getElementById('char').value = char;
              document.getElementById('homophone').value = homophone;
              
              // 顯示字符預覽
              const preview = document.getElementById('charPreview');
              preview.textContent = char;
              preview.style.display = 'block';
              
              // 切換按鈕顯示
              document.getElementById('addButton').style.display = 'none';
              document.getElementById('updateButton').style.display = 'inline-block';
              document.getElementById('cancelButton').style.display = 'inline-block';
              
              // 保存原始數據用於更新
              window.editingData = {
                originalChar: char,
                originalChapterTitle: actualChapterTitle,
                originalLineNumber: actualLineNumber
              };
              
              showMessage('請修改同音字信息，然後點擊「更新」按鈕', false);
            }
            
            function updateHomophone() {
              const chapterTitle = document.getElementById('chapterTitle').value.trim();
              const lineNumber = document.getElementById('lineNumber').value.trim();
              const char = document.getElementById('char').value.trim();
              const homophone = document.getElementById('homophone').value.trim();
              
              if (!char) {
                showMessage('請輸入漢字', true);
                return;
              }
              
              if (!homophone) {
                showMessage('請輸入同音字', true);
                return;
              }
              
              if (!window.editingData) {
                showMessage('編輯數據丟失，請重新選擇要編輯的項目', true);
                return;
              }
              
              // 先刪除舊的同音字
              google.script.run
                .withSuccessHandler(() => {
                  // 再添加新的同音字
                  google.script.run
                    .withSuccessHandler(() => {
                      showMessage('同音字更新成功！');
                      cancelEdit();
                      loadHomophones();
                    })
                    .withFailureHandler(error => {
                      showMessage('更新失敗：' + error, true);
                    })
                    .addHomophone(char, homophone, chapterTitle, lineNumber);
                })
                .withFailureHandler(error => {
                  showMessage('刪除舊數據失敗：' + error, true);
                })
                .removeHomophone(window.editingData.originalChar, window.editingData.originalChapterTitle, window.editingData.originalLineNumber);
            }
            
            function cancelEdit() {
              // 清空表單
              document.getElementById('chapterTitle').value = '';
              document.getElementById('lineNumber').value = '';
              document.getElementById('char').value = '';
              document.getElementById('homophone').value = '';
              document.getElementById('charPreview').style.display = 'none';
              
              // 切換按鈕顯示
              document.getElementById('addButton').style.display = 'inline-block';
              document.getElementById('updateButton').style.display = 'none';
              document.getElementById('cancelButton').style.display = 'none';
              
              // 清除編輯數據
              window.editingData = null;
              
              showMessage('已取消編輯', false);
            }
            
            function showImportDialog() {
              document.getElementById('importDialog').style.display = 'block';
            }
            
            function hideImportDialog() {
              document.getElementById('importDialog').style.display = 'none';
            }
            
            function importHomophones() {
              const data = document.getElementById('importData').value;
              if (!data) {
                alert('請輸入要導入的數據');
                return;
              }
              
              google.script.run
                .withSuccessHandler(result => {
                  alert(result.message);
                  document.getElementById('importData').value = '';
                  hideImportDialog();
                  loadHomophones();
                })
                .withFailureHandler(error => {
                  alert('導入失敗：' + error);
                })
                .batchImportHomophones(data);
            }
            
            function exportHomophones(format = 'simple') {
              google.script.run
                .withSuccessHandler(result => {
                  if (result.success) {
                    // 創建下載鏈接
                    const link = document.createElement('a');
                    link.href = 'data:text/csv;charset=utf-8,\ufeff' + encodeURIComponent(result.csv);
                    link.download = format === 'simple' ? 'homophones_simple.csv' : 'homophones_full.csv';
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                  } else {
                    showError('導出失敗');
                  }
                })
                .withFailureHandler(showError)
                .exportHomophones(format);
            }
            
            function clearAllHomophones() {
              if (confirm('確定要清除所有同音字嗎？此操作無法復原！')) {
                google.script.run
                  .withSuccessHandler(result => {
                    alert(result);
                    loadHomophones();
                  })
                  .withFailureHandler(error => {
                    alert('清除失敗：' + error);
                  })
                  .clearAllHomophones();
              }
            }
            

            
            function showError(error) {
              alert('發生錯誤：' + error);
            }
            
            function closeDialog() {
              google.script.host.close();
            }
            
            // 載入品名下拉選單
            function loadSheetNames() {
              google.script.run.withSuccessHandler(function(names) {
                const select = document.getElementById('chapterTitle');
                // 保留全域選項
                select.innerHTML = '<option value="">全域</option>';
                names.forEach(name => {
                  const option = document.createElement('option');
                  option.value = name;
                  option.textContent = name;
                  select.appendChild(option);
                });
              }).getAllSheetAndChapterNames();
            }
            window.onload = function() {
              loadSheetNames();
              loadHomophones();
            };
          </script>
        </body>
      </html>
    `)
    .setWidth(800)
    .setHeight(700);
    
    SpreadsheetApp.getUi().showModalDialog(html, '同音字管理');
  } catch (error) {
    Logger.log('顯示同音字管理界面錯誤：' + error.toString());
    SpreadsheetApp.getUi().alert('發生錯誤：' + error.message);
  }
}

// 獲取所有經文資料
function getAllSutraData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const data = {};
    
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const range = sheet.getDataRange();
      const values = range.getValues();
      
      // 檢查是否有資料
      if (values.length > 0 && values[0].length > 0) {
        data[sheetName] = values;
      }
    });
    
    Logger.log('成功讀取所有經文資料，共 ' + Object.keys(data).length + ' 部經文');
    return data;
  } catch (error) {
    Logger.log('獲取經文資料錯誤：' + error.toString());
    return {};
  }
}

// 翻譯選取的欄位（優化版本，支援配額檢查）
function translateSelectedRange() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();
    
    // 檢查是否為同音字字典工作表
    if (sheetName === '同音字字典') {
      throw new Error('同音字字典工作表不需要翻譯');
    }
    
    // 先檢查配額狀態
    const quotaStatus = checkApiQuotaStatus();
    if (quotaStatus.quotaExceeded) {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        '配額限制',
        `API 配額已用完，請等待 ${quotaStatus.waitTime || 60} 秒後重試。\n\n建議：\n1. 等待配額重置\n2. 升級到付費方案\n3. 分批處理翻譯`,
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 獲取選取的範圍
    const selectedRange = sheet.getActiveRange();
    if (!selectedRange) {
      throw new Error('請先選取要翻譯的欄位範圍');
    }
    
    const ui = SpreadsheetApp.getUi();
    const rangeA1 = selectedRange.getA1Notation();
    const response = ui.alert(
      '確認翻譯',
      `確定要翻譯選取的範圍「${rangeA1}」嗎？\n這可能需要一些時間。`,
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      const values = selectedRange.getValues();
      const newValues = [...values];
      const startRow = selectedRange.getRow();
      const startCol = selectedRange.getColumn();
      
      Logger.log(`開始翻譯選取範圍：${rangeA1}，起始行：${startRow}，起始列：${startCol}`);
      
      let successCount = 0;
      let errorCount = 0;
      let quotaExceeded = false;
      
      // 遍歷選取範圍內的每一行
      for (let i = 0; i < values.length; i++) {
        // 遍歷選取範圍內的每一列
        for (let j = 0; j < values[i].length; j++) {
          const actualRow = startRow + i;
          const actualCol = startCol + j;
          
          // 檢查是否為原文欄位（奇數欄，從1開始計算）
          if (actualCol % 2 === 1) {
            const originalText = values[i][j];
            const translationColumn = actualCol + 1; // 對應的翻譯欄位
            
            // 只翻譯有原文且對應翻譯欄位為空的內容
            if (originalText && originalText.trim() !== '' && 
                (!values[i][translationColumn - startCol] || values[i][translationColumn - startCol].trim() === '')) {
              try {
                Logger.log(`正在翻譯：第 ${actualRow} 行，第 ${actualCol} 列 -> 第 ${translationColumn} 列`);
                Logger.log(`原文：${originalText}`);
                
                // 使用 AI 翻譯
                const translation = translateText(originalText);
                
                // 檢查對應的翻譯欄位是否在選取範圍內
                const translationColIndex = translationColumn - startCol;
                if (translationColIndex >= 0 && translationColIndex < values[i].length) {
                  // 翻譯欄位在選取範圍內，直接更新
                  newValues[i][translationColIndex] = translation;
                } else {
                  // 翻譯欄位不在選取範圍內，直接寫入工作表
                  sheet.getRange(actualRow, translationColumn).setValue(translation);
                }
                
                Logger.log(`翻譯成功：${translation}`);
                successCount++;
                
                // 添加延遲以避免觸發 API 限制
                Utilities.sleep(2000);
              } catch (error) {
                Logger.log(`翻譯失敗：第 ${actualRow} 行，第 ${actualCol} 列 -> 第 ${translationColumn} 列：${error.message}`);
                
                // 檢查是否為配額限制錯誤
                if (isQuotaError(error)) {
                  quotaExceeded = true;
                  Logger.log('檢測到配額限制，停止翻譯');
                  break;
                }
                
                errorCount++;
              }
            }
          }
        }
        
        // 如果遇到配額限制，跳出外層循環
        if (quotaExceeded) break;
      }
      
      // 更新選取範圍
      selectedRange.setValues(newValues);
      
      if (quotaExceeded) {
        ui.alert('翻譯中斷', `選取範圍「${rangeA1}」的翻譯因配額限制中斷：成功 ${successCount} 個，失敗 ${errorCount} 個。請等待配額重置後繼續。`, ui.ButtonSet.OK);
      } else {
        ui.alert('翻譯完成', `選取範圍「${rangeA1}」的翻譯已完成：成功 ${successCount} 個，失敗 ${errorCount} 個！`, ui.ButtonSet.OK);
      }
    }
  } catch (error) {
    Logger.log('翻譯錯誤：' + error.toString());
    SpreadsheetApp.getUi().alert('翻譯失敗：' + error.message);
  }
}

// 翻譯當前工作表
function translateCurrentSheet() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();
    
    // 檢查是否為同音字字典工作表
    if (sheetName === '同音字字典') {
      throw new Error('同音字字典工作表不需要翻譯');
    }
    
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '確認翻譯',
      `確定要翻譯「${sheetName}」嗎？\n這可能需要一些時間。`,
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      const range = sheet.getDataRange();
      const values = range.getValues();
      const newValues = [...values];
      
      // 遍歷每一列
      for (let i = 1; i < values.length; i++) {
        // 遍歷所有欄位
        for (let j = 0; j < values[i].length; j++) {
          // 檢查是否為原文欄位（奇數欄）
          if (j % 2 === 0) {
            const originalText = values[i][j];
            const translationColumn = j + 1;
            
            // 只翻譯有原文且對應翻譯欄位為空的內容
            if (originalText && originalText.trim() !== '' && 
                (!values[i][translationColumn] || values[i][translationColumn].trim() === '')) {
              try {
                Logger.log(`正在翻譯：第 ${i + 1} 行，第 ${j + 1} 列 -> 第 ${translationColumn + 1} 列`);
                Logger.log(`原文：${originalText}`);
                
                // 使用 AI 翻譯
                const translation = translateText(originalText);
                
                // 將翻譯放入對應的偶數欄
                newValues[i][translationColumn] = translation;
                Logger.log(`翻譯成功：${translation}`);
                
                // 添加延遲以避免觸發 API 限制
                Utilities.sleep(2000);
              } catch (error) {
                Logger.log(`翻譯失敗：第 ${i + 1} 行，第 ${j + 1} 列 -> 第 ${translationColumn + 1} 列：${error.message}`);
              }
            }
          }
        }
      }
      
      // 更新工作表
      range.setValues(newValues);
      ui.alert('翻譯完成', '當前工作表的翻譯已完成！', ui.ButtonSet.OK);
    }
  } catch (error) {
    Logger.log('翻譯錯誤：' + error.toString());
    SpreadsheetApp.getUi().alert('翻譯失敗：' + error.message);
  }
}

// 翻譯所有工作表
function translateAllSheets() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '確認翻譯',
      '確定要翻譯所有工作表嗎？\n這可能需要較長時間。',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      batchTranslateAllSheets();
      ui.alert('翻譯完成', '所有工作表的翻譯已完成！', ui.ButtonSet.OK);
    }
  } catch (error) {
    Logger.log('翻譯錯誤：' + error.toString());
    SpreadsheetApp.getUi().alert('翻譯失敗：' + error.message);
  }
}

// 設置欄位寬度
function setColumnWidth() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();
    
    // 檢查是否為同音字字典工作表
    if (sheetName === '同音字字典') {
      throw new Error('同音字字典工作表不需要設置欄位寬度');
    }
    
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '確認設置',
      `確定要將「${sheetName}」的所有欄位寬度設置為 700 嗎？`,
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      const numColumns = sheet.getLastColumn();
      
      // 設置所有欄位的寬度為 700
      for (let i = 1; i <= numColumns; i++) {
        sheet.setColumnWidth(i, 700);
      }
      
      ui.alert('設置完成', '欄位寬度已設置完成！', ui.ButtonSet.OK);
    }
  } catch (error) {
    Logger.log('設置欄位寬度錯誤：' + error.toString());
    SpreadsheetApp.getUi().alert('設置失敗：' + error.message);
  }
}

// 設置所有經文工作表的欄位寬度
function setAllSheetsColumnWidth() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const ui = SpreadsheetApp.getUi();
    
    const response = ui.alert(
      '確認設置',
      '確定要將所有經文工作表的欄位寬度設置為 700 嗎？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      sheets.forEach(sheet => {
        const sheetName = sheet.getName();
        if (sheetName !== '同音字字典') {
          const numColumns = sheet.getLastColumn();
          for (let i = 1; i <= numColumns; i++) {
            sheet.setColumnWidth(i, 700);
          }
        }
      });
      
      ui.alert('設置完成', '所有工作表的欄位寬度已設置完成！', ui.ButtonSet.OK);
    }
  } catch (error) {
    Logger.log('設置所有工作表欄位寬度錯誤：' + error.toString());
    SpreadsheetApp.getUi().alert('設置失敗：' + error.message);
  }
}

// 移除錯誤的同音字條目
function removeIncorrectHomophone() {
  try {
    if (!isAdmin()) {
      throw new Error('只有管理員可以移除錯誤的同音字條目');
    }
    
    const dict = getHomophoneDict();
    
    // 移除「合」字的錯誤同音字條目
    if (dict['合']) {
      delete dict['合'];
      setHomophoneDict(dict);
      return '已移除錯誤的「合」字同音字條目';
    } else {
      return '「合」字同音字條目不存在';
    }
  } catch (error) {
    Logger.log('移除錯誤同音字條目錯誤：' + error.toString());
    throw error;
  }
}



// 清除所有同音字字典（重置為空白）
function clearAllHomophones() {
  try {
    if (!isAdmin()) {
      throw new Error('只有管理員可以清除同音字字典');
    }
    
    const properties = PropertiesService.getDocumentProperties();
    properties.deleteProperty('FULL_HOMOPHONE_DICT');
    
    // 重新初始化正確的字典結構
    const correctDict = { "": { "": {} } };
    properties.setProperty('FULL_HOMOPHONE_DICT', JSON.stringify(correctDict));
    
    Logger.log('字典已重置為正確結構：' + JSON.stringify(correctDict));
    
    return '同音字字典已完全清除並重置為正確結構！';
  } catch (error) {
    Logger.log('清除同音字字典錯誤：' + error.toString());
    throw error;
  }
}



// 取得所有工作表名稱與所有章節名稱（奇數欄位第1行），去重、去空
function getAllSheetAndChapterNames() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const nameSet = new Set();
    // 加入所有工作表名稱
    sheets.forEach(sheet => nameSet.add(sheet.getName()));
    // 加入所有工作表第1行奇數欄位內容
    sheets.forEach(sheet => {
      const firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      for (let col = 0; col < firstRow.length; col += 2) {
        const value = (firstRow[col] || '').toString().trim();
        if (value) nameSet.add(value);
      }
    });
    return Array.from(nameSet);
  } catch (error) {
    Logger.log('獲取品名與章節名稱錯誤：' + error.toString());
    return [];
  }
}

// ==================== 閱讀進度追蹤功能 ====================

// 保存閱讀進度
function saveReadingProgress(sheetName, chapterIndex, lineNumber) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const properties = PropertiesService.getUserProperties();
    const progressKey = `READING_PROGRESS_${userEmail}`;
    
    // 獲取現有進度
    let allProgress = {};
    const existingProgress = properties.getProperty(progressKey);
    if (existingProgress) {
      try {
        allProgress = JSON.parse(existingProgress);
      } catch (e) {
        Logger.log('解析現有進度失敗，重新初始化');
        allProgress = {};
      }
    }
    
    // 更新或創建進度記錄
    const progressData = {
      sheetName: sheetName,
      chapterIndex: chapterIndex,
      lineNumber: lineNumber,
      timestamp: new Date().toISOString()
    };
    
    // 使用工作表名稱作為鍵，一個用戶只能有一個當前進度
    allProgress.current = progressData;
    
    // 也保存歷史記錄（最多保留最近10個）
    if (!allProgress.history) {
      allProgress.history = [];
    }
    allProgress.history.unshift(progressData);
    if (allProgress.history.length > 10) {
      allProgress.history = allProgress.history.slice(0, 10);
    }
    
    // 保存到 Properties
    properties.setProperty(progressKey, JSON.stringify(allProgress));
    
    Logger.log(`保存閱讀進度：${sheetName} - 品名索引 ${chapterIndex} - 行號 ${lineNumber}`);
    return { success: true, message: '進度已保存' };
  } catch (error) {
    Logger.log('保存閱讀進度錯誤：' + error.toString());
    return { success: false, message: '保存失敗：' + error.message };
  }
}

// 獲取閱讀進度
function getReadingProgress() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const properties = PropertiesService.getUserProperties();
    const progressKey = `READING_PROGRESS_${userEmail}`;
    
    const existingProgress = properties.getProperty(progressKey);
    if (existingProgress) {
      return JSON.parse(existingProgress);
    }
    
    return { current: null, history: [] };
  } catch (error) {
    Logger.log('獲取閱讀進度錯誤：' + error.toString());
    return { current: null, history: [] };
  }
}

// 清除閱讀進度
function clearReadingProgress() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const properties = PropertiesService.getUserProperties();
    const progressKey = `READING_PROGRESS_${userEmail}`;
    properties.deleteProperty(progressKey);
    return { success: true, message: '進度已清除' };
  } catch (error) {
    Logger.log('清除閱讀進度錯誤：' + error.toString());
    return { success: false, message: '清除失敗：' + error.message };
  }
}

// ==================== 書籤功能 ====================

// 添加書籤
function addBookmark(sheetName, chapterIndex, chapterTitle, lineNumber, lineText) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const properties = PropertiesService.getUserProperties();
    const bookmarkKey = `BOOKMARKS_${userEmail}`;
    
    // 獲取現有書籤
    let bookmarks = [];
    const existingBookmarks = properties.getProperty(bookmarkKey);
    if (existingBookmarks) {
      try {
        bookmarks = JSON.parse(existingBookmarks);
      } catch (e) {
        Logger.log('解析現有書籤失敗，重新初始化');
        bookmarks = [];
      }
    }
    
    // 檢查是否已存在相同位置的書籤
    const existingIndex = bookmarks.findIndex(b => 
      b.sheetName === sheetName && 
      b.chapterIndex === chapterIndex && 
      b.lineNumber === lineNumber
    );
    
    if (existingIndex >= 0) {
      return { success: false, message: '此位置已有書籤' };
    }
    
    // 創建新書籤
    const bookmark = {
      id: Date.now().toString(),
      sheetName: sheetName,
      chapterIndex: chapterIndex,
      chapterTitle: chapterTitle || '',
      lineNumber: lineNumber,
      lineText: lineText || '',
      timestamp: new Date().toISOString()
    };
    
    bookmarks.unshift(bookmark); // 新書籤放在前面
    
    // 保存書籤
    properties.setProperty(bookmarkKey, JSON.stringify(bookmarks));
    
    Logger.log(`添加書籤：${sheetName} - ${chapterTitle} - 行號 ${lineNumber}`);
    return { success: true, message: '書籤已添加', bookmark: bookmark };
  } catch (error) {
    Logger.log('添加書籤錯誤：' + error.toString());
    return { success: false, message: '添加失敗：' + error.message };
  }
}

// 獲取所有書籤
function getAllBookmarks() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const properties = PropertiesService.getUserProperties();
    const bookmarkKey = `BOOKMARKS_${userEmail}`;
    
    const existingBookmarks = properties.getProperty(bookmarkKey);
    if (existingBookmarks) {
      return JSON.parse(existingBookmarks);
    }
    
    return [];
  } catch (error) {
    Logger.log('獲取書籤錯誤：' + error.toString());
    return [];
  }
}

// 刪除書籤
function deleteBookmark(bookmarkId) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const properties = PropertiesService.getUserProperties();
    const bookmarkKey = `BOOKMARKS_${userEmail}`;
    
    // 獲取現有書籤
    let bookmarks = [];
    const existingBookmarks = properties.getProperty(bookmarkKey);
    if (existingBookmarks) {
      try {
        bookmarks = JSON.parse(existingBookmarks);
      } catch (e) {
        return { success: false, message: '書籤數據錯誤' };
      }
    }
    
    // 刪除指定的書籤
    const filteredBookmarks = bookmarks.filter(b => b.id !== bookmarkId);
    
    if (filteredBookmarks.length === bookmarks.length) {
      return { success: false, message: '找不到指定的書籤' };
    }
    
    // 保存更新後的書籤列表
    properties.setProperty(bookmarkKey, JSON.stringify(filteredBookmarks));
    
    Logger.log(`刪除書籤：${bookmarkId}`);
    return { success: true, message: '書籤已刪除' };
  } catch (error) {
    Logger.log('刪除書籤錯誤：' + error.toString());
    return { success: false, message: '刪除失敗：' + error.message };
  }
}

// 清除所有書籤
function clearAllBookmarks() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const properties = PropertiesService.getUserProperties();
    const bookmarkKey = `BOOKMARKS_${userEmail}`;
    properties.deleteProperty(bookmarkKey);
    return { success: true, message: '所有書籤已清除' };
  } catch (error) {
    Logger.log('清除書籤錯誤：' + error.toString());
    return { success: false, message: '清除失敗：' + error.message };
  }
}
  