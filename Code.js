// ===== 設定管理 =====
/**
 * スクリプトプロパティから設定を取得
 */
function getConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return {
    token: scriptProperties.getProperty('LINE_CHANNEL_ACCESS_TOKEN'),
    groupId: scriptProperties.getProperty('LINE_GROUP_ID')
  };
}

// ===== 初期設定用関数 =====
/**
 * 初期設定画面を表示
 */
function showSettingsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Settings')
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'LINE通知 初期設定');
}

/**
 * スクリプトプロパティに設定を保存（HTMLから呼び出される）
 */
function saveSettings(token, groupId) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperties({
    'LINE_CHANNEL_ACCESS_TOKEN': token,
    'LINE_GROUP_ID': groupId
  });
  return '設定を保存しました';
}

/**
 * 現在の設定を取得（HTMLから呼び出される）
 */
function getCurrentSettings() {
  const config = getConfig();
  return {
    hasToken: !!config.token,
    hasGroupId: !!config.groupId,
    tokenPreview: config.token ? config.token.substring(0, 20) + '...' : '',
    groupId: config.groupId || ''
  };
}

// ===== メイン機能 =====
/**
 * LINEグループにメッセージを送信する関数
 * @param {string} message - 送信するメッセージ
 */
function sendLineNotification(message) {
  const config = getConfig();
  
  if (!config.token || !config.groupId) {
    throw new Error('LINE設定が見つかりません。メニューから「初期設定」を実行してください。');
  }
  
  const url = 'https://api.line.me/v2/bot/message/push';
  
  const payload = {
    to: config.groupId,
    messages: [{
      type: 'text',
      text: message
    }]
  };
  
  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + config.token
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      console.log('送信成功');
      return { success: true };
    } else {
      const responseText = response.getContentText();
      console.error('送信エラー:', responseCode, responseText);
      return { success: false, error: `エラーコード: ${responseCode}` };
    }
  } catch (error) {
    console.error('送信エラー:', error);
    return { success: false, error: error.toString() };
  }
}


// ===== フォーム/スプレッドシート連携 =====
/**
 * Google Formの送信時に実行される関数
 * @param {Object} e - イベントオブジェクト
 */
function onFormSubmit(e) {
  const responses = e.values;
  const timestamp = responses[0];
  
  let message = `📝 新しいフォーム回答\n`;
  message += `━━━━━━━━━━━━\n`;
  message += `送信日時: ${timestamp}\n\n`;
  
  // フォームの質問と回答を整形
  const sheet = e.range.getSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  for (let i = 1; i < responses.length; i++) {
    if (headers[i] && responses[i]) {
      message += `【${headers[i]}】\n${responses[i]}\n\n`;
    }
  }
  
  sendLineNotification(message);
}

/**
 * カスタムメッセージを送信（HTMLダイアログから呼び出される）
 */
function sendCustomNotification(title, message, options = {}) {
  let formattedMessage = '';
  
  // 絵文字の選択
  const emoji = options.emoji || '📢';
  
  formattedMessage += `${emoji} ${title}\n`;
  formattedMessage += `━━━━━━━━━━━━\n`;
  formattedMessage += message;
  
  // 送信者情報を追加する場合
  if (options.includeSender) {
    const userEmail = Session.getActiveUser().getEmail();
    formattedMessage += `\n\n送信者: ${userEmail}`;
  }
  
  // タイムスタンプを追加する場合
  if (options.includeTimestamp) {
    formattedMessage += `\n送信日時: ${new Date().toLocaleString('ja-JP')}`;
  }
  
  return sendLineNotification(formattedMessage);
}

// ===== UI関連 =====

/**
 * メニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📱 LINE通知')
    .addItem('📊 昨日のPL情報を送信', 'sendYesterdayPLReport')
    .addItem('📅 指定期間のPL情報を送信', 'sendPLReportWithPeriod')
    .addItem('💰 現預金残高を送信', 'sendCashBalanceReport')
    .addSeparator()
    .addItem('⚙️ 初期設定', 'showSettingsDialog')
    .addItem('🔍 設定確認', 'checkSettings')
    .addToUi();
}

/**
 * 設定を確認
 */
function checkSettings() {
  const config = getConfig();
  
  let message = '=== 現在の設定 ===\n\n';
  
  if (config.token) {
    message += '✅ アクセストークン: 設定済み\n';
    message += `   (${config.token.substring(0, 20)}...)\n\n`;
  } else {
    message += '❌ アクセストークン: 未設定\n\n';
  }
  
  if (config.groupId) {
    message += '✅ グループID: 設定済み\n';
    message += `   (${config.groupId})\n`;
  } else {
    message += '❌ グループID: 未設定\n';
  }
  
  if (!config.token || !config.groupId) {
    message += '\n⚠️ 初期設定を完了してください';
  }
  
  SpreadsheetApp.getUi().alert(message);
}

// ===== Webhook関連（グループID取得用） =====
/**
 * WebhookでグループIDを取得するための関数
 */
function doPost(e) {
  if (!e.postData || !e.postData.contents) {
    return ContentService.createTextOutput(JSON.stringify({'status': 'no data'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  try {
    const json = JSON.parse(e.postData.contents);
    const events = json.events || [];
    
    events.forEach(event => {
      // グループからのメッセージの場合
      if (event.type === 'message' && event.source.type === 'group') {
        const groupId = event.source.groupId;
        console.log('グループID検出:', groupId);
        
        // グループIDを自動的に保存（オプション）
        // saveDetectedGroupId(groupId);
        
        // 管理者に通知（オプション）
        notifyGroupIdDetected(groupId);
      }
    });
  } catch (error) {
    console.error('Webhook処理エラー:', error);
  }
  
  return ContentService.createTextOutput(JSON.stringify({'status': 'ok'}))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 検出したグループIDを通知
 */
function notifyGroupIdDetected(groupId) {
  // スプレッドシートに記録
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GroupID_Log') 
    || SpreadsheetApp.getActiveSpreadsheet().insertSheet('GroupID_Log');
  
  sheet.appendRow([
    new Date(),
    'Group ID Detected',
    groupId
  ]);
}

// ===== 全社PLシート連携機能 =====

/**
 * 昨日のPLデータを取得してLINE通知
 */
function sendYesterdayPLReport() {
  try {
    // 昨日の日付を取得
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const yesterdayString = formatDateString(yesterday);
    
    const plData = getPLData(yesterdayString);
    
    if (!plData) {
      SpreadsheetApp.getUi().alert('昨日（' + yesterdayString + '）のデータが見つかりませんでした。');
      return;
    }
    
    // メッセージを作成
    const message = formatPLMessage(plData);
    
    // LINE送信
    const result = sendLineNotification(message);
    
    if (result.success) {
      SpreadsheetApp.getUi().alert('✅ 昨日のPL情報を送信しました\n\nLINEグループを確認してください。');
    } else {
      SpreadsheetApp.getUi().alert('❌ 送信失敗\n\n' + result.error);
    }
  } catch (error) {
    console.error('PLレポート送信エラー:', error);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.toString());
  }
}

/**
 * 指定期間のPLデータを送信（カレンダーダイアログで期間選択）
 */
function sendPLReportWithPeriod() {
  const html = HtmlService.createHtmlOutputFromFile('PeriodSelector')
    .setWidth(400)
    .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, '期間を選択');
}

/**
 * 期間指定でPLデータを取得して送信（HTMLダイアログから呼び出される）
 */
function sendPLReportForPeriod(startDate, endDate) {
  try {
    const start = new Date(startDate);
    const end = new Date(endDate);
    
    if (start > end) {
      return { success: false, error: '開始日は終了日より前の日付を選択してください。' };
    }
    
    // 期間が長すぎる場合の警告（開始日と終了日を含む日数を計算）
    const daysDiff = Math.floor((end - start) / (1000 * 60 * 60 * 24)) + 1;
    if (daysDiff > 31) {
      return { success: false, error: '期間は31日以内で指定してください。' };
    }
    
    // 期間内のPLデータを集計
    const periodData = getPLDataForPeriod(startDate, endDate);
    
    if (!periodData || periodData.length === 0) {
      return { success: false, error: '指定期間のデータが見つかりませんでした。' };
    }
    
    // メッセージを作成
    const message = formatPeriodPLMessage(periodData, startDate, endDate);
    
    // LINE送信
    const result = sendLineNotification(message);
    
    if (result.success) {
      return { success: true };
    } else {
      return { success: false, error: result.error };
    }
  } catch (error) {
    console.error('期間PLレポート送信エラー:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * 指定期間のPLデータを取得
 * @param {string} startDate - 開始日
 * @param {string} endDate - 終了日
 * @return {Array} PLデータの配列
 */
function getPLDataForPeriod(startDate, endDate) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const plSheet = spreadsheet.getSheetByName('全社PLシート');
  
  if (!plSheet) {
    throw new Error('「全社PLシート」が見つかりません');
  }
  
  // 開始日と終了日を0時0分0秒にリセット（時刻の影響を排除）
  const start = new Date(startDate);
  start.setHours(0, 0, 0, 0);
  
  const end = new Date(endDate);
  end.setHours(23, 59, 59, 999); // 終了日の最後の瞬間まで含む
  
  const results = [];
  
  // A6から最終行までのデータを取得
  const lastRow = plSheet.getLastRow();
  const dataRange = plSheet.getRange(6, 1, lastRow - 5, 12).getValues();
  
  console.log('期間検索:', formatDateString(start), '～', formatDateString(end));
  
  for (let i = 0; i < dataRange.length; i++) {
    const cellDate = dataRange[i][0];
    if (!cellDate) continue;
    
    const rowDate = new Date(cellDate);
    rowDate.setHours(0, 0, 0, 0); // 比較用に時刻をリセット
    
    // 日付が期間内かチェック（開始日と終了日を含む）
    if (rowDate >= start && rowDate <= end) {
      const plData = {
        date: formatDateString(rowDate),
        dailySales: dataRange[i][2],        // C列
        monthlySales: dataRange[i][3],      // D列
        dailyGrossProfit: dataRange[i][6],  // G列
        monthlyGrossProfit: dataRange[i][7], // H列
        dailyOperatingProfit: dataRange[i][10], // K列
        monthlyOperatingProfit: dataRange[i][11] // L列
      };
      results.push(plData);
      console.log('データ追加:', plData.date);
    }
  }
  
  console.log('取得データ数:', results.length);
  return results;
}

/**
 * 期間PLデータをLINEメッセージ形式にフォーマット
 * @param {Array} periodData - 期間のPLデータ配列
 * @param {string} startDate - 開始日
 * @param {string} endDate - 終了日
 * @return {string} フォーマットされたメッセージ
 */
function formatPeriodPLMessage(periodData, startDate, endDate) {
  const formatCurrency = (num) => {
    if (typeof num !== 'number') return num;
    return '¥' + num.toLocaleString('ja-JP');
  };
  
  // 期間の合計を計算
  let totalSales = 0;
  let totalGrossProfit = 0;
  let totalOperatingProfit = 0;
  
  periodData.forEach(data => {
    totalSales += data.dailySales || 0;
    totalGrossProfit += data.dailyGrossProfit || 0;
    totalOperatingProfit += data.dailyOperatingProfit || 0;
  });
  
  // 実際のデータ件数を表示
  let message = `📊 期間PL情報\n`;
  message += `━━━━━━━━━━━━\n`;
  message += `📅 ${startDate} ～ ${endDate}\n`;
  message += `（データ件数: ${periodData.length}件）\n\n`;
  
  message += `【期間合計】\n`;
  message += `売上: ${formatCurrency(totalSales)}\n`;
  message += `粗利: ${formatCurrency(totalGrossProfit)}\n`;
  message += `営業利益: ${formatCurrency(totalOperatingProfit)}\n\n`;
  
  message += `【日平均】\n`;
  message += `売上: ${formatCurrency(Math.round(totalSales / periodData.length))}\n`;
  message += `粗利: ${formatCurrency(Math.round(totalGrossProfit / periodData.length))}\n`;
  message += `営業利益: ${formatCurrency(Math.round(totalOperatingProfit / periodData.length))}`;
  
  return message;
}

/**
 * 全社PLシートからデータを取得
 * @param {string} targetDate - 対象日付（省略時はA1の日付を使用）
 * @return {Object|null} PLデータオブジェクト
 */
function getPLData(targetDate = null) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const plSheet = spreadsheet.getSheetByName('全社PLシート');
  
  if (!plSheet) {
    throw new Error('「全社PLシート」が見つかりません');
  }
  
  // A1の基準日付を取得（targetDateが指定されていない場合）
  const baseDateCell = plSheet.getRange('A1').getValue();
  const searchDate = targetDate || formatDateString(baseDateCell);
  
  // A6から最終行までの日付を取得
  const lastRow = plSheet.getLastRow();
  const dateColumn = plSheet.getRange(6, 1, lastRow - 5, 1).getValues();
  
  // 検索日付と一致する行を探す
  let targetRow = -1;
  for (let i = 0; i < dateColumn.length; i++) {
    const rowDate = formatDateString(dateColumn[i][0]);
    if (rowDate === searchDate) {
      targetRow = i + 6; // 実際の行番号（6行目から開始）
      break;
    }
  }
  
  if (targetRow === -1) {
    return null;
  }
  
  // 該当行のデータを取得（C, D, G, H, K, L列）
  const rowData = plSheet.getRange(targetRow, 3, 1, 10).getValues()[0];
  
  return {
    date: searchDate,
    dailySales: rowData[0],        // C列（3列目）
    monthlySales: rowData[1],      // D列（4列目）
    dailyGrossProfit: rowData[4],  // G列（7列目）
    monthlyGrossProfit: rowData[5], // H列（8列目）
    dailyOperatingProfit: rowData[8], // K列（11列目）
    monthlyOperatingProfit: rowData[9] // L列（12列目）
  };
}

/**
 * 日付を統一形式（yyyy/mm/dd）に変換
 * @param {Date|string} dateValue - 日付値
 * @return {string} フォーマットされた日付文字列
 */
function formatDateString(dateValue) {
  if (!dateValue) return '';
  
  let date;
  if (dateValue instanceof Date) {
    date = dateValue;
  } else {
    date = new Date(dateValue);
  }
  
  if (isNaN(date.getTime())) {
    return String(dateValue); // 変換できない場合は元の値を返す
  }
  
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  
  return `${year}/${month}/${day}`;
}

/**
 * PLデータをLINEメッセージ形式にフォーマット
 * @param {Object} plData - PLデータオブジェクト
 * @return {string} フォーマットされたメッセージ
 */
function formatPLMessage(plData) {
  const formatNumber = (num) => {
    if (typeof num !== 'number') return num;
    return num.toLocaleString('ja-JP');
  };
  
  const formatCurrency = (num) => {
    if (typeof num !== 'number') return num;
    return '¥' + num.toLocaleString('ja-JP');
  };
  
  let message = `📊 全社PL情報\n`;
  message += `━━━━━━━━━━━━\n`;
  message += `📅 ${plData.date}\n\n`;
  
  message += `【売上】\n`;
  message += `単日: ${formatCurrency(plData.dailySales)}\n`;
  message += `当月累計: ${formatCurrency(plData.monthlySales)}\n\n`;
  
  message += `【粗利】\n`;
  message += `単日: ${formatCurrency(plData.dailyGrossProfit)}\n`;
  message += `当月累計: ${formatCurrency(plData.monthlyGrossProfit)}\n\n`;
  
  message += `【営業利益】\n`;
  message += `単日: ${formatCurrency(plData.dailyOperatingProfit)}\n`;
  message += `当月累計: ${formatCurrency(plData.monthlyOperatingProfit)}`;
  
  return message;
}

/**
 * 毎日定時にPLレポートを自動送信するためのトリガー関数（昨日のデータを送信）
 */
function sendDailyPLReportAuto() {
  try {
    // 昨日の日付を取得
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const yesterdayString = formatDateString(yesterday);
    
    const plData = getPLData(yesterdayString);
    
    if (!plData) {
      console.log('昨日のPLデータが見つかりません');
      return;
    }
    
    const message = formatPLMessage(plData);
    const result = sendLineNotification(message);
    
    if (result.success) {
      console.log('PLレポート送信成功:', new Date());
    } else {
      console.error('PLレポート送信失敗:', result.error);
    }
  } catch (error) {
    console.error('自動PLレポート送信エラー:', error);
  }
}

// ===== 現預金残高機能 =====

/**
 * 現預金残高を送信（前日の実残高と翌月末の予算残高）
 */
function sendCashBalanceReport() {
  try {
    // 昨日の日付を取得
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    
    // 翌月末の日付を取得
    const nextMonthEnd = new Date();
    nextMonthEnd.setMonth(nextMonthEnd.getMonth() + 2, 0); // 翌月の最終日
    
    // 現預金データを取得
    const cashData = getCashBalanceData(yesterday, nextMonthEnd);
    
    if (!cashData) {
      SpreadsheetApp.getUi().alert('現預金データが見つかりませんでした。CFシートを確認してください。');
      return;
    }
    
    // メッセージを作成
    const message = formatCashBalanceMessage(cashData);
    
    // LINE送信
    const result = sendLineNotification(message);
    
    if (result.success) {
      SpreadsheetApp.getUi().alert('✅ 現預金残高情報を送信しました\n\nLINEグループを確認してください。');
    } else {
      SpreadsheetApp.getUi().alert('❌ 送信失敗\n\n' + result.error);
    }
  } catch (error) {
    console.error('現預金残高送信エラー:', error);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.toString());
  }
}

/**
 * 現預金残高データを取得
 * @param {Date} actualDate - 実残高の日付（前日）
 * @param {Date} budgetDate - 予算残高の日付（翌月末）
 * @return {Object|null} 現預金データオブジェクト
 */
function getCashBalanceData(actualDate, budgetDate) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 実残高のシート名を決定（例: 202508CF）
  const actualSheetName = formatCFSheetName(actualDate);
  const actualSheet = spreadsheet.getSheetByName(actualSheetName);
  
  // 予算残高のシート名を決定（例: 202509CF）
  const budgetSheetName = formatCFSheetName(budgetDate);
  const budgetSheet = spreadsheet.getSheetByName(budgetSheetName);
  
  if (!actualSheet && !budgetSheet) {
    console.error('CFシートが見つかりません:', actualSheetName, budgetSheetName);
    return null;
  }
  
  const data = {
    actualDate: formatDateString(actualDate),
    budgetDate: formatDateString(budgetDate),
    actualBalance: null,
    budgetBalance: null,
    actualSheetName: actualSheetName,
    budgetSheetName: budgetSheetName
  };
  
  // 実残高を取得（前日のデータ）
  if (actualSheet) {
    const actualData = findCashBalanceInSheet(actualSheet, actualDate);
    if (actualData) {
      data.actualBalance = actualData.actualBalance;
    }
  }
  
  // 予算残高を取得（翌月末のデータ）
  if (budgetSheet) {
    const budgetData = findCashBalanceInSheet(budgetSheet, budgetDate);
    if (budgetData) {
      data.budgetBalance = budgetData.budgetBalance;
    }
  }
  
  return data;
}

/**
 * CFシート内から指定日付の現預金残高を検索
 * @param {Sheet} sheet - 検索対象のシート
 * @param {Date} targetDate - 検索する日付
 * @return {Object|null} 残高データ
 */
function findCashBalanceInSheet(sheet, targetDate) {
  // 7行目から日付を探す（最大100列まで検索）
  const dateRow = sheet.getRange(7, 1, 1, 100).getValues()[0];
  const budgetRow = sheet.getRange(8, 1, 1, 100).getValues()[0];
  const actualRow = sheet.getRange(9, 1, 1, 100).getValues()[0];
  
  // 日付を検索
  for (let i = 0; i < dateRow.length; i++) {
    if (!dateRow[i]) continue;
    
    const cellDate = new Date(dateRow[i]);
    // 日付の比較（日単位で一致するか確認）
    if (formatDateString(cellDate) === formatDateString(targetDate)) {
      return {
        date: formatDateString(targetDate),
        budgetBalance: budgetRow[i],
        actualBalance: actualRow[i]
      };
    }
  }
  
  return null;
}

/**
 * CFシート名をフォーマット（例: 202508CF）
 * @param {Date} date - 日付
 * @return {string} シート名
 */
function formatCFSheetName(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  return `${year}${month}CF`;
}

/**
 * 現預金残高メッセージをフォーマット
 * @param {Object} cashData - 現預金データ
 * @return {string} フォーマットされたメッセージ
 */
function formatCashBalanceMessage(cashData) {
  const formatCurrency = (num) => {
    if (typeof num !== 'number') return '取得不可';
    return '¥' + num.toLocaleString('ja-JP');
  };
  
  let message = `💰 現預金残高情報\n`;
  message += `━━━━━━━━━━━━\n\n`;
  
  message += `【実残高】\n`;
  message += `📅 ${cashData.actualDate}時点\n`;
  message += `💵 ${formatCurrency(cashData.actualBalance)}\n`;
  message += `（${cashData.actualSheetName}より取得）\n\n`;
  
  message += `【予算残高】\n`;
  message += `📅 ${cashData.budgetDate}時点\n`;
  message += `💴 ${formatCurrency(cashData.budgetBalance)}\n`;
  message += `（${cashData.budgetSheetName}より取得）\n\n`;
  
  // 差額を計算
  if (typeof cashData.actualBalance === 'number' && typeof cashData.budgetBalance === 'number') {
    const difference = cashData.budgetBalance - cashData.actualBalance;
    const sign = difference >= 0 ? '📈' : '📉';
    message += `【予実差額】\n`;
    message += `${sign} ${formatCurrency(difference)}`;
  }
  
  return message;
}

/**
 * 毎日定時に現預金残高を自動送信するためのトリガー関数
 */
function sendDailyCashBalanceAuto() {
  try {
    // 昨日の日付を取得
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    
    // 翌月末の日付を取得
    const nextMonthEnd = new Date();
    nextMonthEnd.setMonth(nextMonthEnd.getMonth() + 2, 0);
    
    const cashData = getCashBalanceData(yesterday, nextMonthEnd);
    
    if (!cashData || (!cashData.actualBalance && !cashData.budgetBalance)) {
      console.log('現預金データが見つかりません');
      return;
    }
    
    const message = formatCashBalanceMessage(cashData);
    const result = sendLineNotification(message);
    
    if (result.success) {
      console.log('現預金残高送信成功:', new Date());
    } else {
      console.error('現預金残高送信失敗:', result.error);
    }
  } catch (error) {
    console.error('自動現預金残高送信エラー:', error);
  }
}

// ===== デバッグ用関数 =====

/**
 * 期間のデータ取得をテスト（デバッグ用）
 */
function debugTestPeriodData() {
  // 2025年6月のデータを取得してテスト
  const startDate = '2025-06-01';
  const endDate = '2025-06-30';
  
  console.log('=== 期間データ取得テスト ===');
  console.log('期間:', startDate, '～', endDate);
  
  const periodData = getPLDataForPeriod(startDate, endDate);
  
  console.log('取得データ数:', periodData.length);
  console.log('最初のデータ:', periodData[0]);
  console.log('最後のデータ:', periodData[periodData.length - 1]);
  
  // 6月は30日あるはずなので確認
  if (periodData.length !== 30) {
    console.warn('警告: 6月のデータが30件ではありません。実際:', periodData.length);
  }
  
  // 日付の連続性をチェック
  const dates = periodData.map(d => d.date);
  console.log('取得された日付:', dates);
}

/**
 * 詳細なデバッグ情報付きでLINE送信をテスト
 */
function debugLineNotification() {
  console.log('=== デバッグ開始 ===');
  
  // 1. 設定の取得と確認
  console.log('1. 設定を取得中...');
  const config = getConfig();
  
  console.log('トークン:', config.token ? `${config.token.substring(0, 30)}...` : '未設定');
  console.log('グループID:', config.groupId || '未設定');
  
  if (!config.token || !config.groupId) {
    console.error('エラー: 設定が不完全です');
    SpreadsheetApp.getUi().alert('設定が不完全です。初期設定を行ってください。');
    return;
  }
  
  // 2. メッセージの準備
  console.log('2. メッセージを準備中...');
  const testMessage = '🧪 デバッグテスト送信\n' + new Date().toLocaleString('ja-JP');
  console.log('送信メッセージ:', testMessage);
  
  // 3. APIリクエストの準備
  console.log('3. APIリクエストを準備中...');
  const url = 'https://api.line.me/v2/bot/message/push';
  
  const payload = {
    to: config.groupId,
    messages: [{
      type: 'text',
      text: testMessage
    }]
  };
  console.log('ペイロード:', JSON.stringify(payload, null, 2));
  
  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + config.token
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  // 4. API呼び出し
  console.log('4. LINE APIを呼び出し中...');
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    const responseHeaders = response.getHeaders();
    
    console.log('レスポンスコード:', responseCode);
    console.log('レスポンスヘッダー:', JSON.stringify(responseHeaders, null, 2));
    console.log('レスポンス本文:', responseText);
    
    if (responseCode === 200) {
      console.log('✅ 送信成功！');
      SpreadsheetApp.getUi().alert('送信成功！LINEグループを確認してください。');
    } else {
      console.error('❌ 送信失敗');
      let errorMessage = `エラーコード: ${responseCode}\n`;
      
      try {
        const errorDetail = JSON.parse(responseText);
        errorMessage += `エラーメッセージ: ${errorDetail.message}\n`;
        if (errorDetail.details) {
          errorMessage += `詳細: ${JSON.stringify(errorDetail.details, null, 2)}`;
        }
      } catch (e) {
        errorMessage += `レスポンス: ${responseText}`;
      }
      
      console.error('エラー詳細:', errorMessage);
      SpreadsheetApp.getUi().alert('送信失敗\n\n' + errorMessage);
    }
  } catch (error) {
    console.error('例外エラー:', error.toString());
    console.error('スタックトレース:', error.stack);
    SpreadsheetApp.getUi().alert('エラーが発生しました:\n' + error.toString());
  }
  
  console.log('=== デバッグ終了 ===');
}

/**
 * 設定値を詳細に確認
 */
function debugCheckSettings() {
  const config = getConfig();
  const scriptProperties = PropertiesService.getScriptProperties();
  const allProperties = scriptProperties.getProperties();
  
  console.log('=== 設定値の確認 ===');
  console.log('全プロパティ:', JSON.stringify(allProperties, null, 2));
  
  // トークンの検証
  if (config.token) {
    console.log('トークン長:', config.token.length);
    console.log('トークン先頭:', config.token.substring(0, 20));
    console.log('トークン末尾:', config.token.substring(config.token.length - 20));
  }
  
  // グループIDの検証
  if (config.groupId) {
    console.log('グループID:', config.groupId);
    console.log('グループID長:', config.groupId.length);
    console.log('グループID先頭文字:', config.groupId.charAt(0));
  }
  
  let message = '=== 詳細な設定情報 ===\n\n';
  message += `トークン: ${config.token ? '設定済み' : '未設定'}\n`;
  if (config.token) {
    message += `  - 長さ: ${config.token.length}文字\n`;
    message += `  - 先頭: ${config.token.substring(0, 20)}...\n`;
  }
  message += `\nグループID: ${config.groupId || '未設定'}\n`;
  if (config.groupId) {
    message += `  - 値: ${config.groupId}\n`;
    message += `  - 長さ: ${config.groupId.length}文字\n`;
  }
  
  SpreadsheetApp.getUi().alert(message);
}

/**
 * LINE APIの接続テスト（認証のみ）
 */
function testLineAPIConnection() {
  console.log('=== LINE API接続テスト ===');
  
  const config = getConfig();
  if (!config.token) {
    console.error('トークンが設定されていません');
    SpreadsheetApp.getUi().alert('トークンが設定されていません');
    return;
  }
  
  // Bot情報を取得するAPIを呼び出し
  const url = 'https://api.line.me/v2/bot/info';
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + config.token
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('レスポンスコード:', responseCode);
    console.log('レスポンス:', responseText);
    
    if (responseCode === 200) {
      const botInfo = JSON.parse(responseText);
      SpreadsheetApp.getUi().alert(
        '✅ API接続成功！\n\n' +
        `Bot名: ${botInfo.displayName}\n` +
        `Bot ID: ${botInfo.userId}`
      );
    } else {
      SpreadsheetApp.getUi().alert(
        '❌ API接続失敗\n\n' +
        `エラーコード: ${responseCode}\n` +
        `詳細: ${responseText}`
      );
    }
  } catch (error) {
    console.error('エラー:', error);
    SpreadsheetApp.getUi().alert('エラー: ' + error.toString());
  }
}