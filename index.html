/**
 * Gmail → Notion 自動連携ツール
 * 
 * 機能：
 * - Gmailの特定ラベル付きメールをNotionデータベースに自動追加
 * - ユーザーフレンドリーなメニューインターフェース
 * - 設定とログの管理
 * - 自動実行トリガー設定
 */

// ==================== メニュー設定 ====================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📧 Gmail → Notion')
    .addItem('🔧 初期設定', 'initializeSheets')
    .addItem('🔗 接続テスト', 'testConnection')
    .addItem('▶️ 手動実行', 'manualExecution')
    .addItem('⏰ 自動実行設定', 'setupAutoTrigger')
    .addItem('⏹️ 自動実行停止', 'stopAutoTrigger')
    .addItem('📊 ログ確認', 'showLogStats')
    .addSeparator()
    .addItem('🔍 デバッグ情報', 'debugLabelStatus')
    .addToUi();
}

// ==================== 初期設定 ====================
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // 設定シートの作成
    let settingsSheet = ss.getSheetByName('設定');
    if (!settingsSheet) {
      settingsSheet = ss.insertSheet('設定');
    }
    
    // 設定シートのヘッダー設定
    settingsSheet.clear();
    settingsSheet.getRange('A1:B1').setValues([['設定項目', '値']]);
    settingsSheet.getRange('A2:B9').setValues([
      ['NotionデータベースID', ''],
      ['Notionシークレットトークン', ''],
      ['Notionタイトルプロパティ名', 'Name'],
      ['Gmailラベル名', ''],
      ['処理後の動作', 'ラベル削除'], // ラベル削除/完了ラベル追加/アーカイブ
      ['完了ラベル名', 'Notion処理済み'],
      ['最終処理日時', ''],
      ['処理間隔（分）', '60']
    ]);
    
    // 設定シートの書式設定
    settingsSheet.getRange('A1:B1').setFontWeight('bold').setBackground('#e6f3ff');
    settingsSheet.getRange('A2:A9').setFontWeight('bold');
    settingsSheet.setColumnWidth(1, 200);
    settingsSheet.setColumnWidth(2, 300);
    
    // ログシートの作成
    let logSheet = ss.getSheetByName('ログ');
    if (!logSheet) {
      logSheet = ss.insertSheet('ログ');
    }
    
    // ログシートのヘッダー設定
    logSheet.clear();
    logSheet.getRange('A1:F1').setValues([['処理日時', 'メール件名', '送信者', 'ステータス', 'エラー内容', '処理時間(秒)']]);
    logSheet.getRange('A1:F1').setFontWeight('bold').setBackground('#e6f3ff');
    logSheet.setColumnWidth(1, 150);
    logSheet.setColumnWidth(2, 300);
    logSheet.setColumnWidth(3, 200);
    logSheet.setColumnWidth(4, 100);
    logSheet.setColumnWidth(5, 300);
    logSheet.setColumnWidth(6, 100);
    
    // 統計シートの作成
    let statsSheet = ss.getSheetByName('統計');
    if (!statsSheet) {
      statsSheet = ss.insertSheet('統計');
    }
    
    // 統計シートのヘッダー設定
    statsSheet.clear();
    statsSheet.getRange('A1:B1').setValues([['項目', '値']]);
    statsSheet.getRange('A2:B8').setValues([
      ['総処理件数', '0'],
      ['成功件数', '0'],
      ['エラー件数', '0'],
      ['成功率', '0%'],
      ['最新処理日時', ''],
      ['平均処理時間', '0秒'],
      ['今日の処理件数', '0']
    ]);
    
    statsSheet.getRange('A1:B1').setFontWeight('bold').setBackground('#e6f3ff');
    statsSheet.getRange('A2:A8').setFontWeight('bold');
    statsSheet.setColumnWidth(1, 200);
    statsSheet.setColumnWidth(2, 200);
    
    // 成功メッセージ
    SpreadsheetApp.getUi().alert(
      '✅ 初期設定完了',
      '設定シート、ログシート、統計シートが作成されました。\n\n次のステップ：\n1. 設定シートに必要な情報を入力してください\n2. 接続テストを実行してください',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    // 設定シートをアクティブにする
    ss.setActiveSheet(settingsSheet);
    
  } catch (error) {
    console.error(`初期設定エラー: ${error.toString()}`);
    SpreadsheetApp.getUi().alert('❌ エラー', '初期設定中にエラーが発生しました：\n' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ==================== 設定読み込み ====================
function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const settingsSheet = ss.getSheetByName('設定');
    
    if (!settingsSheet) {
      throw new Error('設定シートが見つかりません。初期設定を実行してください。');
    }
    
    const values = settingsSheet.getRange('B2:B9').getValues();
    
    return {
      databaseId: values[0][0],
      notionToken: values[1][0],
      titleProperty: values[2][0] || 'Name',
      gmailLabel: values[3][0],
      postProcessAction: values[4][0] || 'ラベル削除',
      completedLabel: values[5][0] || 'Notion処理済み',
      lastProcessTime: values[6][0],
      intervalMinutes: parseInt(values[7][0]) || 60
    };
  } catch (error) {
    throw new Error('設定シートの読み込みに失敗しました：' + error.toString());
  }
}

// ==================== 設定値検証 ====================
function validateSettings(settings) {
  const errors = [];
  
  if (!settings.databaseId) {
    errors.push('NotionデータベースIDが設定されていません');
  }
  
  if (!settings.notionToken) {
    errors.push('Notionシークレットトークンが設定されていません');
  }
  
  if (!settings.titleProperty) {
    errors.push('Notionタイトルプロパティ名が設定されていません');
  }
  
  if (!settings.gmailLabel) {
    errors.push('Gmailラベル名が設定されていません');
  }
  
  if (settings.intervalMinutes < 1 || settings.intervalMinutes > 1440) {
    errors.push('処理間隔は1分から1440分（24時間）の間で設定してください');
  }
  
  return errors;
}

// ==================== 接続テスト ====================
function testConnection() {
  const startTime = new Date();
  
  try {
    const settings = getSettings();
    
    // 設定値の検証
    const validationErrors = validateSettings(settings);
    if (validationErrors.length > 0) {
      SpreadsheetApp.getUi().alert(
        '⚠️ 設定不備',
        '以下の設定項目を確認してください：\n' + validationErrors.join('\n'),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Notion API テスト
    const response = UrlFetchApp.fetch(`https://api.notion.com/v1/databases/${settings.databaseId}`, {
      'method': 'GET',
      'headers': {
        'Authorization': `Bearer ${settings.notionToken}`,
        'Notion-Version': '2022-06-28'
      }
    });
    
    if (response.getResponseCode() === 200) {
      const databaseData = JSON.parse(response.getContentText());
      
      // タイトルプロパティの存在確認
      const properties = databaseData.properties;
      const titlePropertyExists = properties[settings.titleProperty] && properties[settings.titleProperty].type === 'title';
      
      if (!titlePropertyExists) {
        SpreadsheetApp.getUi().alert(
          '⚠️ タイトルプロパティエラー',
          `指定されたタイトルプロパティ「${settings.titleProperty}」が見つかりません。\n\n利用可能なプロパティ:\n${Object.keys(properties).map(key => `・${key} (${properties[key].type})`).join('\n')}`,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
        return;
      }
      
      // Gmail ラベルテスト
      const labelThreads = GmailApp.search(`label:${settings.gmailLabel}`);
      
      const processingTime = (new Date() - startTime) / 1000;
      
      SpreadsheetApp.getUi().alert(
        '✅ 接続テスト成功',
        `Notion接続成功！\nデータベース名: ${databaseData.title[0].text.content}\nタイトルプロパティ: ${settings.titleProperty} ✓\n\nGmailラベル「${settings.gmailLabel}」のメール数: ${labelThreads.length}件\n\n処理時間: ${processingTime}秒`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
      logSuccess('接続テスト', '接続テスト成功', '', processingTime);
      
    } else {
      throw new Error(`Notion API エラー: ${response.getResponseCode()} - ${response.getContentText()}`);
    }
    
  } catch (error) {
    const processingTime = (new Date() - startTime) / 1000;
    logError('接続テスト', error.toString(), '', processingTime);
    
    SpreadsheetApp.getUi().alert(
      '❌ 接続テスト失敗',
      'エラー詳細：\n' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ==================== 手動実行 ====================
function manualExecution() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '手動実行確認',
    '指定されたラベルのメールをNotionに転送します。\n処理後、該当メールからラベルが削除されます。\n\n実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    processGmailToNotion();
  }
}

// ==================== メイン処理 ====================
function processGmailToNotion() {
  const startTime = new Date();
  
  try {
    const settings = getSettings();
    
    // 設定値の検証
    const validationErrors = validateSettings(settings);
    if (validationErrors.length > 0) {
      throw new Error('設定が不完全です：' + validationErrors.join(', '));
    }
    
    // 指定ラベルのメールを取得
    const threads = GmailApp.search(`label:${settings.gmailLabel}`);
    let processedCount = 0;
    let errorCount = 0;
    const errors = [];
    
    console.log(`処理対象メール数: ${threads.length}`);
    
    // 大量メール処理時のタイムアウト対策
    const maxProcessTime = 5 * 60 * 1000; // 5分
    const maxEmailsPerRun = 50; // 1回の実行で処理する最大メール数
    
    const threadsToProcess = threads.slice(0, maxEmailsPerRun);
    
    for (const thread of threadsToProcess) {
      // 実行時間チェック
      if (new Date() - startTime > maxProcessTime) {
        console.log('実行時間制限のため処理を中断します');
        break;
      }
      
      const messages = thread.getMessages();
      let threadProcessed = false;
      
      for (const message of messages) {
        try {
          // メール情報の取得
          const subject = message.getSubject();
          const body = message.getPlainBody();
          const sender = message.getFrom();
          const date = message.getDate();
          const messageId = message.getId();
          
          // 重複チェック（簡易版）
          if (isDuplicateEmail(messageId)) {
            console.log(`重複メールをスキップ: ${subject}`);
            continue;
          }
          
          // Notionに送信
          const success = sendToNotion(settings, subject, body, sender, date, messageId);
          
          if (success) {
            logSuccess(subject, 'Notion追加成功', sender, (new Date() - startTime) / 1000);
            processedCount++;
            threadProcessed = true;
          } else {
            errorCount++;
            errors.push(`${subject}: Notion送信失敗`);
          }
          
        } catch (error) {
          console.error(`メール処理エラー: ${error.toString()}`);
          logError(message.getSubject(), error.toString(), message.getFrom(), (new Date() - startTime) / 1000);
          errorCount++;
          errors.push(`${message.getSubject()}: ${error.toString()}`);
        }
      }
      
      // スレッド単位で処理後の動作を実行（1回のみ）
      if (threadProcessed) {
        handlePostProcessAction(thread, settings);
      }
    }
    
    const totalProcessingTime = (new Date() - startTime) / 1000;
    
    // 処理結果をログに記録
    logSuccess('処理完了', `成功: ${processedCount}件, エラー: ${errorCount}件`, '', totalProcessingTime);
    
    // 最終処理日時を更新
    updateLastProcessTime();
    
    // 統計を更新
    updateStatistics(processedCount, errorCount, totalProcessingTime);
    
    console.log(`処理完了 - 成功: ${processedCount}件, エラー: ${errorCount}件, 処理時間: ${totalProcessingTime}秒`);
    
    // エラーがある場合は通知
    if (errorCount > 0 && errors.length > 0) {
      console.warn('処理中にエラーが発生しました:', errors.slice(0, 5).join('\n')); // 最初の5件のみ表示
    }
    
  } catch (error) {
    const processingTime = (new Date() - startTime) / 1000;
    console.error(`メイン処理エラー: ${error.toString()}`);
    logError('メイン処理', error.toString(), '', processingTime);
  }
}

// ==================== 重複チェック ====================
function isDuplicateEmail(messageId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('ログ');
    
    if (!logSheet) return false;
    
    const data = logSheet.getDataRange().getValues();
    
    // メッセージIDを含む行があるかチェック（簡易版）
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().includes(messageId)) {
        return true;
      }
    }
    
    return false;
  } catch (error) {
    console.error(`重複チェックエラー: ${error.toString()}`);
    return false;
  }
}

// ==================== Notion送信 ====================
function sendToNotion(settings, subject, body, sender, date, messageId) {
  try {
    // メール本文の長さ制限（Notion APIの制限対応）
    const maxBodyLength = 2000;
    const truncatedBody = body.length > maxBodyLength ? 
      body.substring(0, maxBodyLength) + '\n\n...(本文が長いため省略されました)' : body;
    
    const payload = {
      'parent': {
        'database_id': settings.databaseId
      },
      'properties': {
        [settings.titleProperty]: {
          'title': [
            {
              'text': {
                'content': subject || '(件名なし)'
              }
            }
          ]
        }
      },
      'children': [
        {
          'object': 'block',
          'type': 'paragraph',
          'paragraph': {
            'rich_text': [
              {
                'type': 'text',
                'text': {
                  'content': `📧 送信者: ${sender}\n📅 受信日時: ${date}\n🆔 メッセージID: ${messageId}\n\n📝 本文:\n${truncatedBody}`
                }
              }
            ]
          }
        }
      ]
    };
    
    const response = UrlFetchApp.fetch('https://api.notion.com/v1/pages', {
      'method': 'POST',
      'headers': {
        'Authorization': `Bearer ${settings.notionToken}`,
        'Content-Type': 'application/json',
        'Notion-Version': '2022-06-28'
      },
      'payload': JSON.stringify(payload)
    });
    
    if (response.getResponseCode() === 200) {
      return true;
    } else {
      console.error(`Notion API エラー: ${response.getResponseCode()} - ${response.getContentText()}`);
      return false;
    }
    
  } catch (error) {
    console.error(`Notion送信エラー: ${error.toString()}`);
    return false;
  }
}

// ==================== ログ機能 ====================
function logSuccess(subject, status, sender, processingTime) {
  addLog(subject, status, sender, '', processingTime);
}

function logError(subject, error, sender, processingTime) {
  addLog(subject, 'エラー', sender, error, processingTime);
}

function addLog(subject, status, sender, errorDetails, processingTime) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (!ss) {
      console.error('スプレッドシートにアクセスできません。');
      return;
    }
    
    const logSheet = ss.getSheetByName('ログ');
    
    if (!logSheet) {
      console.error('ログシートが見つかりません。初期設定を実行してください。');
      return;
    }
    
    const timestamp = new Date();
    logSheet.appendRow([timestamp, subject, sender, status, errorDetails, processingTime || '']);
    
    // ログの行数制限（古いログを削除）
    const maxLogRows = 1000;
    const currentRows = logSheet.getLastRow();
    if (currentRows > maxLogRows) {
      logSheet.deleteRows(2, currentRows - maxLogRows); // ヘッダーを除いて古い行を削除
    }
    
  } catch (error) {
    console.error(`ログ記録エラー: ${error.toString()}`);
  }
}

// ==================== 統計更新 ====================
function updateStatistics(successCount, errorCount, processingTime) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const statsSheet = ss.getSheetByName('統計');
    
    if (!statsSheet) {
      console.error('統計シートが見つかりません。');
      return;
    }
    
    // 現在の統計を取得
    const currentStats = statsSheet.getRange('B2:B8').getValues();
    const totalProcessed = parseInt(currentStats[0][0]) + successCount + errorCount;
    const totalSuccess = parseInt(currentStats[1][0]) + successCount;
    const totalError = parseInt(currentStats[2][0]) + errorCount;
    const successRate = totalProcessed > 0 ? ((totalSuccess / totalProcessed) * 100).toFixed(1) + '%' : '0%';
    const latestProcessTime = new Date();
    
    // 平均処理時間の計算（簡易版）
    const avgProcessingTime = processingTime ? `${processingTime.toFixed(2)}秒` : currentStats[5][0];
    
    // 今日の処理件数（簡易版）
    const today = new Date().toDateString();
    const todayProcessed = new Date().toDateString() === new Date().toDateString() ? 
      successCount + errorCount : parseInt(currentStats[6][0]);
    
    // 統計を更新
    statsSheet.getRange('B2:B8').setValues([
      [totalProcessed],
      [totalSuccess],
      [totalError],
      [successRate],
      [latestProcessTime],
      [avgProcessingTime],
      [todayProcessed]
    ]);
    
  } catch (error) {
    console.error(`統計更新エラー: ${error.toString()}`);
  }
}

// ==================== 最終処理日時更新 ====================
function updateLastProcessTime() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('設定');
    
    if (!settingsSheet) {
      console.error('設定シートが見つかりません。初期設定を実行してください。');
      return;
    }
    
    const timestamp = new Date();
    settingsSheet.getRange('B8').setValue(timestamp);
    
  } catch (error) {
    console.error(`最終処理日時更新エラー: ${error.toString()}`);
  }
}

// ==================== ログ統計表示 ====================
function showLogStats() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const statsSheet = ss.getSheetByName('統計');
    
    if (!statsSheet) {
      SpreadsheetApp.getUi().alert('統計シートが見つかりません。初期設定を実行してください。');
      return;
    }
    
    const stats = statsSheet.getRange('B2:B8').getValues();
    
    const message = `
📊 処理統計

総処理件数: ${stats[0][0]}件
成功件数: ${stats[1][0]}件
エラー件数: ${stats[2][0]}件
成功率: ${stats[3][0]}
最新処理日時: ${stats[4][0]}
平均処理時間: ${stats[5][0]}
今日の処理件数: ${stats[6][0]}件
    `;
    
    SpreadsheetApp.getUi().alert('📊 処理統計', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
    // 統計シートをアクティブにする
    ss.setActiveSheet(statsSheet);
    
  } catch (error) {
    console.error(`統計表示エラー: ${error.toString()}`);
    SpreadsheetApp.getUi().alert('エラー', '統計の表示中にエラーが発生しました：\n' + error.toString());
  }
}

// ==================== トリガー管理 ====================
function setupAutoTrigger() {
  try {
    const settings = getSettings();
    
    // 設定値の検証
    const validationErrors = validateSettings(settings);
    if (validationErrors.length > 0) {
      SpreadsheetApp.getUi().alert(
        '⚠️ 設定不備',
        '自動実行を設定する前に以下の設定項目を確認してください：\n' + validationErrors.join('\n'),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // 既存のトリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'processGmailToNotion') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // 新しいトリガーを設定
    const intervalMinutes = settings.intervalMinutes;
    
    if (intervalMinutes >= 60) {
      // 1時間以上の場合は時間ベース
      const hours = Math.floor(intervalMinutes / 60);
      ScriptApp.newTrigger('processGmailToNotion')
        .timeBased()
        .everyHours(hours)
        .create();
    } else {
      // 1時間未満の場合は分ベース
      ScriptApp.newTrigger('processGmailToNotion')
        .timeBased()
        .everyMinutes(intervalMinutes)
        .create();
    }
    
    logSuccess('自動実行設定', `${intervalMinutes}分ごとの自動実行を設定しました`, '', 0);
    
    SpreadsheetApp.getUi().alert(
      '✅ 自動実行設定完了',
      `${intervalMinutes}分ごとの自動実行が設定されました。`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    logError('自動実行設定', error.toString(), '', 0);
    
    SpreadsheetApp.getUi().alert(
      '❌ 設定エラー',
      '自動実行設定中にエラーが発生しました：\n' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function stopAutoTrigger() {
  try {
    // 該当するトリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'processGmailToNotion') {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    });
    
    logSuccess('自動実行停止', `${deletedCount}個のトリガーを停止しました`, '', 0);
    
    SpreadsheetApp.getUi().alert(
      '✅ 自動実行停止完了',
      `自動実行が停止されました。（${deletedCount}個のトリガーを削除）`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    logError('自動実行停止', error.toString(), '', 0);
    
    SpreadsheetApp.getUi().alert(
      '❌ 停止エラー',
      '自動実行停止中にエラーが発生しました：\n' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ==================== 処理後動作 ====================
function handlePostProcessAction(thread, settings) {
  try {
    const action = settings.postProcessAction;
    const originalLabel = GmailApp.getUserLabelByName(settings.gmailLabel);
    
    console.log(`=== 処理後動作開始 ===`);
    console.log(`動作: ${action}`);
    console.log(`スレッド件名: ${thread.getFirstMessageSubject()}`);
    console.log(`現在のラベル: ${thread.getLabels().map(l => l.getName()).join(', ')}`);
    
    switch (action) {
      case 'ラベル削除':
        if (originalLabel && thread.getLabels().includes(originalLabel)) {
          thread.removeLabel(originalLabel);
          console.log(`✅ ラベル「${settings.gmailLabel}」を削除しました`);
        } else {
          console.log(`⚠️ ラベル「${settings.gmailLabel}」が見つからないか、既に削除済みです`);
        }
        break;
        
      case '完了ラベル追加':
        // 完了ラベルを取得または作成
        let completedLabel = GmailApp.getUserLabelByName(settings.completedLabel);
        if (!completedLabel) {
          try {
            completedLabel = GmailApp.createLabel(settings.completedLabel);
            console.log(`✅ 新しいラベル「${settings.completedLabel}」を作成しました`);
          } catch (labelError) {
            console.error(`❌ ラベル作成エラー: ${labelError.toString()}`);
            throw labelError;
          }
        } else {
          console.log(`ℹ️ ラベル「${settings.completedLabel}」は既に存在します`);
        }
        
        // 元のラベルを削除
        if (originalLabel && thread.getLabels().includes(originalLabel)) {
          thread.removeLabel(originalLabel);
          console.log(`✅ 元のラベル「${settings.gmailLabel}」を削除しました`);
        } else {
          console.log(`⚠️ 元のラベル「${settings.gmailLabel}」が見つからないか、既に削除済みです`);
        }
        
        // 完了ラベルを追加（重複チェック）
        const currentLabels = thread.getLabels();
        const hasCompletedLabel = currentLabels.some(label => label.getName() === settings.completedLabel);
        
        if (!hasCompletedLabel) {
          thread.addLabel(completedLabel);
          console.log(`✅ 完了ラベル「${settings.completedLabel}」を追加しました`);
          
          // 追加後の確認
          Utilities.sleep(1000); // 1秒待機
          const updatedLabels = thread.getLabels();
          const finalCheck = updatedLabels.some(label => label.getName() === settings.completedLabel);
          console.log(`🔍 ラベル追加確認: ${finalCheck ? '成功' : '失敗'}`);
          console.log(`🔍 最終ラベル: ${updatedLabels.map(l => l.getName()).join(', ')}`);
        } else {
          console.log(`ℹ️ 完了ラベル「${settings.completedLabel}」は既に付いています`);
        }
        break;
        
      case 'アーカイブ':
        if (originalLabel && thread.getLabels().includes(originalLabel)) {
          thread.removeLabel(originalLabel);
          console.log(`✅ ラベル「${settings.gmailLabel}」を削除しました`);
        }
        thread.moveToArchive();
        console.log('✅ メールをアーカイブしました');
        break;
        
      default:
        console.log(`⚠️ 不明な動作: ${action}、デフォルト処理を実行`);
        if (originalLabel && thread.getLabels().includes(originalLabel)) {
          thread.removeLabel(originalLabel);
          console.log(`✅ デフォルト: ラベル「${settings.gmailLabel}」を削除しました`);
        }
    }
    
    console.log(`=== 処理後動作完了 ===`);
    
  } catch (error) {
    console.error(`❌ 処理後動作エラー: ${error.toString()}`);
    console.error(`エラースタック: ${error.stack}`);
    
    // エラー発生時のフォールバック処理
    try {
      const fallbackLabel = GmailApp.getUserLabelByName(settings.gmailLabel);
      if (fallbackLabel && thread.getLabels().includes(fallbackLabel)) {
        thread.removeLabel(fallbackLabel);
        console.log('✅ エラー時フォールバック: 元のラベルを削除しました');
      }
    } catch (fallbackError) {
      console.error(`❌ フォールバック処理もエラー: ${fallbackError.toString()}`);
    }
    
    // エラーをログに記録
    logError('処理後動作', error.toString(), thread.getFirstMessageSubject(), 0);
  }
}

// ==================== デバッグ用関数 ====================
function debugLabelStatus() {
  try {
    const settings = getSettings();
    
    // 現在の設定を確認
    console.log('=== 設定確認 ===');
    console.log(`処理後の動作: ${settings.postProcessAction}`);
    console.log(`完了ラベル名: ${settings.completedLabel}`);
    console.log(`Gmailラベル名: ${settings.gmailLabel}`);
    
    // ラベルの存在確認
    const originalLabel = GmailApp.getUserLabelByName(settings.gmailLabel);
    const completedLabel = GmailApp.getUserLabelByName(settings.completedLabel);
    
    console.log('=== ラベル存在確認 ===');
    console.log(`元ラベル「${settings.gmailLabel}」存在: ${originalLabel ? 'あり' : 'なし'}`);
    console.log(`完了ラベル「${settings.completedLabel}」存在: ${completedLabel ? 'あり' : 'なし'}`);
    
    // 対象メールスレッドの確認
    const threads = GmailApp.search(`label:${settings.gmailLabel}`);
    console.log(`=== 対象メール確認 ===`);
    console.log(`対象スレッド数: ${threads.length}`);
    
    if (threads.length > 0) {
      const firstThread = threads[0];
      const labels = firstThread.getLabels();
      console.log(`最初のスレッドのラベル: ${labels.map(l => l.getName()).join(', ')}`);
    }
    
    // 完了ラベル付きメールの確認
    if (completedLabel) {
      const completedThreads = GmailApp.search(`label:${settings.completedLabel}`);
      console.log(`完了ラベル付きスレッド数: ${completedThreads.length}`);
    }
    
    // UI表示
    const message = `
🔍 デバッグ情報

設定:
・処理後の動作: ${settings.postProcessAction}
・完了ラベル名: ${settings.completedLabel}
・Gmailラベル名: ${settings.gmailLabel}

ラベル存在:
・元ラベル: ${originalLabel ? '✅' : '❌'}
・完了ラベル: ${completedLabel ? '✅' : '❌'}

メール状況:
・処理対象: ${threads.length}件
・処理済み: ${completedLabel ? GmailApp.search(`label:${settings.completedLabel}`).length : 0}件

詳細は実行ログを確認してください。
    `;
    
    SpreadsheetApp.getUi().alert('🔍 デバッグ情報', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error(`デバッグエラー: ${error.toString()}`);
    SpreadsheetApp.getUi().alert('❌ デバッグエラー', error.toString());
  }
}

// ==================== ユーティリティ関数 ====================

/**
 * 現在のトリガー状態を確認
 */
function checkTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  const activeTriggers = triggers.filter(trigger => trigger.getHandlerFunction() === 'processGmailToNotion');
  
  console.log(`アクティブなトリガー数: ${activeTriggers.length}`);
  activeTriggers.forEach(trigger => {
    console.log(`トリガーID: ${trigger.getUniqueId()}, 種類: ${trigger.getEventType()}`);
  });
  
  return activeTriggers.length;
}

/**
 * エラー通知機能（オプション）
 */
function sendErrorNotification(errorMessage) {
  try {
    // 管理者メールアドレスが設定されている場合のみ送信
    const adminEmail = Session.getActiveUser().getEmail();
    
    if (adminEmail) {
      MailApp.sendEmail({
        to: adminEmail,
        subject: 'Gmail → Notion 自動連携エラー通知',
        body: `エラーが発生しました：\n\n${errorMessage}\n\n詳細はログシートを確認してください。`
      });
    }
  } catch (error) {
    console.error(`エラー通知送信失敗: ${error.toString()}`);
  }
}
