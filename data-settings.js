/**
 * Tournament Pro - 大会設定データ管理
 */

/**
 * 大会設定を取得
 * @return {Object} 設定データ
 */
function getTournamentSettings() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Settings');
    
    if (!sheet) {
      return { success: false, error: 'Settingsシートが見つかりません' };
    }
    
    var data = sheet.getDataRange().getValues();
    var settings = {};
    
    // 1行目はヘッダーなので2行目から
    for (var i = 1; i < data.length; i++) {
      var key = data[i][0];
      var value = data[i][1];
      settings[key] = value;
    }
    
    return {
      success: true,
      data: settings
    };
  } catch (error) {
    Logger.log('getTournamentSettings Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 大会設定を更新
 * @param {Object} settings - 設定データ
 * @return {Object} 結果
 */
function updateTournamentSettings(settings) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Settings');
    
    if (!sheet) {
      return { success: false, error: 'Settingsシートが見つかりません' };
    }
    
    var data = sheet.getDataRange().getValues();
    
    // 各設定値を更新
    for (var key in settings) {
      var found = false;
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
          sheet.getRange(i + 1, 2).setValue(settings[key]);
          found = true;
          break;
        }
      }
      
      // 新しいキーの場合は追加
      if (!found) {
        sheet.appendRow([key, settings[key]]);
      }
    }
    
    return {
      success: true,
      message: '設定を更新しました'
    };
  } catch (error) {
    Logger.log('updateTournamentSettings Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 大会要項データを取得
 * @return {Object} 要項データ
 */
function getTournamentDocument() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('TournamentDocument');
    
    if (!sheet) {
      return { success: false, error: 'TournamentDocumentシートが見つかりません' };
    }
    
    var data = sheet.getDataRange().getValues();
    var document = {};
    
    for (var i = 1; i < data.length; i++) {
      var key = data[i][0];
      var value = data[i][1];
      
      // customFieldsはJSON形式で保存されているのでパース
      if (key === 'customFields' && value) {
        try {
          document[key] = JSON.parse(value);
        } catch (e) {
          document[key] = [];
        }
      } else {
        document[key] = value;
      }
    }
    
    return {
      success: true,
      data: document
    };
  } catch (error) {
    Logger.log('getTournamentDocument Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 大会要項データを更新
 * @param {Object} document - 要項データ
 * @return {Object} 結果
 */
function updateTournamentDocument(document) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('TournamentDocument');
    
    if (!sheet) {
      return { success: false, error: 'TournamentDocumentシートが見つかりません' };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var key in document) {
      var value = document[key];
      
      // customFieldsはJSON形式で保存
      if (key === 'customFields') {
        value = JSON.stringify(value);
      }
      
      var found = false;
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
          sheet.getRange(i + 1, 2).setValue(value);
          found = true;
          break;
        }
      }
      
      if (!found) {
        sheet.appendRow([key, value]);
      }
    }
    
    return {
      success: true,
      message: '要項データを更新しました'
    };
  } catch (error) {
    Logger.log('updateTournamentDocument Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}
