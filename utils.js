/**
 * Tournament Pro - ユーティリティ関数
 */

/**
 * 日付を YYYY-MM-DD 形式にフォーマット
 * @param {Date} date - 日付オブジェクト
 * @return {string} フォーマットされた日付
 */
function formatDate(date) {
  if (!date) return '';
  
  var d = new Date(date);
  var year = d.getFullYear();
  var month = ('0' + (d.getMonth() + 1)).slice(-2);
  var day = ('0' + d.getDate()).slice(-2);
  
  return year + '-' + month + '-' + day;
}

/**
 * 時刻を HH:MM 形式にフォーマット
 * @param {Date|string} time - 時刻
 * @return {string} フォーマットされた時刻
 */
function formatTime(time) {
  if (!time) return '';
  
  if (typeof time === 'string') return time;
  
  var d = new Date(time);
  var hours = ('0' + d.getHours()).slice(-2);
  var minutes = ('0' + d.getMinutes()).slice(-2);
  
  return hours + ':' + minutes;
}

/**
 * 配列をシャッフル
 * @param {Array} array - 配列
 * @return {Array} シャッフルされた配列
 */
function shuffleArray(array) {
  var shuffled = array.slice();
  for (var i = shuffled.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = shuffled[i];
    shuffled[i] = shuffled[j];
    shuffled[j] = temp;
  }
  return shuffled;
}

/**
 * チーム名をIDから取得
 * @param {number|string} teamId - チームID
 * @return {string} チーム名
 */
function getTeamName(teamId) {
  try {
    if (!teamId) return '';
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Teams');
    
    if (!sheet) return '';
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == teamId) {
        return data[i][1];
      }
    }
    
    return '';
  } catch (error) {
    Logger.log('getTeamName Error: ' + error.toString());
    return '';
  }
}

/**
 * エラーログを記録
 * @param {string} functionName - 関数名
 * @param {Error} error - エラーオブジェクト
 */
function logError(functionName, error) {
  var message = '[' + new Date().toISOString() + '] ' + functionName + ': ' + error.toString();
  Logger.log(message);
  
  // 将来的にはエラーログシートに記録することも可能
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var logSheet = ss.getSheetByName('ErrorLog');
  // if (logSheet) {
  //   logSheet.appendRow([new Date(), functionName, error.toString()]);
  // }
}

/**
 * 設定値を安全に取得
 * @param {Object} settings - 設定オブジェクト
 * @param {string} key - キー
 * @param {*} defaultValue - デフォルト値
 * @return {*} 設定値
 */
function getSetting(settings, key, defaultValue) {
  if (settings && settings[key] !== undefined && settings[key] !== null && settings[key] !== '') {
    return settings[key];
  }
  return defaultValue;
}
