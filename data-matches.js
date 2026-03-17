/**
 * Tournament Pro - 試合データ管理
 */

/**
 * 試合データを取得
 * @return {Object} 試合データ
 */
function getMatchesData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Matches');
    
    if (!sheet) {
      return { success: false, error: 'Matchesシートが見つかりません' };
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, data: [] };
    }
    
    var data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    var matches = [];
    
    data.forEach(function(row) {
      if (row[0]) { // 試合IDが存在する行のみ
        matches.push({
          matchId: row[0],
          stage: row[1],
          teamA: row[2],
          teamB: row[3],
          court: row[4],
          startTime: row[5],
          referee: row[6],
          scoreA_S1: row[7],
          scoreB_S1: row[8],
          scoreA_S2: row[9],
          scoreB_S2: row[10],
          scoreA_S3: row[11],
          scoreB_S3: row[12],
          status: row[13]
        });
      }
    });
    
    return {
      success: true,
      data: matches
    };
  } catch (error) {
    Logger.log('getMatchesData Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 試合を追加
 * @param {Object} match - 試合情報
 * @return {Object} 結果
 */
function addMatch(match) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Matches');
    
    if (!sheet) {
      return { success: false, error: 'Matchesシートが見つかりません' };
    }
    
    sheet.appendRow([
      match.matchId || '',
      match.stage || '',
      match.teamA || '',
      match.teamB || '',
      match.court || '',
      match.startTime || '',
      match.referee || '',
      match.scoreA_S1 || '',
      match.scoreB_S1 || '',
      match.scoreA_S2 || '',
      match.scoreB_S2 || '',
      match.scoreA_S3 || '',
      match.scoreB_S3 || '',
      match.status || '未開始'
    ]);
    
    return {
      success: true,
      message: '試合を追加しました'
    };
  } catch (error) {
    Logger.log('addMatch Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 試合スコアを更新
 * @param {string} matchId - 試合ID
 * @param {Object} scores - スコア情報
 * @return {Object} 結果
 */
function updateMatchScore(matchId, scores) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Matches');
    
    if (!sheet) {
      return { success: false, error: 'Matchesシートが見つかりません' };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == matchId) {
        if (scores.scoreA_S1 !== undefined) sheet.getRange(i + 1, 8).setValue(scores.scoreA_S1);
        if (scores.scoreB_S1 !== undefined) sheet.getRange(i + 1, 9).setValue(scores.scoreB_S1);
        if (scores.scoreA_S2 !== undefined) sheet.getRange(i + 1, 10).setValue(scores.scoreA_S2);
        if (scores.scoreB_S2 !== undefined) sheet.getRange(i + 1, 11).setValue(scores.scoreB_S2);
        if (scores.scoreA_S3 !== undefined) sheet.getRange(i + 1, 12).setValue(scores.scoreA_S3);
        if (scores.scoreB_S3 !== undefined) sheet.getRange(i + 1, 13).setValue(scores.scoreB_S3);
        if (scores.status !== undefined) sheet.getRange(i + 1, 14).setValue(scores.status);
        
        // スコア更新後、順位表も更新
        updateStandingsFromMatches();
        
        return {
          success: true,
          message: 'スコアを更新しました'
        };
      }
    }
    
    return {
      success: false,
      error: '試合が見つかりません'
    };
  } catch (error) {
    Logger.log('updateMatchScore Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 試合状態を更新
 * @param {string} matchId - 試合ID
 * @param {string} status - 状態（未開始/進行中/完了）
 * @return {Object} 結果
 */
function updateMatchStatus(matchId, status) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Matches');
    
    if (!sheet) {
      return { success: false, error: 'Matchesシートが見つかりません' };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == matchId) {
        sheet.getRange(i + 1, 14).setValue(status);
        return {
          success: true,
          message: '試合状態を更新しました'
        };
      }
    }
    
    return {
      success: false,
      error: '試合が見つかりません'
    };
  } catch (error) {
    Logger.log('updateMatchStatus Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 複数試合を一括追加
 * @param {Array} matches - 試合情報の配列
 * @return {Object} 結果
 */
function bulkAddMatches(matches) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Matches');
    
    if (!sheet) {
      return { success: false, error: 'Matchesシートが見つかりません' };
    }
    
    var rows = [];
    matches.forEach(function(match) {
      rows.push([
        match.matchId || '',
        match.stage || '',
        match.teamA || '',
        match.teamB || '',
        match.court || '',
        match.startTime || '',
        match.referee || '',
        match.scoreA_S1 || '',
        match.scoreB_S1 || '',
        match.scoreA_S2 || '',
        match.scoreB_S2 || '',
        match.scoreA_S3 || '',
        match.scoreB_S3 || '',
        match.status || '未開始'
      ]);
    });
    
    if (rows.length > 0) {
      var lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, rows.length, 14).setValues(rows);
    }
    
    return {
      success: true,
      message: matches.length + '件の試合を追加しました'
    };
  } catch (error) {
    Logger.log('bulkAddMatches Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}
