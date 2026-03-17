/**
 * Tournament Pro - チームデータ管理
 */

/**
 * チームデータを取得
 * @return {Object} チームデータ
 */
function getTeamsData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Teams');
    
    if (!sheet) {
      return { success: false, error: 'Teamsシートが見つかりません' };
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, data: [] };
    }
    
    var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    var teams = [];
    
    data.forEach(function(row) {
      if (row[0]) { // IDが存在する行のみ
        teams.push({
          id: row[0],
          name: row[1],
          group: row[2],
          rank: row[3],
          contact: row[4],
          paymentStatus: row[5]
        });
      }
    });
    
    return {
      success: true,
      data: teams
    };
  } catch (error) {
    Logger.log('getTeamsData Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * チームを追加
 * @param {Object} team - チーム情報
 * @return {Object} 結果
 */
function addTeam(team) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Teams');
    
    if (!sheet) {
      return { success: false, error: 'Teamsシートが見つかりません' };
    }
    
    // 新しいIDを生成（既存の最大ID + 1）
    var lastRow = sheet.getLastRow();
    var newId = 1;
    
    if (lastRow > 1) {
      var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      var maxId = Math.max.apply(null, ids.map(function(row) { return row[0] || 0; }));
      newId = maxId + 1;
    }
    
    sheet.appendRow([
      newId,
      team.name || '',
      team.group || '',
      team.rank || '',
      team.contact || '',
      team.paymentStatus || false
    ]);
    
    return {
      success: true,
      message: 'チームを追加しました',
      teamId: newId
    };
  } catch (error) {
    Logger.log('addTeam Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * チームを更新
 * @param {Object} team - チーム情報
 * @return {Object} 結果
 */
function updateTeam(team) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Teams');
    
    if (!sheet) {
      return { success: false, error: 'Teamsシートが見つかりません' };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == team.id) {
        sheet.getRange(i + 1, 2).setValue(team.name || '');
        sheet.getRange(i + 1, 3).setValue(team.group || '');
        sheet.getRange(i + 1, 4).setValue(team.rank || '');
        sheet.getRange(i + 1, 5).setValue(team.contact || '');
        sheet.getRange(i + 1, 6).setValue(team.paymentStatus || false);
        
        return {
          success: true,
          message: 'チームを更新しました'
        };
      }
    }
    
    return {
      success: false,
      error: 'チームが見つかりません'
    };
  } catch (error) {
    Logger.log('updateTeam Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * チームを削除
 * @param {number} teamId - チームID
 * @return {Object} 結果
 */
function deleteTeam(teamId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Teams');
    
    if (!sheet) {
      return { success: false, error: 'Teamsシートが見つかりません' };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == teamId) {
        sheet.deleteRow(i + 1);
        return {
          success: true,
          message: 'チームを削除しました'
        };
      }
    }
    
    return {
      success: false,
      error: 'チームが見つかりません'
    };
  } catch (error) {
    Logger.log('deleteTeam Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 複数チームを一括追加
 * @param {Array} teams - チーム情報の配列
 * @return {Object} 結果
 */
function bulkAddTeams(teams) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Teams');
    
    if (!sheet) {
      return { success: false, error: 'Teamsシートが見つかりません' };
    }
    
    var lastRow = sheet.getLastRow();
    var startId = 1;
    
    if (lastRow > 1) {
      var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      var maxId = Math.max.apply(null, ids.map(function(row) { return row[0] || 0; }));
      startId = maxId + 1;
    }
    
    var rows = [];
    teams.forEach(function(team, index) {
      rows.push([
        startId + index,
        team.name || '',
        team.group || '',
        team.rank || '',
        team.contact || '',
        team.paymentStatus || false
      ]);
    });
    
    if (rows.length > 0) {
      sheet.getRange(lastRow + 1, 1, rows.length, 6).setValues(rows);
    }
    
    return {
      success: true,
      message: teams.length + '件のチームを追加しました'
    };
  } catch (error) {
    Logger.log('bulkAddTeams Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}
