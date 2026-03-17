/**
 * Tournament Pro - 審判割り当て最適化
 */

/**
 * 審判を自動割り当て
 * @return {Object} 結果
 */
function autoAssignReferees() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var matchesSheet = ss.getSheetByName('Matches');
    var teamsSheet = ss.getSheetByName('Teams');
    
    if (!matchesSheet || !teamsSheet) {
      return { success: false, error: '必要なシートが見つかりません' };
    }
    
    // 全チームを取得
    var teamsData = teamsSheet.getDataRange().getValues();
    var allTeams = [];
    for (var i = 1; i < teamsData.length; i++) {
      if (teamsData[i][0]) {
        allTeams.push(teamsData[i][0]);
      }
    }
    
    // 審判回数をカウント
    var refereeCount = {};
    allTeams.forEach(function(teamId) {
      refereeCount[teamId] = 0;
    });
    
    // 試合データを取得
    var matchesData = matchesSheet.getDataRange().getValues();
    
    // 各試合に審判を割り当て
    for (var i = 1; i < matchesData.length; i++) {
      var teamA = matchesData[i][2];
      var teamB = matchesData[i][3];
      var currentTime = matchesData[i][5];
      
      // 候補チームを取得（対戦チーム以外）
      var candidates = allTeams.filter(function(teamId) {
        // 対戦チーム以外
        if (teamId === teamA || teamId === teamB) return false;
        
        // 同じ時間帯に試合があるチームは除外
        var hasMatchAtSameTime = false;
        for (var j = 1; j < matchesData.length; j++) {
          if (matchesData[j][5] === currentTime) {
            if (matchesData[j][2] === teamId || matchesData[j][3] === teamId) {
              hasMatchAtSameTime = true;
              break;
            }
          }
        }
        return !hasMatchAtSameTime;
      });
      
      // 審判回数が最も少ないチームを選択
      if (candidates.length > 0) {
        candidates.sort(function(a, b) {
          return refereeCount[a] - refereeCount[b];
        });
        
        var assignedReferee = candidates[0];
        matchesSheet.getRange(i + 1, 7).setValue(assignedReferee);
        refereeCount[assignedReferee]++;
      }
    }
    
    return {
      success: true,
      message: '審判を自動割り当てしました'
    };
  } catch (error) {
    Logger.log('autoAssignReferees Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 審判割り当てのバランスを確認
 * @return {Object} 各チームの審判回数
 */
function checkRefereeBalance() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var matchesSheet = ss.getSheetByName('Matches');
    var teamsSheet = ss.getSheetByName('Teams');
    
    if (!matchesSheet || !teamsSheet) {
      return { success: false, error: '必要なシートが見つかりません' };
    }
    
    var teamsData = teamsSheet.getDataRange().getValues();
    var matchesData = matchesSheet.getDataRange().getValues();
    
    var refereeCount = {};
    
    // 全チームの審判回数を初期化
    for (var i = 1; i < teamsData.length; i++) {
      if (teamsData[i][0]) {
        refereeCount[teamsData[i][0]] = 0;
      }
    }
    
    // 審判回数をカウント
    for (var i = 1; i < matchesData.length; i++) {
      var referee = matchesData[i][6];
      if (referee && refereeCount[referee] !== undefined) {
        refereeCount[referee]++;
      }
    }
    
    return {
      success: true,
      data: refereeCount
    };
  } catch (error) {
    Logger.log('checkRefereeBalance Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}
