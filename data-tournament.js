/**
 * Tournament Pro - トーナメントデータ管理
 */

/**
 * トーナメントデータを取得
 * @return {Object} トーナメントデータ
 */
function getTournamentData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Tournament');
    
    if (!sheet) {
      return { success: false, error: 'Tournamentシートが見つかりません' };
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, data: [] };
    }
    
    var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    var tournament = [];
    
    data.forEach(function(row) {
      if (row[0]) { // ラウンドが存在する行のみ
        tournament.push({
          round: row[0],
          matchNo: row[1],
          teamA: row[2],
          teamB: row[3],
          winner: row[4],
          scoreA_S1: row[5],
          scoreB_S1: row[6],
          scoreA_S2: row[7],
          scoreB_S2: row[8],
          scoreA_S3: row[9],
          scoreB_S3: row[10],
          status: row[11]
        });
      }
    });
    
    return {
      success: true,
      data: tournament
    };
  } catch (error) {
    Logger.log('getTournamentData Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * トーナメント試合を更新
 * @param {Object} match - 試合情報
 * @return {Object} 結果
 */
function updateTournamentMatch(match) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Tournament');
    
    if (!sheet) {
      return { success: false, error: 'Tournamentシートが見つかりません' };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === match.round && data[i][1] === match.matchNo) {
        if (match.teamA !== undefined) sheet.getRange(i + 1, 3).setValue(match.teamA);
        if (match.teamB !== undefined) sheet.getRange(i + 1, 4).setValue(match.teamB);
        if (match.winner !== undefined) sheet.getRange(i + 1, 5).setValue(match.winner);
        if (match.scoreA_S1 !== undefined) sheet.getRange(i + 1, 6).setValue(match.scoreA_S1);
        if (match.scoreB_S1 !== undefined) sheet.getRange(i + 1, 7).setValue(match.scoreB_S1);
        if (match.scoreA_S2 !== undefined) sheet.getRange(i + 1, 8).setValue(match.scoreA_S2);
        if (match.scoreB_S2 !== undefined) sheet.getRange(i + 1, 9).setValue(match.scoreB_S2);
        if (match.scoreA_S3 !== undefined) sheet.getRange(i + 1, 10).setValue(match.scoreA_S3);
        if (match.scoreB_S3 !== undefined) sheet.getRange(i + 1, 11).setValue(match.scoreB_S3);
        if (match.status !== undefined) sheet.getRange(i + 1, 12).setValue(match.status);
        
        return {
          success: true,
          message: 'トーナメント試合を更新しました'
        };
      }
    }
    
    return {
      success: false,
      error: '試合が見つかりません'
    };
  } catch (error) {
    Logger.log('updateTournamentMatch Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 順位表からトーナメント出場チームを自動設定
 * @return {Object} 結果
 */
function generateTournamentBracket() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var standingsSheet = ss.getSheetByName('Standings');
    var tournamentSheet = ss.getSheetByName('Tournament');
    var settingsResult = getTournamentSettings();
    
    if (!standingsSheet || !tournamentSheet || !settingsResult.success) {
      return { success: false, error: '必要なシートまたは設定が見つかりません' };
    }
    
    var settings = settingsResult.data;
    var teamsAdvancing = parseInt(settings.teamsAdvancing) || 2;
    
    // 各グループの上位チームを取得
    var standingsData = standingsSheet.getDataRange().getValues();
    var qualifiedTeams = [];
    
    var groups = ['A', 'B', 'C', 'D'];
    groups.forEach(function(group) {
      var groupTeams = [];
      for (var i = 1; i < standingsData.length; i++) {
        if (standingsData[i][0] === group) {
          groupTeams.push({
            group: standingsData[i][0],
            teamId: standingsData[i][1],
            rank: standingsData[i][8]
          });
        }
      }
      
      // 順位でソート
      groupTeams.sort(function(a, b) { return a.rank - b.rank; });
      
      // 上位N チームを取得
      for (var j = 0; j < Math.min(teamsAdvancing, groupTeams.length); j++) {
        qualifiedTeams.push(groupTeams[j]);
      }
    });
    
    // トーナメントシートをクリア
    tournamentSheet.clearContents();
    tournamentSheet.appendRow(['ラウンド', 'マッチNo', 'チームA_ID', 'チームB_ID', '勝者ID', 'スコアA_S1', 'スコアB_S1', 'スコアA_S2', 'スコアB_S2', 'スコアA_S3', 'スコアB_S3', '状態']);
    
    // 8チームの場合のトーナメント組み合わせ
    if (qualifiedTeams.length === 8) {
      // 準々決勝
      tournamentSheet.appendRow(['準々決勝', 1, qualifiedTeams[0].teamId, qualifiedTeams[7].teamId, '', '', '', '', '', '', '', '未開始']);
      tournamentSheet.appendRow(['準々決勝', 2, qualifiedTeams[3].teamId, qualifiedTeams[4].teamId, '', '', '', '', '', '', '', '未開始']);
      tournamentSheet.appendRow(['準々決勝', 3, qualifiedTeams[1].teamId, qualifiedTeams[6].teamId, '', '', '', '', '', '', '', '未開始']);
      tournamentSheet.appendRow(['準々決勝', 4, qualifiedTeams[2].teamId, qualifiedTeams[5].teamId, '', '', '', '', '', '', '', '未開始']);
      
      // 準決勝
      tournamentSheet.appendRow(['準決勝', 1, '', '', '', '', '', '', '', '', '', '未開始']);
      tournamentSheet.appendRow(['準決勝', 2, '', '', '', '', '', '', '', '', '', '未開始']);
      
      // 3位決定戦
      tournamentSheet.appendRow(['3位決定戦', 1, '', '', '', '', '', '', '', '', '', '未開始']);
      
      // 決勝
      tournamentSheet.appendRow(['決勝', 1, '', '', '', '', '', '', '', '', '', '未開始']);
    }
    
    return {
      success: true,
      message: 'トーナメント表を生成しました',
      qualifiedTeams: qualifiedTeams.length
    };
  } catch (error) {
    Logger.log('generateTournamentBracket Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}
