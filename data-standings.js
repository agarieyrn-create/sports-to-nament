/**
 * Tournament Pro - 順位表データ管理
 */

/**
 * 順位表データを取得
 * @return {Object} 順位表データ
 */
function getStandingsData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Standings');
    
    if (!sheet) {
      return { success: false, error: 'Standingsシートが見つかりません' };
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, data: [] };
    }
    
    var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    var standings = [];
    
    data.forEach(function(row) {
      if (row[1]) { // チームIDが存在する行のみ
        standings.push({
          group: row[0],
          teamId: row[1],
          points: row[2],
          wins: row[3],
          losses: row[4],
          scored: row[5],
          conceded: row[6],
          difference: row[7],
          rank: row[8]
        });
      }
    });
    
    return {
      success: true,
      data: standings
    };
  } catch (error) {
    Logger.log('getStandingsData Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 試合結果から順位表を更新
 * @return {Object} 結果
 */
function updateStandingsFromMatches() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var matchesSheet = ss.getSheetByName('Matches');
    var standingsSheet = ss.getSheetByName('Standings');
    var teamsSheet = ss.getSheetByName('Teams');
    
    if (!matchesSheet || !standingsSheet || !teamsSheet) {
      return { success: false, error: '必要なシートが見つかりません' };
    }
    
    // チームデータを取得
    var teamsData = teamsSheet.getDataRange().getValues();
    var teams = {};
    for (var i = 1; i < teamsData.length; i++) {
      if (teamsData[i][0]) {
        teams[teamsData[i][0]] = {
          id: teamsData[i][0],
          name: teamsData[i][1],
          group: teamsData[i][2]
        };
      }
    }
    
    // 順位表を初期化
    var standings = {};
    for (var teamId in teams) {
      var team = teams[teamId];
      if (team.group) {
        if (!standings[team.group]) {
          standings[team.group] = {};
        }
        standings[team.group][teamId] = {
          teamId: teamId,
          points: 0,
          wins: 0,
          losses: 0,
          scored: 0,
          conceded: 0
        };
      }
    }
    
    // 完了した予選リーグ試合を集計
    var matchesData = matchesSheet.getDataRange().getValues();
    for (var i = 1; i < matchesData.length; i++) {
      var match = matchesData[i];
      var stage = match[1];
      var teamA = match[2];
      var teamB = match[3];
      var status = match[13];
      
      // 予選リーグかつ完了した試合のみ
      if (stage !== '予選リーグ' || status !== '完了') continue;
      
      var scoreA_S1 = match[7] || 0;
      var scoreB_S1 = match[8] || 0;
      var scoreA_S2 = match[9] || 0;
      var scoreB_S2 = match[10] || 0;
      var scoreA_S3 = match[11] || 0;
      var scoreB_S3 = match[12] || 0;
      
      // セット数を計算
      var setsA = 0;
      var setsB = 0;
      if (scoreA_S1 > scoreB_S1) setsA++; else setsB++;
      if (scoreA_S2 > scoreB_S2) setsA++; else setsB++;
      if (scoreA_S3 && scoreB_S3) {
        if (scoreA_S3 > scoreB_S3) setsA++; else setsB++;
      }
      
      // 得点・失点を計算
      var totalScoreA = scoreA_S1 + scoreA_S2 + scoreA_S3;
      var totalScoreB = scoreB_S1 + scoreB_S2 + scoreB_S3;
      
      // チームAとチームBのグループを取得
      var groupA = teams[teamA] ? teams[teamA].group : null;
      var groupB = teams[teamB] ? teams[teamB].group : null;
      
      if (groupA && standings[groupA] && standings[groupA][teamA]) {
        standings[groupA][teamA].scored += totalScoreA;
        standings[groupA][teamA].conceded += totalScoreB;
        
        if (setsA > setsB) {
          standings[groupA][teamA].wins++;
          standings[groupA][teamA].points += 2;
        } else {
          standings[groupA][teamA].losses++;
        }
      }
      
      if (groupB && standings[groupB] && standings[groupB][teamB]) {
        standings[groupB][teamB].scored += totalScoreB;
        standings[groupB][teamB].conceded += totalScoreA;
        
        if (setsB > setsA) {
          standings[groupB][teamB].wins++;
          standings[groupB][teamB].points += 2;
        } else {
          standings[groupB][teamB].losses++;
        }
      }
    }
    
    // 順位表シートをクリアして再作成
    standingsSheet.clearContents();
    standingsSheet.appendRow(['グループ', 'チームID', '勝点', '勝', '敗', '得点', '失点', '得失点差', '順位']);
    
    // グループごとに順位付け
    for (var group in standings) {
      var groupTeams = [];
      for (var teamId in standings[group]) {
        var teamStanding = standings[group][teamId];
        teamStanding.difference = teamStanding.scored - teamStanding.conceded;
        groupTeams.push(teamStanding);
      }
      
      // ソート: 勝点 > 得失点差 > 総得点
      groupTeams.sort(function(a, b) {
        if (b.points !== a.points) return b.points - a.points;
        if (b.difference !== a.difference) return b.difference - a.difference;
        return b.scored - a.scored;
      });
      
      // 順位を付与してシートに追加
      groupTeams.forEach(function(team, index) {
        standingsSheet.appendRow([
          group,
          team.teamId,
          team.points,
          team.wins,
          team.losses,
          team.scored,
          team.conceded,
          team.difference,
          index + 1
        ]);
      });
    }
    
    return {
      success: true,
      message: '順位表を更新しました'
    };
  } catch (error) {
    Logger.log('updateStandingsFromMatches Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}
