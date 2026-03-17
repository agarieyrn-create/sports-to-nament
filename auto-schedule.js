/**
 * Tournament Pro - 試合スケジュール自動生成
 */

/**
 * 予選リーグのスケジュールを自動生成
 * @param {Object} options - オプション設定
 * @return {Object} 結果
 */
function generateLeagueSchedule(options) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var teamsSheet = ss.getSheetByName('Teams');
    var matchesSheet = ss.getSheetByName('Matches');
    var settingsResult = getTournamentSettings();
    
    if (!teamsSheet || !matchesSheet || !settingsResult.success) {
      return { success: false, error: '必要なシートまたは設定が見つかりません' };
    }
    
    var settings = settingsResult.data;
    var courtCount = parseInt(settings.courtCount) || 4;
    var startTime = options.startTime || '13:00';
    var matchDuration = options.matchDuration || 20; // 分
    
    // チームをグループ別に取得
    var teamsData = teamsSheet.getDataRange().getValues();
    var groups = {};
    
    for (var i = 1; i < teamsData.length; i++) {
      if (teamsData[i][0] && teamsData[i][2]) {
        var teamId = teamsData[i][0];
        var group = teamsData[i][2];
        
        if (!groups[group]) {
          groups[group] = [];
        }
        groups[group].push(teamId);
      }
    }
    
    // 各グループで総当たり戦の組み合わせを生成
    var allMatches = [];
    var matchId = 1;
    var courts = ['A', 'B', 'C', 'D'];
    var currentCourt = 0;
    var currentTime = startTime;
    
    for (var group in groups) {
      var teams = groups[group];
      
      // 総当たり戦の組み合わせ
      for (var i = 0; i < teams.length; i++) {
        for (var j = i + 1; j < teams.length; j++) {
          var court = courts[currentCourt % courtCount];
          
          allMatches.push({
            matchId: group + '-' + matchId,
            stage: '予選リーグ',
            teamA: teams[i],
            teamB: teams[j],
            court: court,
            startTime: currentTime,
            referee: '',
            status: '未開始'
          });
          
          matchId++;
          currentCourt++;
          
          // コートを一巡したら時間を進める
          if (currentCourt % courtCount === 0) {
            currentTime = addMinutesToTime(currentTime, matchDuration);
          }
        }
      }
    }
    
    // 既存の試合データをクリア（予選のみ）
    var existingData = matchesSheet.getDataRange().getValues();
    for (var i = existingData.length - 1; i > 0; i--) {
      if (existingData[i][1] === '予選リーグ') {
        matchesSheet.deleteRow(i + 1);
      }
    }
    
    // 新しい試合を追加
    var result = bulkAddMatches(allMatches);
    
    // 審判を自動割り当て
    if (result.success) {
      autoAssignReferees();
    }
    
    return {
      success: true,
      message: allMatches.length + '試合のスケジュールを生成しました',
      matches: allMatches.length
    };
  } catch (error) {
    Logger.log('generateLeagueSchedule Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 時間に分を加算
 * @param {string} time - 時刻（HH:MM形式）
 * @param {number} minutes - 加算する分
 * @return {string} 新しい時刻
 */
function addMinutesToTime(time, minutes) {
  var parts = time.split(':');
  var hours = parseInt(parts[0]);
  var mins = parseInt(parts[1]);
  
  mins += minutes;
  hours += Math.floor(mins / 60);
  mins = mins % 60;
  
  return (hours < 10 ? '0' : '') + hours + ':' + (mins < 10 ? '0' : '') + mins;
}
