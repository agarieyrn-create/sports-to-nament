const SPREADSHEET_ID = '1XlrauDVN1xLowyapSW5haIio1LMB9t0Kn2WQuu5jxa4';

function doGet(e) {
  const page = e.parameter.page || 'index';
  
  if (page === 'setup') {
    return HtmlService.createHtmlOutputFromFile('setup')
      .setTitle('初期設定 - Tournament Pro')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('admin')
      .setTitle('管理画面 - Tournament Pro')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (page === 'public') {
    return HtmlService.createHtmlOutputFromFile('public')
      .setTitle('観客画面 - Tournament Pro')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Tournament Pro - トップページ')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// 全データを取得
function getAllData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 設定を読み込み
    const settings = getSettings();
    
    // チームを読み込み
    const teams = getTeams();
    
    // 試合を読み込み
    const matches = getMatches();
    
    Logger.log('getAllData 実行完了');
    Logger.log('設定: ' + JSON.stringify(settings));
    Logger.log('チーム数: ' + teams.length);
    Logger.log('試合数: ' + matches.length);
    
    return {
      settings: settings,
      teams: teams,
      matches: matches
    };
  } catch (error) {
    Logger.log('getAllData エラー: ' + error.toString());
    throw new Error('データ読み込みエラー: ' + error.toString());
  }
}

// 設定を取得
function getSettings() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('大会設定');
    
    // シートが存在しない場合は作成
    if (!sheet) {
      Logger.log('「大会設定」シートが存在しないため作成します');
      sheet = ss.insertSheet('大会設定');
      sheet.appendRow(['項目', '値']);
      sheet.appendRow(['大会名', '']);
      sheet.appendRow(['開催日', '']);
      sheet.appendRow(['会場名', '']);
      sheet.appendRow(['第1・2セット点数', '25']);
      sheet.appendRow(['第3セット点数', '15']);
      sheet.appendRow(['デュース差', '2']);
      
      return {
        name: '',
        date: '',
        venue: '',
        set12Points: 25,
        set3Points: 15,
        deuceMargin: 2
      };
    }
    
    const data = sheet.getDataRange().getValues();
    const settings = {};
    
    for (let i = 1; i < data.length; i++) {
      const key = data[i][0];
      const value = data[i][1];
      
      if (key === '大会名') settings.name = value || '';
      if (key === '開催日') settings.date = value || '';
      if (key === '会場名') settings.venue = value || '';
      if (key === '第1・2セット点数') settings.set12Points = parseInt(value) || 25;
      if (key === '第3セット点数') settings.set3Points = parseInt(value) || 15;
      if (key === 'デュース差') settings.deuceMargin = parseInt(value) || 2;
    }
    
    Logger.log('設定読み込み完了: ' + JSON.stringify(settings));
    return settings;
  } catch (error) {
    Logger.log('getSettings エラー: ' + error.toString());
    return {
      name: '',
      date: '',
      venue: '',
      set12Points: 25,
      set3Points: 15,
      deuceMargin: 2
    };
  }
}

// 設定を保存
function saveSettings(settings) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('大会設定');
    
    if (!sheet) {
      sheet = ss.insertSheet('大会設定');
      sheet.appendRow(['項目', '値']);
    }
    
    // 既存データをクリア（ヘッダー以外）
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
    }
    
    // 新しいデータを書き込み
    const data = [
      ['大会名', settings.name || ''],
      ['開催日', settings.date || ''],
      ['会場名', settings.venue || ''],
      ['第1・2セット点数', settings.set12Points || 25],
      ['第3セット点数', settings.set3Points || 15],
      ['デュース差', settings.deuceMargin || 2]
    ];
    
    sheet.getRange(2, 1, data.length, 2).setValues(data);
    Logger.log('設定保存完了');
    return { success: true };
  } catch (error) {
    Logger.log('saveSettings エラー: ' + error.toString());
    throw new Error('設定保存エラー: ' + error.toString());
  }
}

// チームを取得
function getTeams() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('チーム');
    
    // シートが存在しない場合は作成
    if (!sheet) {
      Logger.log('「チーム」シートが存在しないため作成します');
      sheet = ss.insertSheet('チーム');
      sheet.appendRow(['チームID', 'チーム名']);
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const teams = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        teams.push({
          id: data[i][0],
          name: data[i][1] || ''
        });
      }
    }
    
    Logger.log('チーム読み込み完了: ' + teams.length + '件');
    return teams;
  } catch (error) {
    Logger.log('getTeams エラー: ' + error.toString());
    return [];
  }
}

// チームを保存
function saveTeams(teams) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('チーム');
    
    if (!sheet) {
      sheet = ss.insertSheet('チーム');
      sheet.appendRow(['チームID', 'チーム名']);
    } else {
      // 既存データをクリア（ヘッダー以外）
      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
      }
    }
    
    // 新しいデータを書き込み
    const data = teams.map(team => [team.id, team.name]);
    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, 2).setValues(data);
    }
    
    Logger.log('チーム保存完了: ' + teams.length + '件');
    return { success: true };
  } catch (error) {
    Logger.log('saveTeams エラー: ' + error.toString());
    throw new Error('チーム保存エラー: ' + error.toString());
  }
}

// 試合を取得
function getMatches() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('試合');
    
    // シートが存在しない場合は作成
    if (!sheet) {
      Logger.log('「試合」シートが存在しないため作成します');
      sheet = ss.insertSheet('試合');
      sheet.appendRow([
        '試合ID', 'ラウンド', 'ラウンド名', '試合番号',
        'チーム1ID', 'チーム1名', 'チーム2ID', 'チーム2名',
        '審判', 
        '第1セット(チーム1)', '第1セット(チーム2)',
        '第2セット(チーム1)', '第2セット(チーム2)',
        '第3セット(チーム1)', '第3セット(チーム2)',
        'チーム1セット数', 'チーム2セット数',
        '勝者ID', '勝者名', '状態'
      ]);
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const matches = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        matches.push({
          id: data[i][0],
          round: data[i][1],
          roundName: data[i][2],
          matchNumber: data[i][3],
          team1: data[i][4] ? { id: data[i][4], name: data[i][5] } : null,
          team2: data[i][6] ? { id: data[i][6], name: data[i][7] } : null,
          referee: data[i][8] || null,
          sets: [
            { team1Score: data[i][9] || null, team2Score: data[i][10] || null },
            { team1Score: data[i][11] || null, team2Score: data[i][12] || null },
            { team1Score: data[i][13] || null, team2Score: data[i][14] || null }
          ],
          team1Sets: data[i][15] || 0,
          team2Sets: data[i][16] || 0,
          winner: data[i][17] ? { id: data[i][17], name: data[i][18] } : null,
          status: data[i][19] || 'pending'
        });
      }
    }
    
    Logger.log('試合読み込み完了: ' + matches.length + '件');
    return matches;
  } catch (error) {
    Logger.log('getMatches エラー: ' + error.toString());
    return [];
  }
}

// 試合を保存
function saveMatches(matches) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('試合');
    
    if (!sheet) {
      sheet = ss.insertSheet('試合');
      sheet.appendRow([
        '試合ID', 'ラウンド', 'ラウンド名', '試合番号',
        'チーム1ID', 'チーム1名', 'チーム2ID', 'チーム2名',
        '審判', 
        '第1セット(チーム1)', '第1セット(チーム2)',
        '第2セット(チーム1)', '第2セット(チーム2)',
        '第3セット(チーム1)', '第3セット(チーム2)',
        'チーム1セット数', 'チーム2セット数',
        '勝者ID', '勝者名', '状態'
      ]);
    } else {
      // 既存データをクリア（ヘッダー以外）
      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, 20).clearContent();
      }
    }
    
    // 新しいデータを書き込み
    const data = matches.map(match => [
      match.id,
      match.round,
      match.roundName,
      match.matchNumber,
      match.team1 ? match.team1.id : '',
      match.team1 ? match.team1.name : '',
      match.team2 ? match.team2.id : '',
      match.team2 ? match.team2.name : '',
      match.referee || '',
      match.sets[0].team1Score || '',
      match.sets[0].team2Score || '',
      match.sets[1].team1Score || '',
      match.sets[1].team2Score || '',
      match.sets[2].team1Score || '',
      match.sets[2].team2Score || '',
      match.team1Sets || 0,
      match.team2Sets || 0,
      match.winner ? match.winner.id : '',
      match.winner ? match.winner.name : '',
      match.status || 'pending'
    ]);
    
    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, 20).setValues(data);
    }
    
    Logger.log('試合保存完了: ' + matches.length + '件');
    return { success: true };
  } catch (error) {
    Logger.log('saveMatches エラー: ' + error.toString());
    throw new Error('試合保存エラー: ' + error.toString());
  }
}

// 初期設定を保存（setup.html から呼ばれる）
function saveInitialSetup(setupData) {
  try {
    Logger.log('初期設定保存開始');
    Logger.log('受信データ: ' + JSON.stringify(setupData));
    
    // 設定を保存
    saveSettings({
      name: setupData.tournamentName,
      date: setupData.date,
      venue: setupData.venue,
      set12Points: setupData.set12Points,
      set3Points: setupData.set3Points,
      deuceMargin: setupData.deuceMargin
    });
    
    // チームを保存
    saveTeams(setupData.teams);
    
    Logger.log('初期設定保存完了');
    return { success: true };
  } catch (error) {
    Logger.log('saveInitialSetup エラー: ' + error.toString());
    throw new Error('初期設定保存エラー: ' + error.toString());
  }
}