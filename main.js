/**
 * Tournament Pro - メインスクリプト
 * バレーボール大会管理システム
 */

/** スプレッドシートID */
var SPREADSHEET_ID = '1XlrauDVN1xLowyapSW5haIio1LMB9t0Kn2WQuu5jxa4';

/** シート名の定義（全ファイル共通） */
var SHEET = {
  SETTINGS: '大会設定',
  TEAMS: 'チーム一覧',
  MATCHES: '試合表',
  STANDINGS: '順位表',
  TOURNAMENT: 'トーナメント表',
  TASKS: 'タスク管理',
  DOCUMENT: '大会要項',
  ERROR_LOG: 'エラーログ'
};

/** スプレッドシートを取得 */
function getSS_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ========================================
//  Webアプリ エントリーポイント
// ========================================

function doGet(e) {
  e = e || { parameter: {} };
  var page = e.parameter.page || '';
  try {
    var html;
    switch (page) {
      case 'admin':  html = HtmlService.createHtmlOutputFromFile('admin'); break;
      case 'public': html = HtmlService.createHtmlOutputFromFile('public'); break;
      case 'setup':  html = HtmlService.createHtmlOutputFromFile('setup'); break;
      default:       html = HtmlService.createHtmlOutputFromFile('index');
    }
    return html.setTitle('Tournament Pro')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    logError_('doGet', error);
    return HtmlService.createHtmlOutput('<h1>エラー</h1><pre>' + error.toString() + '</pre>');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getCurrentUser() {
  try {
    var email = Session.getActiveUser().getEmail();
    return { email: email, isAuthenticated: email !== '' };
  } catch (error) {
    return { email: '', isAuthenticated: false };
  }
}

// ========================================
//  全データ取得（admin.html, public.html）
// ========================================

function getAllData() {
  try {
    var ss = getSS_();
    ensureSheetsExist_(ss);
    var settings = readSettings_(ss);
    var teams = readTeams_(ss);
    var matches = readMatches_(ss, teams);
    return { settings: settings, teams: teams, matches: matches };
  } catch (error) {
    logError_('getAllData', error);
    throw new Error('データ読み込みエラー: ' + error.toString());
  }
}

// ========================================
//  設定の読み書き
// ========================================

function readSettings_(ss) {
  var sheet = ss.getSheetByName(SHEET.SETTINGS);
  if (!sheet) return { name:'',date:'',venue:'',address:'',startTime:'13:00',endTime:'17:00',set12Points:25,set3Points:15,deuceMargin:2,courtCount:4,matchFormat:'tournament',groupCount:4,teamsAdvancing:2,setFormat:'3set',timeoutCount:2,notes:'' };
  var data = sheet.getDataRange().getValues();
  var m = {};
  for (var i = 1; i < data.length; i++) { if (data[i][0]) m[String(data[i][0])] = data[i][1]; }
  return {
    name: m['大会名']||'', date: m['開催日']||'', venue: m['会場名']||'',
    address: m['会場住所']||'', startTime: m['開始時間']||'13:00', endTime: m['終了時間']||'17:00',
    set12Points: parseInt(m['第1・2セット点数'])||25, set3Points: parseInt(m['第3セット点数'])||15,
    deuceMargin: parseInt(m['デュース差'])||2, courtCount: parseInt(m['コート数'])||4,
    matchFormat: m['試合形式']||'tournament', groupCount: parseInt(m['グループ数'])||4,
    teamsAdvancing: parseInt(m['通過チーム数'])||2, setFormat: m['セット形式']||'3set',
    timeoutCount: parseInt(m['タイムアウト数'])||2, notes: m['備考']||''
  };
}

function saveSettings(settings) {
  try {
    var ss = getSS_();
    var sheet = ss.getSheetByName(SHEET.SETTINGS);
    if (!sheet) sheet = ss.insertSheet(SHEET.SETTINGS);
    sheet.clearContents();
    var rows = [
      ['項目','値'],['大会名',settings.name||''],['開催日',settings.date||''],
      ['会場名',settings.venue||''],['会場住所',settings.address||''],
      ['開始時間',settings.startTime||'13:00'],['終了時間',settings.endTime||'17:00'],
      ['第1・2セット点数',settings.set12Points||25],['第3セット点数',settings.set3Points||15],
      ['デュース差',settings.deuceMargin||2],['コート数',settings.courtCount||4],
      ['試合形式',settings.matchFormat||'tournament'],['グループ数',settings.groupCount||4],
      ['通過チーム数',settings.teamsAdvancing||2],['セット形式',settings.setFormat||'3set'],
      ['タイムアウト数',settings.timeoutCount||2],['備考',settings.notes||'']
    ];
    sheet.getRange(1,1,rows.length,2).setValues(rows);
    formatHeaderRow_(sheet,2);
    sheet.setColumnWidth(1,200); sheet.setColumnWidth(2,300);
    return { success: true };
  } catch (error) {
    logError_('saveSettings', error);
    throw new Error('設定保存エラー: ' + error.toString());
  }
}

// ========================================
//  チームの読み込み
// ========================================

function readTeams_(ss) {
  var sheet = ss.getSheetByName(SHEET.TEAMS);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  var data = sheet.getRange(2,1,lastRow-1,6).getValues();
  var teams = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) {
      teams.push({ id:data[i][0], name:data[i][1]||'', group:data[i][2]||'', rank:data[i][3]||'', contact:data[i][4]||'', paymentStatus:data[i][5]||false });
    }
  }
  return teams;
}

// ========================================
//  試合の読み書き
// ========================================

function readMatches_(ss, teams) {
  var sheet = ss.getSheetByName(SHEET.MATCHES);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  var data = sheet.getRange(2,1,lastRow-1,20).getValues();
  var tm = {};
  teams.forEach(function(t){ tm[t.id]=t.name; });
  var matches = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) {
      var t1=data[i][4], t2=data[i][6], w=data[i][17];
      matches.push({
        id:data[i][0], round:data[i][1], roundName:data[i][2], matchNumber:data[i][3],
        team1: t1?{id:t1,name:data[i][5]||tm[t1]||''}:null,
        team2: t2?{id:t2,name:data[i][7]||tm[t2]||''}:null,
        referee: data[i][8]||null,
        sets:[
          {team1Score:data[i][9]!==''?data[i][9]:null, team2Score:data[i][10]!==''?data[i][10]:null},
          {team1Score:data[i][11]!==''?data[i][11]:null, team2Score:data[i][12]!==''?data[i][12]:null},
          {team1Score:data[i][13]!==''?data[i][13]:null, team2Score:data[i][14]!==''?data[i][14]:null}
        ],
        team1Sets:data[i][15]||0, team2Sets:data[i][16]||0,
        winner: w?{id:w,name:data[i][18]||tm[w]||''}:null,
        status:data[i][19]||'pending'
      });
    }
  }
  return matches;
}

function saveMatches(matches) {
  try {
    var ss = getSS_();
    var sheet = ss.getSheetByName(SHEET.MATCHES);
    if (!sheet) sheet = ss.insertSheet(SHEET.MATCHES);
    sheet.clearContents();
    var headers = ['試合ID','ラウンド','ラウンド名','試合番号','チーム1_ID','チーム1名','チーム2_ID','チーム2名','審判','第1S(T1)','第1S(T2)','第2S(T1)','第2S(T2)','第3S(T1)','第3S(T2)','T1セット数','T2セット数','勝者ID','勝者名','状態'];
    sheet.getRange(1,1,1,20).setValues([headers]);
    if (matches.length > 0) {
      var data = matches.map(function(m){
        var s = m.sets||[{},{},{}];
        return [
          m.id,m.round,m.roundName,m.matchNumber,
          m.team1?m.team1.id:'', m.team1?m.team1.name:'',
          m.team2?m.team2.id:'', m.team2?m.team2.name:'',
          m.referee||'',
          s[0]&&s[0].team1Score!=null?s[0].team1Score:'',
          s[0]&&s[0].team2Score!=null?s[0].team2Score:'',
          s[1]&&s[1].team1Score!=null?s[1].team1Score:'',
          s[1]&&s[1].team2Score!=null?s[1].team2Score:'',
          s[2]&&s[2].team1Score!=null?s[2].team1Score:'',
          s[2]&&s[2].team2Score!=null?s[2].team2Score:'',
          m.team1Sets||0, m.team2Sets||0,
          m.winner?m.winner.id:'', m.winner?m.winner.name:'',
          m.status||'pending'
        ];
      });
      sheet.getRange(2,1,data.length,20).setValues(data);
    }
    formatMatchSheet_(sheet, matches.length);
    return { success: true };
  } catch (error) {
    logError_('saveMatches', error);
    throw new Error('試合保存エラー: ' + error.toString());
  }
}

// ========================================
//  初期設定保存（setup.html）
// ========================================

function saveInitialSetup(setupData) {
  try {
    var ss = getSS_();
    ensureSheetsExist_(ss);
    saveSettings({
      name:setupData.tournamentName, date:setupData.date, venue:setupData.venue,
      address:setupData.address||'', startTime:setupData.startTime||'13:00',
      endTime:setupData.endTime||'17:00', set12Points:setupData.set12Points||25,
      set3Points:setupData.set3Points||15, deuceMargin:setupData.deuceMargin||2,
      matchFormat:setupData.format||'tournament', setFormat:setupData.setFormat||'3set',
      notes:setupData.notes||''
    });
    var teamSheet = ss.getSheetByName(SHEET.TEAMS);
    teamSheet.clearContents();
    teamSheet.getRange(1,1,1,6).setValues([['チームID','チーム名','グループ','ランク','連絡先','入金状態']]);
    if (setupData.teams && setupData.teams.length > 0) {
      var td = setupData.teams.map(function(t){ return [t.id,t.name,'','','',false]; });
      teamSheet.getRange(2,1,td.length,6).setValues(td);
    }
    formatTeamSheet_(teamSheet, setupData.teams?setupData.teams.length:0);
    setupDefaultTasks_(ss);
    return { success: true };
  } catch (error) {
    logError_('saveInitialSetup', error);
    throw new Error('初期設定保存エラー: ' + error.toString());
  }
}

// ========================================
//  シート初期化
// ========================================

function ensureSheetsExist_(ss) {
  var configs = [
    {name:SHEET.SETTINGS, h:['項目','値']},
    {name:SHEET.TEAMS, h:['チームID','チーム名','グループ','ランク','連絡先','入金状態']},
    {name:SHEET.MATCHES, h:['試合ID','ラウンド','ラウンド名','試合番号','チーム1_ID','チーム1名','チーム2_ID','チーム2名','審判','第1S(T1)','第1S(T2)','第2S(T1)','第2S(T2)','第3S(T1)','第3S(T2)','T1セット数','T2セット数','勝者ID','勝者名','状態']},
    {name:SHEET.STANDINGS, h:['グループ','チームID','チーム名','勝点','勝','敗','得セット','失セット','セット率','得点','失点','得失点差','順位']},
    {name:SHEET.TOURNAMENT, h:['ラウンド','マッチNo','チームA_ID','チームA名','チームB_ID','チームB名','勝者ID','勝者名','S1_A','S1_B','S2_A','S2_B','S3_A','S3_B','状態']},
    {name:SHEET.TASKS, h:['No','カテゴリ','タスク','担当者','期限','状態']},
    {name:SHEET.DOCUMENT, h:['項目','値']},
    {name:SHEET.ERROR_LOG, h:['日時','関数名','エラー内容']}
  ];
  configs.forEach(function(c){
    var s = ss.getSheetByName(c.name);
    if (!s) {
      s = ss.insertSheet(c.name);
      s.getRange(1,1,1,c.h.length).setValues([c.h]);
      formatHeaderRow_(s, c.h.length);
      s.setFrozenRows(1);
    }
  });
  migrateOldSheets_(ss);
}

function migrateOldSheets_(ss) {
  [['チーム',SHEET.TEAMS],['試合',SHEET.MATCHES],['Settings',SHEET.SETTINGS],['Teams',SHEET.TEAMS],['Matches',SHEET.MATCHES],['Standings',SHEET.STANDINGS],['Tournament',SHEET.TOURNAMENT],['Tasks',SHEET.TASKS],['TournamentDocument',SHEET.DOCUMENT]].forEach(function(p){
    var o=ss.getSheetByName(p[0]), n=ss.getSheetByName(p[1]);
    if(o&&n&&o.getSheetId()!==n.getSheetId()&&o.getLastRow()>1&&n.getLastRow()<=1){
      var d=o.getDataRange().getValues();
      if(d.length>1){var r=d.slice(1),c=Math.min(r[0].length,n.getLastColumn()||r[0].length);
        var t=r.map(function(x){return x.slice(0,c);});
        if(t.length>0&&c>0) n.getRange(2,1,t.length,c).setValues(t);
      }
    }
  });
}

// ========================================
//  書式設定ヘルパー
// ========================================

function formatHeaderRow_(sheet, colCount) {
  sheet.getRange(1,1,1,colCount).setFontWeight('bold').setBackground('#667eea').setFontColor('white').setHorizontalAlignment('center');
}

function formatMatchSheet_(sheet, rowCount) {
  formatHeaderRow_(sheet,20);
  sheet.setFrozenRows(1);
  var mr = Math.max(rowCount,50);
  var sr = sheet.getRange(2,20,mr,1);
  var rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('completed').setBackground('#c6f6d5').setFontColor('#22543d').setRanges([sr]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ongoing').setBackground('#fefcbf').setFontColor('#744210').setRanges([sr]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('pending').setBackground('#e2e8f0').setFontColor('#4a5568').setRanges([sr]).build());
  sheet.setConditionalFormatRules(rules);
  sr.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['pending','ongoing','completed'],true).setAllowInvalid(false).build());
  sheet.getRange(2,10,mr,6).setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(0,99).setAllowInvalid(true).build());
  sheet.setColumnWidth(1,100); sheet.setColumnWidth(2,100); sheet.setColumnWidth(3,100);
  sheet.setColumnWidth(6,120); sheet.setColumnWidth(8,120); sheet.setColumnWidth(9,100);
}

function formatTeamSheet_(sheet, rowCount) {
  formatHeaderRow_(sheet,6);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1,80); sheet.setColumnWidth(2,200); sheet.setColumnWidth(3,100);
  sheet.setColumnWidth(4,80); sheet.setColumnWidth(5,200); sheet.setColumnWidth(6,100);
  if (rowCount<=0) return;
  var n = Math.max(rowCount,30);
  sheet.getRange(2,3,n,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['','A','B','C','D','E','F'],true).setAllowInvalid(true).build());
  sheet.getRange(2,4,n,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['','A','B','C'],true).setAllowInvalid(true).build());
  var pr = sheet.getRange(2,6,n,1);
  var rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('TRUE').setBackground('#c6f6d5').setRanges([pr]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('FALSE').setBackground('#fed7d7').setRanges([pr]).build());
  sheet.setConditionalFormatRules(rules);
}

function setupDefaultTasks_(ss) {
  var sheet = ss.getSheetByName(SHEET.TASKS);
  if (!sheet||sheet.getLastRow()>1) return;
  var tasks = [
    [1,'事前準備','会場の予約確認','','','未完了'],[2,'事前準備','参加チームへの案内送付','','','未完了'],
    [3,'事前準備','審判員の手配','','','未完了'],[4,'事前準備','ネット・ボール等の備品確認','','','未完了'],
    [5,'事前準備','救急箱・AEDの確認','','','未完了'],[6,'事前準備','参加費の集金','','','未完了'],
    [7,'事前準備','対戦表・トーナメント表の作成','','','未完了'],[8,'事前準備','大会要項の印刷','','','未完了'],
    [9,'当日準備','コート設営（ネット張り・ライン確認）','','','未完了'],[10,'当日準備','受付準備（名簿・筆記用具）','','','未完了'],
    [11,'当日準備','放送設備の確認','','','未完了'],[12,'当日準備','スコアボード・得点板の設置','','','未完了'],
    [13,'当日運営','開会式（挨拶・ルール説明）','','','未完了'],[14,'当日運営','試合進行管理','','','未完了'],
    [15,'当日運営','閉会式・表彰式','','','未完了'],[16,'事後処理','会場の原状復帰・清掃','','','未完了'],
    [17,'事後処理','結果報告の作成・共有','','','未完了'],[18,'事後処理','会計報告の作成','','','未完了']
  ];
  sheet.getRange(2,1,tasks.length,6).setValues(tasks);
  var sr = sheet.getRange(2,6,50,1);
  sr.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['未完了','進行中','完了'],true).setAllowInvalid(false).build());
  var rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('完了').setBackground('#c6f6d5').setFontColor('#22543d').setRanges([sr]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('進行中').setBackground('#fefcbf').setFontColor('#744210').setRanges([sr]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('未完了').setBackground('#fed7d7').setFontColor('#744210').setRanges([sr]).build());
  sheet.setConditionalFormatRules(rules);
}

function initializeSpreadsheet() {
  try {
    var ss = getSS_();
    ensureSheetsExist_(ss);
    setupStandingsSheet_(ss);
    setupDefaultTasks_(ss);
    return { success:true, message:'スプレッドシートを初期化しました' };
  } catch (error) {
    logError_('initializeSpreadsheet', error);
    return { success:false, error:error.toString() };
  }
}

function setupStandingsSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET.STANDINGS);
  if (!sheet) return;
  var h = ['グループ','チームID','チーム名','勝点','勝','敗','得セット','失セット','セット率','得点','失点','得失点差','順位'];
  sheet.getRange(1,1,1,h.length).setValues([h]);
  formatHeaderRow_(sheet,h.length);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1,80); sheet.setColumnWidth(2,80); sheet.setColumnWidth(3,150);
  for(var i=4;i<=13;i++) sheet.setColumnWidth(i,80);
  var rr = sheet.getRange(2,13,50,1);
  var rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(1).setBackground('#fef3c7').setFontColor('#92400e').setRanges([rr]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(2).setBackground('#e0e7ff').setFontColor('#3730a3').setRanges([rr]).build());
  sheet.setConditionalFormatRules(rules);
}

function logError_(functionName, error) {
  try {
    Logger.log('[ERROR] ' + functionName + ': ' + error.toString());
    var ss = getSS_();
    var sheet = ss.getSheetByName(SHEET.ERROR_LOG);
    if (sheet) sheet.appendRow([new Date(), functionName, error.toString()]);
  } catch (e) {
    Logger.log('logError_ failed: ' + e.toString());
  }
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('Tournament Pro')
      .addItem('スプレッドシート初期化','initializeSpreadsheet')
      .addItem('順位表を更新','updateStandingsFromMatches')
      .addItem('審判を自動割り当て','autoAssignReferees')
      .addToUi();
  } catch(e){}
}
