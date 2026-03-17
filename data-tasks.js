/**
 * Tournament Pro - タスク管理
 */

/**
 * タスクデータを取得
 * @return {Object} タスクデータ
 */
function getTasksData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Tasks');
    
    if (!sheet) {
      return { success: false, error: 'Tasksシートが見つかりません' };
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, data: [] };
    }
    
    var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    var tasks = [];
    
    data.forEach(function(row) {
      if (row[0]) { // Noが存在する行のみ
        tasks.push({
          no: row[0],
          category: row[1],
          task: row[2],
          assignee: row[3],
          deadline: row[4],
          status: row[5]
        });
      }
    });
    
    return {
      success: true,
      data: tasks
    };
  } catch (error) {
    Logger.log('getTasksData Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * タスクを追加
 * @param {Object} task - タスク情報
 * @return {Object} 結果
 */
function addTask(task) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Tasks');
    
    if (!sheet) {
      return { success: false, error: 'Tasksシートが見つかりません' };
    }
    
    var lastRow = sheet.getLastRow();
    var newNo = lastRow; // ヘッダー行を除いた行数が新しいNo
    
    sheet.appendRow([
      newNo,
      task.category || '',
      task.task || '',
      task.assignee || '',
      task.deadline || '',
      task.status || '未完了'
    ]);
    
    return {
      success: true,
      message: 'タスクを追加しました'
    };
  } catch (error) {
    Logger.log('addTask Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * タスクを更新
 * @param {Object} task - タスク情報
 * @return {Object} 結果
 */
function updateTask(task) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Tasks');
    
    if (!sheet) {
      return { success: false, error: 'Tasksシートが見つかりません' };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == task.no) {
        sheet.getRange(i + 1, 2).setValue(task.category || '');
        sheet.getRange(i + 1, 3).setValue(task.task || '');
        sheet.getRange(i + 1, 4).setValue(task.assignee || '');
        sheet.getRange(i + 1, 5).setValue(task.deadline || '');
        sheet.getRange(i + 1, 6).setValue(task.status || '');
        
        return {
          success: true,
          message: 'タスクを更新しました'
        };
      }
    }
    
    return {
      success: false,
      error: 'タスクが見つかりません'
    };
  } catch (error) {
    Logger.log('updateTask Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * タスクの状態を切り替え（完了⇔未完了）
 * @param {number} taskNo - タスクNo
 * @return {Object} 結果
 */
function toggleTaskStatus(taskNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Tasks');
    
    if (!sheet) {
      return { success: false, error: 'Tasksシートが見つかりません' };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == taskNo) {
        var currentStatus = data[i][5];
        var newStatus = (currentStatus === '完了' || currentStatus === '済') ? '未完了' : '完了';
        sheet.getRange(i + 1, 6).setValue(newStatus);
        
        return {
          success: true,
          message: 'タスク状態を変更しました',
          newStatus: newStatus
        };
      }
    }
    
    return {
      success: false,
      error: 'タスクが見つかりません'
    };
  } catch (error) {
    Logger.log('toggleTaskStatus Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * タスクを削除
 * @param {number} taskNo - タスクNo
 * @return {Object} 結果
 */
function deleteTask(taskNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Tasks');
    
    if (!sheet) {
      return { success: false, error: 'Tasksシートが見つかりません' };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == taskNo) {
        sheet.deleteRow(i + 1);
        return {
          success: true,
          message: 'タスクを削除しました'
        };
      }
    }
    
    return {
      success: false,
      error: 'タスクが見つかりません'
    };
  } catch (error) {
    Logger.log('deleteTask Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}
