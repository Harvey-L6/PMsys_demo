/**
 * ============================================================================
 * [Web App 系統設定區] 
 * ============================================================================
 */
const WEBAPP_CONFIG = {
  SHEET_NAMES: {
    MAIN: '專案追蹤總表',
    HISTORY: '總表異動紀錄',
    TASK: '任務牆',             
    TASK_HISTORY: '任務牆異動紀錄',
    TEAM: '組員清冊',           
    EMPLOYEE: '員工主檔'        
  },
  HEADERS_MAIN: {
    TOKEN: '專案Token',
    PROJ_ID: '專案編號',
    PROJ_NAME: '專案名稱',
    PROGRESS: '最新進度回報',
    NEXT_ACTION: '下一步Action',
    FILE_LINK: '檔案連結(選填)',
    UPDATE_DATE: '最新更新日期',
    FOLDER_LINK: '雲端資料夾連結',
    OWNER: 'Owner (一級主管指派)' 
  },
  HEADERS_TEAM: {
    PROJ_ID: '專案編號',
    MEMBER_NAME: '成員姓名'
  },
  HEADERS_EMP: {
    NAME: '人員姓名',
    TITLE: '職稱',
    EMAIL: 'mail',
    DEPT: '單位'
  },
  HEADERS_HISTORY: { PROJ_ID: '專案編號', PROJ_NAME: '專案名稱', SERIAL: '異動編號', PURPOSE: '預期效益(目的)', ROI: 'ROI', ACCEPTANCE: '驗收標準', MILESTONE: 'Milestone(里程碑)', DEADLINE: '預期Deadline', PROGRESS: '進度彙報內容', NEXT_ACTION: '下一步Action', FILE_LINK: '檔案連結', UPDATE_DATE: '更新日期' },
  HEADERS_TASK: { PROJ_ID: '專案編號', PROJ_NAME: '專案名稱', TASK_ID: '任務編號', OWNER: 'Owner', CONTENT: '任務內容', STATUS: '狀態', PROGRESS: '任務進度回報', CREATE_DATE: '建立日期', EXPECTED_DEADLINE: '預期Deadline', FINISH_DATE: '完成日期', UPDATE_DATE: '最後編輯日期', LOG_DATE: '異動日期' }
};

/**
 * ============================================================================
 * [Web App 後端核心]
 * ============================================================================
 */

function doGet(e) {
  const token = e.parameter.token;
  if (!token) return HtmlService.createHtmlOutput('<h1>❌ 錯誤：網址缺少驗證碼 (Token)</h1>');

  const projectData = getProjectDataByToken_(token);
  if (!projectData) return HtmlService.createHtmlOutput('<h1>❌ 錯誤：找不到此專案，或連結已失效。</h1>');

  const projId = projectData[WEBAPP_CONFIG.HEADERS_MAIN.PROJ_ID] || '未知專案';
  const tasks = getTasksByProjectId_(projId);
  
  const employees = getEmployeeMasterData_();
  const teamMembers = getTeamMembers_(projId);

  const template = HtmlService.createTemplateFromFile('Index');
  template.data = projectData;
  template.token = token;
  template.tasks = tasks; 
  template.employees = employees;     
  template.teamMembers = teamMembers; 
  template.scriptUrl = PropertiesService.getScriptProperties().getProperty('WebURL');
  template.activeTab = e.parameter.tab || 'overview'; 
  
  return template.evaluate()
      .setTitle(`[${projId}] 專案回報中心`)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getEmployeeMasterData_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WEBAPP_CONFIG.SHEET_NAMES.EMPLOYEE);
  if (!sheet) return {};
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return {};
  
  const headers = data[0];
  const map = getHeaderMap_(headers);
  let empDict = {};
  
  for (let i = 1; i < data.length; i++) {
    let name = data[i][map[WEBAPP_CONFIG.HEADERS_EMP.NAME]]?.trim();
    if (name) {
      empDict[name] = {
        title: data[i][map[WEBAPP_CONFIG.HEADERS_EMP.TITLE]] || '',
        email: data[i][map[WEBAPP_CONFIG.HEADERS_EMP.EMAIL]] || '',
        dept: data[i][map[WEBAPP_CONFIG.HEADERS_EMP.DEPT]] || ''
      };
    }
  }
  return empDict;
}

function getTeamMembers_(projId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WEBAPP_CONFIG.SHEET_NAMES.TEAM);
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const projIdColIdx = headers.indexOf(WEBAPP_CONFIG.HEADERS_TEAM.PROJ_ID);
  const nameColIdx = headers.indexOf(WEBAPP_CONFIG.HEADERS_TEAM.MEMBER_NAME);
  if (projIdColIdx === -1 || nameColIdx === -1) return [];

  let members = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][projIdColIdx] === projId) {
      members.push(data[i][nameColIdx].trim());
    }
  }
  return members;
}

function addTeamMember(token, memberName) {
  try {
    const projData = getProjectDataByToken_(token);
    if (!projData) throw new Error("無效的 Token。");
    const projId = projData[WEBAPP_CONFIG.HEADERS_MAIN.PROJ_ID];
    const owner = (projData[WEBAPP_CONFIG.HEADERS_MAIN.OWNER] || '').trim();

    if (memberName === owner) throw new Error("該人員已經是專案 Owner，不可重複加入為組員！");

    const empMaster = getEmployeeMasterData_();
    if (!empMaster[memberName]) {
      throw new Error("此人員不存在於員工主檔中！請確認姓名，或先新增員工至主檔。");
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WEBAPP_CONFIG.SHEET_NAMES.TEAM);
    const data = sheet.getDataRange().getValues();
    
    if (data.length > 0) {
      const headers = data[0];
      const projIdColIdx = headers.indexOf(WEBAPP_CONFIG.HEADERS_TEAM.PROJ_ID);
      const nameColIdx = headers.indexOf(WEBAPP_CONFIG.HEADERS_TEAM.MEMBER_NAME);
      for (let i = 1; i < data.length; i++) {
        if (data[i][projIdColIdx] === projId && data[i][nameColIdx].toString().trim() === memberName) {
          throw new Error("該人員已經在組員清單中了！");
        }
      }
    }
    
    sheet.appendRow([projId, memberName, ""]);
    return { success: true, message: `已成功將 ${memberName} 加入組員！` };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function removeTeamMember(token, memberName) {
  try {
    const projData = getProjectDataByToken_(token);
    if (!projData) throw new Error("無效的 Token。");
    const projId = projData[WEBAPP_CONFIG.HEADERS_MAIN.PROJ_ID];

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WEBAPP_CONFIG.SHEET_NAMES.TEAM);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) throw new Error("無組員資料。");

    const headers = data[0];
    const projIdColIdx = headers.indexOf(WEBAPP_CONFIG.HEADERS_TEAM.PROJ_ID);
    const nameColIdx = headers.indexOf(WEBAPP_CONFIG.HEADERS_TEAM.MEMBER_NAME);

    let deleted = false;
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][projIdColIdx] === projId && data[i][nameColIdx].toString().trim() === memberName) {
        sheet.deleteRow(i + 1);
        deleted = true;
        break;
      }
    }
    if (!deleted) throw new Error("找不到該名組員。");
    
    return { success: true, message: `已將 ${memberName} 從組員清單移除！` };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function addNewEmployeeMaster(empData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WEBAPP_CONFIG.SHEET_NAMES.EMPLOYEE);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const nameIdx = headers.indexOf(WEBAPP_CONFIG.HEADERS_EMP.NAME);
    for (let i = 1; i < data.length; i++) {
      if (data[i][nameIdx].toString().trim() === empData.name.trim()) {
        throw new Error("此員工姓名已存在主檔中，請直接從選單選擇！若為同名同姓，請聯絡系統管理員於後台加註區別。");
      }
    }

    const newRow = new Array(headers.length).fill("");
    const map = getHeaderMap_(headers);
    if (map[WEBAPP_CONFIG.HEADERS_EMP.NAME] !== undefined) newRow[map[WEBAPP_CONFIG.HEADERS_EMP.NAME]] = empData.name.trim();
    if (map[WEBAPP_CONFIG.HEADERS_EMP.TITLE] !== undefined) newRow[map[WEBAPP_CONFIG.HEADERS_EMP.TITLE]] = empData.title.trim();
    if (map[WEBAPP_CONFIG.HEADERS_EMP.EMAIL] !== undefined) newRow[map[WEBAPP_CONFIG.HEADERS_EMP.EMAIL]] = empData.email.trim();
    if (map[WEBAPP_CONFIG.HEADERS_EMP.DEPT] !== undefined) newRow[map[WEBAPP_CONFIG.HEADERS_EMP.DEPT]] = empData.dept.trim();
    
    sheet.appendRow(newRow);
    return { success: true, message: "員工主檔新增成功！您現在可以將他加入組員了。" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getProjectDataByToken_(token) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WEBAPP_CONFIG.SHEET_NAMES.MAIN);
  const data = sheet.getDataRange().getDisplayValues(); 
  if (data.length <= 1) return null;
  const headers = data[0]; 
  const tokenColIdx = headers.indexOf(WEBAPP_CONFIG.HEADERS_MAIN.TOKEN);
  if (tokenColIdx === -1) return null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][tokenColIdx] === token) {
      let rowData = {};
      for (let j = 0; j < headers.length; j++) {
        const headerName = headers[j].toString().trim();
        if (headerName) rowData[headerName] = data[i][j];
      }
      return rowData;
    }
  }
  return null;
}

function getTasksByProjectId_(projId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WEBAPP_CONFIG.SHEET_NAMES.TASK);
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const projIdColIdx = headers.indexOf(WEBAPP_CONFIG.HEADERS_TASK.PROJ_ID);
  if (projIdColIdx === -1) return [];

  let tasks = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][projIdColIdx] === projId) {
      let taskObj = {};
      for (let j = 0; j < headers.length; j++) {
        const headerName = headers[j].toString().trim();
        if (headerName) taskObj[headerName] = data[i][j];
      }
      tasks.push(taskObj);
    }
  }
  return tasks;
}

function saveTask(token, taskData) {
  try {
    const projectData = getProjectDataByToken_(token);
    if (!projectData) throw new Error("無效的 Token，無法儲存任務。");
    const projId = projectData[WEBAPP_CONFIG.HEADERS_MAIN.PROJ_ID];
    const projName = projectData[WEBAPP_CONFIG.HEADERS_MAIN.PROJ_NAME];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const taskSheet = ss.getSheetByName(WEBAPP_CONFIG.SHEET_NAMES.TASK);
    const logSheet = ss.getSheetByName(WEBAPP_CONFIG.SHEET_NAMES.TASK_HISTORY);
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    let targetTaskId = taskData.taskId;
    
    if (!targetTaskId) {
      targetTaskId = getNextTaskId_(taskSheet, projId);
      const newLogData = {
        [WEBAPP_CONFIG.HEADERS_TASK.PROJ_ID]: projId,
        [WEBAPP_CONFIG.HEADERS_TASK.PROJ_NAME]: projName,
        [WEBAPP_CONFIG.HEADERS_TASK.TASK_ID]: targetTaskId,
        [WEBAPP_CONFIG.HEADERS_TASK.OWNER]: taskData.owner,
        [WEBAPP_CONFIG.HEADERS_TASK.CONTENT]: taskData.content,
        [WEBAPP_CONFIG.HEADERS_TASK.STATUS]: taskData.status,
        [WEBAPP_CONFIG.HEADERS_TASK.PROGRESS]: taskData.progress,
        [WEBAPP_CONFIG.HEADERS_TASK.CREATE_DATE]: nowStr,
        [WEBAPP_CONFIG.HEADERS_TASK.EXPECTED_DEADLINE]: taskData.expectedDeadline,
        [WEBAPP_CONFIG.HEADERS_TASK.FINISH_DATE]: taskData.finishDate,
        [WEBAPP_CONFIG.HEADERS_TASK.UPDATE_DATE]: nowStr,
        [WEBAPP_CONFIG.HEADERS_TASK.LOG_DATE]: nowStr
      };
      appendHistoryLog_(taskSheet, newLogData);
      appendHistoryLog_(logSheet, newLogData);
      return { success: true, message: "任務新增成功！" };
    } else {
      const data = taskSheet.getDataRange().getDisplayValues();
      const headers = data[0];
      const headerMap = getHeaderMap_(headers);
      const taskIdColIdx = headerMap[WEBAPP_CONFIG.HEADERS_TASK.TASK_ID];
      let targetRow = -1;
      let createDateStr = "";
      for (let i = 1; i < data.length; i++) {
        if (data[i][taskIdColIdx] === targetTaskId) {
          targetRow = i + 1;
          createDateStr = data[i][headerMap[WEBAPP_CONFIG.HEADERS_TASK.CREATE_DATE]];
          break;
        }
      }
      if (targetRow === -1) throw new Error("找不到該任務編號。");
      
      if (headerMap[WEBAPP_CONFIG.HEADERS_TASK.OWNER] !== undefined) taskSheet.getRange(targetRow, headerMap[WEBAPP_CONFIG.HEADERS_TASK.OWNER] + 1).setValue(taskData.owner);
      if (headerMap[WEBAPP_CONFIG.HEADERS_TASK.CONTENT] !== undefined) taskSheet.getRange(targetRow, headerMap[WEBAPP_CONFIG.HEADERS_TASK.CONTENT] + 1).setValue(taskData.content);
      if (headerMap[WEBAPP_CONFIG.HEADERS_TASK.STATUS] !== undefined) taskSheet.getRange(targetRow, headerMap[WEBAPP_CONFIG.HEADERS_TASK.STATUS] + 1).setValue(taskData.status);
      if (headerMap[WEBAPP_CONFIG.HEADERS_TASK.PROGRESS] !== undefined) taskSheet.getRange(targetRow, headerMap[WEBAPP_CONFIG.HEADERS_TASK.PROGRESS] + 1).setValue(taskData.progress);
      if (headerMap[WEBAPP_CONFIG.HEADERS_TASK.EXPECTED_DEADLINE] !== undefined) taskSheet.getRange(targetRow, headerMap[WEBAPP_CONFIG.HEADERS_TASK.EXPECTED_DEADLINE] + 1).setValue(taskData.expectedDeadline);
      if (headerMap[WEBAPP_CONFIG.HEADERS_TASK.FINISH_DATE] !== undefined) taskSheet.getRange(targetRow, headerMap[WEBAPP_CONFIG.HEADERS_TASK.FINISH_DATE] + 1).setValue(taskData.finishDate);
      if (headerMap[WEBAPP_CONFIG.HEADERS_TASK.UPDATE_DATE] !== undefined) taskSheet.getRange(targetRow, headerMap[WEBAPP_CONFIG.HEADERS_TASK.UPDATE_DATE] + 1).setValue(nowStr);
      
      appendHistoryLog_(logSheet, {
        [WEBAPP_CONFIG.HEADERS_TASK.PROJ_ID]: projId,
        [WEBAPP_CONFIG.HEADERS_TASK.PROJ_NAME]: projName,
        [WEBAPP_CONFIG.HEADERS_TASK.TASK_ID]: targetTaskId,
        [WEBAPP_CONFIG.HEADERS_TASK.OWNER]: taskData.owner,
        [WEBAPP_CONFIG.HEADERS_TASK.CONTENT]: taskData.content,
        [WEBAPP_CONFIG.HEADERS_TASK.STATUS]: taskData.status,
        [WEBAPP_CONFIG.HEADERS_TASK.PROGRESS]: taskData.progress,
        [WEBAPP_CONFIG.HEADERS_TASK.CREATE_DATE]: createDateStr,
        [WEBAPP_CONFIG.HEADERS_TASK.EXPECTED_DEADLINE]: taskData.expectedDeadline,
        [WEBAPP_CONFIG.HEADERS_TASK.FINISH_DATE]: taskData.finishDate,
        [WEBAPP_CONFIG.HEADERS_TASK.LOG_DATE]: nowStr
      });
      return { success: true, message: "任務更新成功！" };
    }
  } catch (error) { return { success: false, message: error.toString() }; }
}

function updateSingleField(token, columnName, newValue) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheetByName(WEBAPP_CONFIG.SHEET_NAMES.MAIN);
    const historySheet = ss.getSheetByName(WEBAPP_CONFIG.SHEET_NAMES.HISTORY);
    const mainHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    const mainColMap = getHeaderMap_(mainHeaders);
    const tokenColIdx = mainColMap[WEBAPP_CONFIG.HEADERS_MAIN.TOKEN];
    if (tokenColIdx === undefined) throw new Error("總表找不到『專案Token』欄位");
    const mainData = mainSheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < mainData.length; i++) {
      if (mainData[i][tokenColIdx] === token) { targetRow = i + 1; break; }
    }
    if (targetRow === -1) throw new Error("無效的 Token，更新失敗。");
    const projId = mainData[targetRow - 1][mainColMap[WEBAPP_CONFIG.HEADERS_MAIN.PROJ_ID]];
    const projName = mainData[targetRow - 1][mainColMap[WEBAPP_CONFIG.HEADERS_MAIN.PROJ_NAME]];
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    const targetColIdx = mainColMap[columnName];
    if (targetColIdx !== undefined) mainSheet.getRange(targetRow, targetColIdx + 1).setValue(newValue);
    if (mainColMap[WEBAPP_CONFIG.HEADERS_MAIN.UPDATE_DATE] !== undefined) mainSheet.getRange(targetRow, mainColMap[WEBAPP_CONFIG.HEADERS_MAIN.UPDATE_DATE] + 1).setValue(nowStr);
    
    const nextSerial = getNextSerial_(historySheet, projId);
    appendHistoryLog_(historySheet, { [WEBAPP_CONFIG.HEADERS_HISTORY.PROJ_ID]: projId, [WEBAPP_CONFIG.HEADERS_HISTORY.PROJ_NAME]: projName, [WEBAPP_CONFIG.HEADERS_HISTORY.SERIAL]: nextSerial, [columnName]: newValue, [WEBAPP_CONFIG.HEADERS_HISTORY.UPDATE_DATE]: nowStr });
    return { success: true };
  } catch (error) { return { success: false, message: error.toString() }; }
}

function submitProjectUpdate(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheetByName(WEBAPP_CONFIG.SHEET_NAMES.MAIN);
    const historySheet = ss.getSheetByName(WEBAPP_CONFIG.SHEET_NAMES.HISTORY);
    const mainHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    const mainColMap = getHeaderMap_(mainHeaders);
    const token = formData.token;
    const tokenColIdx = mainColMap[WEBAPP_CONFIG.HEADERS_MAIN.TOKEN];
    if (tokenColIdx === undefined) throw new Error("總表找不到『專案Token』欄位");
    const mainData = mainSheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < mainData.length; i++) {
      if (mainData[i][tokenColIdx] === token) { targetRow = i + 1; break; }
    }
    if (targetRow === -1) throw new Error("無效的 Token，寫入失敗。");

    const projId = mainData[targetRow - 1][mainColMap[WEBAPP_CONFIG.HEADERS_MAIN.PROJ_ID]];
    const projName = mainData[targetRow - 1][mainColMap[WEBAPP_CONFIG.HEADERS_MAIN.PROJ_NAME]];
    const folderLink = mainData[targetRow - 1][mainColMap[WEBAPP_CONFIG.HEADERS_MAIN.FOLDER_LINK]];
    
    let finalFileLinks = [];
    if (formData.files && formData.files.length > 0) {
      let targetFolder = getTargetFolder_(folderLink);
      for (let i = 0; i < formData.files.length; i++) {
        let fileObj = formData.files[i];
        let blob = Utilities.newBlob(Utilities.base64Decode(fileObj.base64), fileObj.mimeType, fileObj.name);
        let createdFile = targetFolder.createFile(blob);
        finalFileLinks.push(createdFile.getUrl());
      }
    }
    
    const fileLinksString = finalFileLinks.join("\n");
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

    if (mainColMap[WEBAPP_CONFIG.HEADERS_MAIN.PROGRESS] !== undefined) mainSheet.getRange(targetRow, mainColMap[WEBAPP_CONFIG.HEADERS_MAIN.PROGRESS] + 1).setValue(formData.progress);
    if (mainColMap[WEBAPP_CONFIG.HEADERS_MAIN.NEXT_ACTION] !== undefined) mainSheet.getRange(targetRow, mainColMap[WEBAPP_CONFIG.HEADERS_MAIN.NEXT_ACTION] + 1).setValue(formData.nextAction);
    if (mainColMap[WEBAPP_CONFIG.HEADERS_MAIN.FILE_LINK] !== undefined) mainSheet.getRange(targetRow, mainColMap[WEBAPP_CONFIG.HEADERS_MAIN.FILE_LINK] + 1).setValue(fileLinksString);
    if (mainColMap[WEBAPP_CONFIG.HEADERS_MAIN.UPDATE_DATE] !== undefined) mainSheet.getRange(targetRow, mainColMap[WEBAPP_CONFIG.HEADERS_MAIN.UPDATE_DATE] + 1).setValue(nowStr);

    const nextSerial = getNextSerial_(historySheet, projId);
    appendHistoryLog_(historySheet, { [WEBAPP_CONFIG.HEADERS_HISTORY.PROJ_ID]: projId, [WEBAPP_CONFIG.HEADERS_HISTORY.PROJ_NAME]: projName, [WEBAPP_CONFIG.HEADERS_HISTORY.SERIAL]: nextSerial, [WEBAPP_CONFIG.HEADERS_HISTORY.PROGRESS]: formData.progress, [WEBAPP_CONFIG.HEADERS_HISTORY.NEXT_ACTION]: formData.nextAction, [WEBAPP_CONFIG.HEADERS_HISTORY.FILE_LINK]: fileLinksString, [WEBAPP_CONFIG.HEADERS_HISTORY.UPDATE_DATE]: nowStr });

    return { success: true, message: "資料已成功更新並紀錄至歷史軌跡！", finalFileLinks: finalFileLinks };
  } catch (error) { return { success: false, message: error.toString() }; }
}

function getHeaderMap_(headers) {
  const map = {};
  headers.forEach((header, index) => { if (header) map[header.toString().trim()] = index; });
  return map;
}

function getNextSerial_(historySheet, projId) {
  const data = historySheet.getDataRange().getValues();
  if (data.length <= 1) return projId + "-001";
  const headers = data[0];
  const serialColIdx = headers.findIndex(h => h.toString().trim() === WEBAPP_CONFIG.HEADERS_HISTORY.SERIAL);
  const projIdColIdx = headers.findIndex(h => h.toString().trim() === WEBAPP_CONFIG.HEADERS_HISTORY.PROJ_ID);
  let maxSerialNum = 0;
  const prefix = projId + "-";
  for (let i = 1; i < data.length; i++) {
    const rowProjId = data[i][projIdColIdx]?.toString().trim();
    const serialStr = data[i][serialColIdx]?.toString().trim();
    if (rowProjId === projId && serialStr && serialStr.startsWith(prefix)) {
      const numPart = parseInt(serialStr.substring(prefix.length), 10);
      if (!isNaN(numPart) && numPart > maxSerialNum) maxSerialNum = numPart;
    }
  }
  return prefix + String(maxSerialNum + 1).padStart(3, '0');
}

function getNextTaskId_(taskSheet, projId) {
  const data = taskSheet.getDataRange().getValues();
  if (data.length <= 1) return projId + "-T-001";
  const headers = data[0];
  const serialColIdx = headers.findIndex(h => h.toString().trim() === WEBAPP_CONFIG.HEADERS_TASK.TASK_ID);
  const projIdColIdx = headers.findIndex(h => h.toString().trim() === WEBAPP_CONFIG.HEADERS_TASK.PROJ_ID);
  let maxSerialNum = 0;
  const prefix = projId + "-T-";
  for (let i = 1; i < data.length; i++) {
    const rowProjId = data[i][projIdColIdx]?.toString().trim();
    const serialStr = data[i][serialColIdx]?.toString().trim();
    if (rowProjId === projId && serialStr && serialStr.startsWith(prefix)) {
      const numPart = parseInt(serialStr.substring(prefix.length), 10);
      if (!isNaN(numPart) && numPart > maxSerialNum) maxSerialNum = numPart;
    }
  }
  return prefix + String(maxSerialNum + 1).padStart(3, '0');
}

function appendHistoryLog_(sheet, logData) {
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return;
  const headers = data[0];
  const newRow = new Array(headers.length).fill("");
  headers.forEach((header, index) => {
    const headerName = header.toString().trim();
    if (logData[headerName] !== undefined) newRow[index] = logData[headerName];
  });
  sheet.appendRow(newRow);
}

function getTargetFolder_(folderUrl) {
  try {
    if (folderUrl) {
      let id = "";
      if (folderUrl.includes("id=")) { id = folderUrl.split("id=")[1].split("&")[0]; } 
      else if (folderUrl.includes("/folders/")) { id = folderUrl.split("/folders/")[1].split("?")[0]; }
      if (id) return DriveApp.getFolderById(id);
    }
  } catch (e) {}
  // 👇 改為直接從全域變數讀取預設資料夾 ID
  return DriveApp.getFolderById(GLOBAL_PARENT_FOLDER_ID);
}
