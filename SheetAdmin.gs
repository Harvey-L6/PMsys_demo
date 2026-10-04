/**
 * ============================================================================
 * [全局變數設定] 
 * ============================================================================
 */
// 💡 記得去後台設定 WebURL！
const GLOBAL_WEBAPP_URL = PropertiesService.getScriptProperties().getProperty('WebURL');

// 💡 記得去後台設定 FolderID！
const GLOBAL_PARENT_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('FolderID');

/**
 * ============================================================================
 * [試算表自訂選單與初始化功能]
 * ============================================================================
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ 專案管理工具')
    .addItem('一鍵初始化新專案 (產生 Token/資料夾/網址)', 'initNewProjects')
    .addSeparator() 
    .addItem('發送：立案通知', 'sendInitiationEmails')
    .addItem('發送：回報提醒', 'sendReminderEmails')
    .addItem('發送：報告通知', 'sendReportEmails')
    .addToUi();
}

/**
 * ============================================================================
 * [寄信核心自動化模組]
 * ============================================================================
 */

function sendInitiationEmails() { processEmails_(1); }
function sendReminderEmails() { processEmails_(2); }
function sendReportEmails() { processEmails_(3); }

function processEmails_(emailType) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName('專案追蹤總表');
  const logSheet = ss.getSheetByName('寄件紀錄');
  const teamSheet = ss.getSheetByName('組員清冊');
  const empSheet = ss.getSheetByName('員工主檔');
  
  if (!mainSheet || !logSheet || !empSheet) {
    ui.alert("❌ 找不到必備的分頁，請確認總表、寄件紀錄、員工主檔分頁是否存在。");
    return;
  }

  let empData = empSheet.getDataRange().getDisplayValues();
  let empMap = {}; 
  let empHeaders = getHeaderMap_(empData[0]);
  for (let i = 1; i < empData.length; i++) {
    let name = empData[i][empHeaders['人員姓名']]?.trim();
    let email = empData[i][empHeaders['mail']]?.trim();
    let title = empData[i][empHeaders['職稱']]?.trim() || "";
    if (name && email) empMap[name] = { email: email, title: title };
  }

  let teamMap = {}; 
  if (teamSheet) {
    let teamData = teamSheet.getDataRange().getDisplayValues();
    let teamHeaders = getHeaderMap_(teamData[0]);
    for (let i = 1; i < teamData.length; i++) {
      let pid = teamData[i][teamHeaders['專案編號']]?.trim();
      let mName = teamData[i][teamHeaders['成員姓名']]?.trim();
      if (pid && mName) {
        if (!teamMap[pid]) teamMap[pid] = [];
        teamMap[pid].push(mName);
      }
    }
  }

  let targetColName = "";
  let typeName = "";
  if (emailType === 1) { targetColName = "立案通知"; typeName = "立案通知"; }
  else if (emailType === 2) { targetColName = "回報提醒"; typeName = "回報提醒"; }
  else if (emailType === 3) { targetColName = "報告通知"; typeName = "報告通知"; }

  const mainData = mainSheet.getDataRange().getValues(); 
  const headers = mainData[0];
  const headerMap = getHeaderMap_(headers);
  const targetColIdx = headerMap[targetColName];
  
  if (targetColIdx === undefined) {
    ui.alert(`❌ 總表中找不到『${targetColName}』欄位！`);
    return;
  }

  let sentCount = 0;
  let failCount = 0;
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");

  for (let i = 1; i < mainData.length; i++) {
    let projId = mainData[i][headerMap['專案編號']]?.toString().trim();
    if (!projId) continue;

    let status = mainData[i][targetColIdx]?.toString().trim();
    if (status !== "") continue; 

    let projName = mainData[i][headerMap['專案名稱']]?.toString().trim();
    let ownerName = mainData[i][headerMap['Owner (一級主管指派)']]?.toString().trim();
    let token = mainData[i][headerMap['專案Token']]?.toString().trim();
    
    // 統一讀取全域變數組合網址
    let projectUrl = GLOBAL_WEBAPP_URL + "?token=" + token;

    let ownerInfo = empMap[ownerName];
    let ownerEmail = ownerInfo ? ownerInfo.email : undefined;
    let ownerTitle = ownerInfo ? ownerInfo.title : "";
    
    if (!ownerName || !ownerEmail) {
      mainSheet.getRange(i + 1, targetColIdx + 1).setValue("失敗(缺Owner信箱)");
      appendToLogSheet_(logSheet, {
        '專案編號': projId,
        '專案名稱': projName,
        '信件種類': typeName,
        '收件者': ownerName || "未指派Owner",
        '寄件結果': "失敗 (找不到Owner或信箱缺失)",
        '寄件時間': nowStr
      });
      failCount++;
      continue;
    }

    let toEmails = ownerEmail;
    let ccEmails = "";
    let recipientsLog = `${ownerName} (${ownerEmail})`;
    let resultMsg = "成功";
    let subject = "";
    let body = "";

    let members = teamMap[projId] || [];
    let validCCs = [];
    let missingCCs = [];
    
    for (let m of members) {
      let mInfo = empMap[m];
      if (mInfo && mInfo.email) {
        validCCs.push(mInfo.email);
        recipientsLog += `\n${m} (${mInfo.email})`;
      } else {
        missingCCs.push(m);
      }
    }
    
    if (validCCs.length > 0) ccEmails = validCCs.join(",");
    if (missingCCs.length > 0) resultMsg = `成功 (已略過缺信箱之組員: ${missingCCs.join(", ")})`;

    let greeting = ownerTitle ? `${ownerName} ${ownerTitle} 以及 專案團隊成員，您們好，` : `${ownerName} 以及 專案團隊成員，您們好，`;

    if (emailType === 1) {
      let requester = mainData[i][headerMap['需求提出者']]?.toString().trim() || '未填寫';
      subject = `【立案通知】${projId} ${projName} 已於系統立案，於此信件提供專屬連結網址`;
      body = `
        <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 16px; color: #202124; line-height: 1.6; max-width: 600px;">
          <p style="font-size: 16px; font-weight: bold; margin-bottom: 5px;">${greeting}</p>
          <p style="margin-top: 0;">下面提供專案相關訊息：</p>
          <div style="margin-top: 20px;">
            <div style="color: #1a73e8; font-weight: bold; font-size: 15px; margin-bottom: 8px;">▍ 專案資訊摘要</div>
            <div style="padding-left: 15px; border-left: 4px solid #e8eaed; margin-left: 5px;">
              <span style="color: #5f6368;">專案編號：</span><strong style="color: #202124;">${projId}</strong><br>
              <span style="color: #5f6368;">專案名稱：</span><strong style="color: #202124;">${projName}</strong><br>
              <span style="color: #5f6368;">需求提出者：</span><strong style="color: #202124;">${requester}</strong><br>
              <span style="color: #5f6368;">專案 Owner：</span><strong style="color: #202124;">${ownerName}</strong><br>
              <span style="font-size: 13px; color: #80868b;">(更多詳細訊息請至系統查看)</span>
            </div>
          </div>
          <div style="margin-top: 20px;">
            <div style="color: #1a73e8; font-weight: bold; font-size: 15px; margin-bottom: 8px;">▍ 專案系統網址</div>
            <div style="padding-left: 15px; border-left: 4px solid #e8eaed; margin-left: 5px;">
              <a href="${projectUrl}" style="color: #1a73e8; word-break: break-all;">${projectUrl}</a><br>
              <span style="font-size: 13px; color: #80868b;">(該專案系統將作為後續進度回報與任務管理之用)</span>
            </div>
          </div>
          <div style="margin-top: 20px;">
            <div style="color: #d93025; font-weight: bold; font-size: 15px; margin-bottom: 8px;">▍ 需完成事項</div>
            <div style="padding-left: 15px; border-left: 4px solid #fce8e6; margin-left: 5px;">
              <ol style="margin-top: 0; margin-bottom: 0; padding-left: 20px;">
                <li><strong>確認系統連結：</strong>確認上述網址是否可以正常使用，並且對應專案是否正確。</li>
                <li><strong>維護基本內容：</strong>如下列專案基本資訊有缺漏或錯誤，再請協助至系統上補充、修正：<br>
                <strong style="color: #d93025;">預期 Deadline、預期效益 (目的)、驗收標準、ROI、Milestone(里程碑)、組員清冊</strong></li>
              </ol>
            </div>
          </div>
          <div style="margin-top: 30px; border-top: 1px solid #e8eaed; padding-top: 15px; font-size: 15px; color: #5f6368;">
            <p>如有任何問題，可直接回覆此信件，或透過其他手段聯繫管理者，謝謝。</p>
          </div>
        </div>
      `;
    }
    else if (emailType === 2) {
      let reportDeadline = formatDateWithDays_(mainData[i][headerMap['回報Deadline']]);
      let nextReportDate = formatDateWithDays_(mainData[i][headerMap['下次進度報告日期']]);
      subject = `【回報提醒】${projId} ${projName} 專案進度回報提醒`;
      body = `
        <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 16px; color: #202124; line-height: 1.6; max-width: 600px;">
          <p style="font-size: 16px; font-weight: bold; margin-bottom: 5px;">${greeting}</p>
          <p style="margin-top: 0;">本專案的時程如下：</p>
          <div style="margin-top: 20px;">
            <div style="color: #1a73e8; font-weight: bold; font-size: 15px; margin-bottom: 8px;">▍ 專案時程資訊</div>
            <div style="padding-left: 15px; border-left: 4px solid #e8eaed; margin-left: 5px;">
              <ul style="margin-top: 0; margin-bottom: 0; padding-left: 20px;">
                <li><strong>回報 Deadline：</strong> ${reportDeadline}</li>
                <li><strong>下次進度報告日期：</strong> ${nextReportDate}</li>
              </ul>
            </div>
          </div>
          <div style="margin-top: 20px;">
            <div style="color: #d93025; font-weight: bold; font-size: 15px; margin-bottom: 8px;">▍ 需完成事項</div>
            <div style="padding-left: 15px; border-left: 4px solid #fce8e6; margin-left: 5px;">
              再請於 Deadline 之前，完成 <strong>最新進度回報</strong> 與 <strong>任務狀態更新</strong>，謝謝。
            </div>
          </div>
          <div style="margin-top: 20px;">
            <div style="color: #1a73e8; font-weight: bold; font-size: 15px; margin-bottom: 8px;">▍ 專案系統網址</div>
            <div style="padding-left: 15px; border-left: 4px solid #e8eaed; margin-left: 5px;">
              <a href="${projectUrl}" style="color: #1a73e8; word-break: break-all;">${projectUrl}</a>
            </div>
          </div>
          <div style="margin-top: 30px; border-top: 1px solid #e8eaed; padding-top: 15px; font-size: 15px; color: #5f6368;">
            <p>如有任何問題，可直接回覆此信件，或透過其他手段聯繫管理者，謝謝。</p>
          </div>
        </div>
      `;
    } 
    else if (emailType === 3) {
      let reportDeadline = formatDateWithDays_(mainData[i][headerMap['回報Deadline']]);
      let nextReportDate = formatDateWithDays_(mainData[i][headerMap['下次進度報告日期']]);
      let rawNextReportDate = mainData[i][headerMap['下次進度報告日期']];
      let cleanTitleDate = (rawNextReportDate instanceof Date) ? Utilities.formatDate(rawNextReportDate, Session.getScriptTimeZone(), "yyyy/MM/dd") : (rawNextReportDate?.toString().trim() || '未定');
      subject = `【報告通知】${projId} ${projName} 安排於 ${cleanTitleDate} 進行專案報告`;
      body = `
        <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 16px; color: #202124; line-height: 1.6; max-width: 600px;">
          <p style="font-size: 16px; font-weight: bold; margin-bottom: 5px;">${greeting}</p>
          <p style="margin-top: 0;">本專案下次進度報告日期為 <strong>${nextReportDate}</strong>，再請於回報Deadline: <strong>${reportDeadline}</strong> 之前完成進度與任務更新<span style="color: #80868b;">(如已完成更新，請無視)</span>，並提前準備報告內容。</p>
          <p style="margin-top: 15px;">當日請記得出席專案追蹤會議，謝謝。<br>
          <span style="font-size: 13px; color: #80868b;">(註：相關者必須出席，無關者可不出席)</span></p>
          <div style="margin-top: 20px;">
            <div style="color: #1a73e8; font-weight: bold; font-size: 15px; margin-bottom: 8px;">▍ 專案系統網址</div>
            <div style="padding-left: 15px; border-left: 4px solid #e8eaed; margin-left: 5px;">
              <a href="${projectUrl}" style="color: #1a73e8; word-break: break-all;">${projectUrl}</a>
            </div>
          </div>
          <div style="margin-top: 30px; border-top: 1px solid #e8eaed; padding-top: 15px; font-size: 15px; color: #5f6368;">
            <p>如有任何問題，可直接回覆此信件，或透過其他手段聯繫管理者，謝謝。</p>
          </div>
        </div>
      `;
    }

    try {
      if (ccEmails) {
        GmailApp.sendEmail(toEmails, subject, "", { htmlBody: body, cc: ccEmails, name: "專案追蹤系統" });
      } else {
        GmailApp.sendEmail(toEmails, subject, "", { htmlBody: body, name: "專案追蹤系統" });
      }
      mainSheet.getRange(i + 1, targetColIdx + 1).setValue(`已發送 ${nowStr}`);
      
      appendToLogSheet_(logSheet, {
        '專案編號': projId,
        '專案名稱': projName,
        '信件種類': typeName,
        '收件者': recipientsLog,
        '寄件結果': resultMsg,
        '寄件時間': nowStr
      });
      sentCount++;
    } catch (e) {
      mainSheet.getRange(i + 1, targetColIdx + 1).setValue("失敗(系統發信異常)");
      
      appendToLogSheet_(logSheet, {
        '專案編號': projId,
        '專案名稱': projName,
        '信件種類': typeName,
        '收件者': recipientsLog,
        '寄件結果': `失敗 (${e.message})`,
        '寄件時間': nowStr
      });
      failCount++;
    }
  }

  ui.alert(`✅ 發送作業完成！\n\n- 成功寄出：${sentCount} 封\n- 發送失敗：${failCount} 封\n\n詳細狀況請至『寄件紀錄』分頁查看。`);
}

function formatDateWithDays_(dateValue) {
  if (!dateValue) return '<span style="color: #d93025;">未填寫</span>';
  let targetDate;
  if (dateValue instanceof Date) {
    targetDate = dateValue;
  } else {
    targetDate = new Date(dateValue);
    if (isNaN(targetDate.getTime())) return dateValue.toString().trim(); 
  }
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const target = new Date(targetDate);
  target.setHours(0, 0, 0, 0);
  
  const diffTime = target.getTime() - today.getTime();
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  
  let dateStr = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
  let dayStr = "";
  if (diffDays > 0) dayStr = `(剩餘 ${diffDays} 天)`;
  else if (diffDays === 0) dayStr = `(今日)`;
  else dayStr = `(已過 ${Math.abs(diffDays)} 天)`;
  return `${dateStr} <strong style="color: #d93025;">${dayStr}</strong>`;
}

/**
 * ============================================================================
 * [一鍵初始化專案]
 * ============================================================================
 */
function initNewProjects() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName('專案追蹤總表');
  
  if (!mainSheet) return;
  const data = mainSheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return;
  
  const headers = data[0];
  const headerMap = getHeaderMap_(headers);
  const projIdIdx = headerMap['專案編號'];
  const projNameIdx = headerMap['專案名稱'];
  const tokenIdx = headerMap['專案Token'];
  const folderIdx = headerMap['雲端資料夾連結'];
  const linkIdx = headerMap['專屬網站連結'];
  
  if (projIdIdx === undefined || tokenIdx === undefined || folderIdx === undefined || linkIdx === undefined) {
    ui.alert("❌ 總表缺少必備欄位 (專案Token / 雲端資料夾連結 / 專屬網站連結)。");
    return;
  }
  
  let parentFolder;
  // 👇 修正：改用我們在上方設定好的全域變數 GLOBAL_PARENT_FOLDER_ID
  try { parentFolder = DriveApp.getFolderById(GLOBAL_PARENT_FOLDER_ID); } catch (e) { return; }
  
  const existingTokens = new Set();
  for (let i = 1; i < data.length; i++) {
    const t = data[i][tokenIdx].toString().trim();
    if (t) existingTokens.add(t);
  }
  
  let tokenAddedCount = 0; let folderCreatedCount = 0; let linkCreatedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const projId = data[i][projIdIdx].toString().trim();
    const projName = data[i][projNameIdx].toString().trim();
    if (projId) {
      let currentToken = data[i][tokenIdx].toString().trim();
      const currentFolder = data[i][folderIdx].toString().trim();
      const currentLink = data[i][linkIdx].toString().trim();

      if (!currentToken) {
        currentToken = generateUniqueToken_(existingTokens);
        existingTokens.add(currentToken);
        mainSheet.getRange(i + 1, tokenIdx + 1).setValue(currentToken);
        tokenAddedCount++;
      }
      
      if (!currentFolder) {
        const folderName = `[${projId}] ${projName}`.trim();
        const newFolder = parentFolder.createFolder(folderName);
        mainSheet.getRange(i + 1, folderIdx + 1).setValue(newFolder.getUrl());
        folderCreatedCount++;
      }

      if (!currentLink && currentToken) {
        const generatedLink = GLOBAL_WEBAPP_URL + "?token=" + currentToken;
        mainSheet.getRange(i + 1, linkIdx + 1).setValue(generatedLink);
        linkCreatedCount++;
      }
    }
  }

  if (tokenAddedCount === 0 && folderCreatedCount === 0 && linkCreatedCount === 0) {
    ui.alert("✅ 檢查完成！所有專案均已有 Token、資料夾與專屬連結。");
  } else {
    ui.alert(`🎉 初始化成功！\n- 新產生 Token：${tokenAddedCount} 筆\n- 新建立資料夾：${folderCreatedCount} 個\n- 新產生專屬連結：${linkCreatedCount} 筆`);
  }
}

function generateUniqueToken_(existingTokens) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let token = '';
  do {
    token = '';
    for (let i = 0; i < 16; i++) token += chars.charAt(Math.floor(Math.random() * chars.length));
  } while (existingTokens.has(token));
  return token;
}

function getHeaderMap_(headers) {
  const map = {};
  headers.forEach((header, index) => { if (header) map[header.toString().trim()] = index; });
  return map;
}

function appendToLogSheet_(sheet, logData) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const newRow = new Array(headers.length).fill("");
  
  headers.forEach((header, index) => {
    const headerName = header.toString().trim();
    if (logData[headerName] !== undefined) {
      newRow[index] = logData[headerName];
    }
  });
  
  sheet.appendRow(newRow);
}