// *************************
// ตั้งค่า CONFIG ที่นี่
// *************************
const SHEET_ID = '1L76maIUpfPLErKzQB4A5rleYYxrBNZiGtwtC1XfJucw'; // ใส่ ID Sheet ของคุณ
const LINE_TOKEN = ' --- ใส่ Line Notify Token ที่นี่ --- ';

// *************************
// CORE FUNCTIONS
// *************************

function doGet(e) {
  let template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
      .setTitle('ACC Maintenance System')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// เชื่อมต่อ Database
function getDbInfo() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return {
    jobSheet: ss.getSheetByName('Jobs'),
    userSheet: ss.getSheetByName('Users')
  };
}

// --- ฟังก์ชัน SETUP (รันครั้งแรกครั้งเดียว) ---
function setupSystem() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // 1. สร้าง Sheet 'Users'
  let userSheet = ss.getSheetByName('Users');
  if (!userSheet) {
    userSheet = ss.insertSheet('Users');
    userSheet.appendRow(['Team Name', 'Passcode', 'Members (Comma separated)']);
    userSheet.setColumnWidth(1, 100); 
    userSheet.setColumnWidth(2, 100); 
    userSheet.setColumnWidth(3, 300);
    
    // ข้อมูลตัวอย่าง
    userSheet.appendRow(['ME1', '1234', 'นาย ก., นาย ข.']);
    userSheet.appendRow(['ME2', '5678', 'นาย ค., นาย ง.']);
    userSheet.appendRow(['ME3', '9012', 'Team Lead, Staff']);
    userSheet.appendRow(['ADMIN', 'admin', 'Admin User']);
  }

  // 2. สร้าง Sheet 'Jobs'
  let jobSheet = ss.getSheetByName('Jobs');
  if (!jobSheet) {
    jobSheet = ss.insertSheet('Jobs');
    jobSheet.appendRow([
      'Job_ID', 'Date', 'Team', 'Supervisor', 'WO_Number', 
      'Description', 'Contractor_Type', 'Contractor_Name', 
      'Plan_Start', 'Plan_Finish', 'Actual_Finish', 'Timestamp', 'Status'
    ]);
  }
}

// ระบบ Login
function loginUser(team, passcode) {
  const db = getDbInfo();
  const data = db.userSheet.getDataRange().getValues();
  // start row 1 to skip header
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == team && data[i][1] == passcode) {
      return { 
        success: true, 
        team: team, 
        members: data[i][2].toString().split(',').map(m => m.trim()) 
      };
    }
  }
  return { success: false };
}

// ดึงรายชื่อทีมทั้งหมดเพื่อไปแสดงใน Dropdown
function getTeamList() {
  const db = getDbInfo();
  const data = db.userSheet.getDataRange().getValues(); 
  const teams = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) { // เอา Column แรกที่เป็นชื่อทีม
      teams.push(data[i][0]);
    }
  }
  return teams;
}

// ดึงข้อมูลงาน
function getJobs(team) {
  const db = getDbInfo();
  const data = db.jobSheet.getDataRange().getValues();
  const headers = data[0];
  const jobs = [];
  // Loop ย้อนหลังเพื่อให้งานใหม่อยู่บน
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    if (team === 'ADMIN' || row[2] === team) {
      let jobObj = {};
      headers.forEach((h, index) => {
        jobObj[h] = row[index];
      });
      jobObj['row_index'] = i + 1;
      jobs.push(jobObj);
    }
  }
  return jobs;
}

// บันทึกงาน
function saveJob(form) {
  const db = getDbInfo();
  const sheet = db.jobSheet;
  const timestamp = new Date();
  const dateStr = Utilities.formatDate(timestamp, "GMT+7", "dd/MM/yyyy");

  try {
    if (form.jobId && form.jobId !== "") {
      // --- UPDATE (แก้ไขงานเดิม) ---
      const data = sheet.getDataRange().getValues();
      let rowIndex = -1;
      for(let i=0; i<data.length; i++){
        if(data[i][0] == form.jobId){
          rowIndex = i + 1;
          break;
        }
      }
      
      if(rowIndex > 0){
        // อัปเดตเฉพาะค่าที่เปลี่ยนได้
        sheet.getRange(rowIndex, 5).setValue(form.woNumber);
        sheet.getRange(rowIndex, 6).setValue(form.description);
        sheet.getRange(rowIndex, 7).setValue(form.contractorType);
        sheet.getRange(rowIndex, 8).setValue(form.contractorName);
        sheet.getRange(rowIndex, 11).setValue(form.actualFinish); 
        sheet.getRange(rowIndex, 12).setValue(timestamp); 
        sheet.getRange(rowIndex, 13).setValue(form.status);
        
        // ข้อมูลใหม่
        sheet.getRange(rowIndex, 15).setValue(form.actualStart);
        sheet.getRange(rowIndex, 16).setValue(form.contractorQty);
        sheet.getRange(rowIndex, 17).setValue(form.spareParts);
        sheet.getRange(rowIndex, 18).setValue(form.externalCost);
      }
      
    } else {
      // --- CREATE (สร้างงานใหม่) ---
      const newId = Utilities.getUuid();
      const newRow = [
        newId, 
        dateStr, 
        form.team, 
        form.supervisor,
        form.woNumber, 
        form.description, 
        form.contractorType,
        form.contractorName, 
        form.planStart, 
        form.planFinish,
        "", // Actual Finish
        timestamp,
        "In Progress",
        // Field ใหม่ที่เพิ่มเข้ามา
        form.userId,        // Col 14
        "",                 // Actual Start (Col 15)
        form.contractorQty, // Col 16
        form.spareParts,    // Col 17
        form.externalCost   // Col 18
      ];
      sheet.appendRow(newRow);
      
      if (form.woNumber === "") {
        const msg = `\n⚠️ *เปิดงานด่วน (No WO)*\nTeam: ${form.team}\nBy: ${form.supervisor}\nJob: ${form.description}`;
        sendLineNotify(msg);
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function sendLineNotify(message) {
  if(!LINE_TOKEN || LINE_TOKEN.includes('ใส่ Line Notify')) return;
  const options = {
    "method": "post",
    "payload": { "message": message },
    "headers": { "Authorization": "Bearer " + LINE_TOKEN }
  };
  UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
