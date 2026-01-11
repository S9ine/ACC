// *************************
// ตั้งค่า CONFIG ที่นี่
// *************************
const SHEET_ID = '1L76maIUpfPLErKzQB4A5rleYYxrBNZiGtwtC1XfJucw';
const LINE_TOKEN = ' --- ใส่ Line Notify Token ที่นี่ --- ';

// *************************
// CORE FUNCTIONS
// *************************

function doGet(e) {
  // รับค่า parameter ?page=tv เพื่อเปิดหน้า TV Mode ทันที (ถ้าต้องการ)
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
        members: data[i][2].split(',').map(m => m.trim()) 
      };
    }
  }
  return { success: false };
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
    // ถ้าเป็น Admin เห็นหมด, ถ้าเป็น Team เห็นแค่ Team ตัวเอง
    if (team === 'ADMIN' || row[2] === team) {
      let jobObj = {};
      headers.forEach((h, index) => {
        jobObj[h] = row[index];
      });
      jobObj['row_index'] = i + 1; // เก็บเลขบรรทัดเพื่อใช้อัพเดท
      jobs.push(jobObj);
    }
  }
  return jobs;
}

// บันทึกงาน (สร้างใหม่ และ อัพเดท)
function saveJob(form) {
  const db = getDbInfo();
  const sheet = db.jobSheet;
  const timestamp = new Date();
  const dateStr = Utilities.formatDate(timestamp, "GMT+7", "dd/MM/yyyy");
  
  try {
    if (form.jobId && form.jobId !== "") {
      // --- กรณีอัพเดท (EDIT / CLOSE) ---
      // หา row จาก jobId (ในที่นี้ใช้ loop หาแบบง่าย หรือใช้ row_index ถ้าส่งมา)
      // เพื่อความชัวร์ จะค้นหาจาก ID ใน Column A
      const data = sheet.getDataRange().getValues();
      let rowIndex = -1;
      for(let i=0; i<data.length; i++){
        if(data[i][0] == form.jobId){
          rowIndex = i + 1;
          break;
        }
      }
      
      if(rowIndex > 0){
        // อัพเดทข้อมูลบางส่วน
        sheet.getRange(rowIndex, 5).setValue(form.woNumber); // Update WO
        sheet.getRange(rowIndex, 11).setValue(form.actualFinish); // Actual Finish
        sheet.getRange(rowIndex, 12).setValue(timestamp); // System Update Time
        sheet.getRange(rowIndex, 13).setValue(form.status); // Status
        
        // ถ้าปิดงาน ให้ส่ง Line แจ้งเตือน (Optional)
        // if(form.status === 'Completed') sendLineNotify(...) 
      }
      
    } else {
      // --- กรณีสร้างใหม่ (CREATE) ---
      const newId = Utilities.getUuid(); // สร้าง ID
      
      // เรียง Data ตาม Column
      const newRow = [
        newId,
        dateStr,
        form.team,
        form.supervisor,
        form.woNumber, // อาจจะว่าง
        form.description,
        form.contractorType,
        form.contractorName,
        form.planStart,
        form.planFinish,
        "", // Actual Finish (ยังไม่มี)
        timestamp,
        "In Progress"
      ];
      
      sheet.appendRow(newRow);
      
      // *** Logic แจ้งเตือน Line เมื่อไม่มี WO ***
      if (form.woNumber === "") {
        const msg = `\n⚠️ *เปิดงานด่วน (No WO)*\nTeam: ${form.team}\nBy: ${form.supervisor}\nJob: ${form.description}\nStatus: In Progress`;
        sendLineNotify(msg);
      }
    }
    
    return { success: true };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ฟังก์ชันส่ง Line
function sendLineNotify(message) {
  if(!LINE_TOKEN) return;
  const options = {
    "method": "post",
    "payload": { "message": message },
    "headers": { "Authorization": "Bearer " + LINE_TOKEN }
  };
  UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
}

// Include files
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ฟังก์ชันดึงรายชื่อทีมทั้งหมดจาก Sheet Users เพื่อไปโชว์หน้า Login
function getTeamList() {
  const db = getDbInfo();
  // ดึงข้อมูลทั้งหมดจาก tab Users
  const data = db.userSheet.getDataRange().getValues(); 
  const teams = [];
  
  // เริ่ม loop ที่ i=1 เพื่อข้าม Header
  for (let i = 1; i < data.length; i++) {
    // column 0 คือชื่อทีม (ME1, ME2...)
    if (data[i][0]) {
      teams.push(data[i][0]);
    }
  }
  return teams;
}
