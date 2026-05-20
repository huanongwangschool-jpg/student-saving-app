// ============================================================
// ระบบออมทรัพย์นักเรียน - โรงเรียนบ้านหัวหนองแวง
// Google Apps Script Backend v3.0
// เพิ่ม: AuditLog (ลบไม่ได้), Export data สำหรับ PDF/Excel
// ============================================================
// กฎ: ทุก function ที่แก้ไขข้อมูล ต้อง writeAudit() เสมอ
// กฎ: ทุก function ต้อง return JSON.stringify() เสมอ
// ============================================================

var SHEET_STUDENTS    = "Students";
var SHEET_TX          = "Transactions";
var SHEET_CONFIG      = "Config";
var SHEET_PROMO       = "PromotionHistory";
var SHEET_AUDIT       = "AuditLog";          // ← ใหม่ — ห้ามลบ

var CLASSES = ["อ.2","อ.3","ป.1","ป.2","ป.3","ป.4","ป.5","ป.6","ม.1","ม.2","ม.3"];
var NEXT_CLASS = {
  "อ.2":"อ.3","อ.3":"ป.1","ป.1":"ป.2","ป.2":"ป.3","ป.3":"ป.4",
  "ป.4":"ป.5","ป.5":"ป.6","ป.6":"ม.1","ม.1":"ม.2","ม.2":"ม.3","ม.3":"จบการศึกษา"
};

// ── Entry Point ───────────────────────────────────────────────
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("ระบบออมทรัพย์ โรงเรียนบ้านหัวหนองแวง")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport","width=device-width, initial-scale=1.0");
}

// ── Sheet Helper ──────────────────────────────────────────────
function getOrCreateSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers) sh.appendRow(headers);
  }
  return sh;
}

// ── Audit Logger ──────────────────────────────────────────────
// action: "ADD_STUDENT" | "EDIT_STUDENT" | "TRANSFER_OUT" | "DELETE_STUDENT"
//       | "ADD_TX" | "DELETE_TX" | "PROMOTE"
// detail: free-text string describing what changed
function writeAudit(action, detail, targetId) {
  try {
    var sh  = getOrCreateSheet(SHEET_AUDIT, [
      "LogID","Timestamp","Action","Detail","TargetID","UserEmail"
    ]);
    var lid = "LOG" + new Date().getTime();
    var email = "";
    try { email = Session.getEffectiveUser().getEmail(); } catch(e) {}
    sh.appendRow([
      lid,
      new Date().toISOString(),
      String(action),
      String(detail),
      String(targetId || ""),
      String(email)
    ]);
  } catch(e) {
    // audit fail ไม่ควรทำให้ operation หลักล้มเหลว — log เงียบๆ
    Logger.log("Audit write failed: " + e.message);
  }
}

// ── Init ──────────────────────────────────────────────────────
function initSheets() {
  getOrCreateSheet(SHEET_STUDENTS, [
    "StudentID","Name","Class","Number","Status","StatusNote","CreatedAt"
  ]);
  getOrCreateSheet(SHEET_TX, [
    "TxID","StudentID","Type","Date","Amount","RunningBalance","Note","CreatedAt"
  ]);
  getOrCreateSheet(SHEET_CONFIG, ["Key","Value"]);
  getOrCreateSheet(SHEET_PROMO, [
    "HistID","AcademicYear","StudentID","StudentName","FromClass","ToClass","PromotedAt"
  ]);
  getOrCreateSheet(SHEET_AUDIT, [
    "LogID","Timestamp","Action","Detail","TargetID","UserEmail"
  ]);
  return JSON.stringify({ success: true });
}

// ════════════════════════════════════════════════════════════
// STUDENTS
// ════════════════════════════════════════════════════════════
function getStudents() {
  try {
    var sh   = getOrCreateSheet(SHEET_STUDENTS);
    var rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return JSON.stringify({ success: true, data: [] });
    var data = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0] || r[0] === "") continue;
      data.push({
        id: String(r[0]), name: String(r[1]), className: String(r[2]),
        number: Number(r[3])||0, status: String(r[4]||"active"),
        statusNote: String(r[5]||""), createdAt: r[6]?String(r[6]):""
      });
    }
    return JSON.stringify({ success: true, data: data });
  } catch(e) { return JSON.stringify({ success: false, error: e.message }); }
}

function addStudent(dataJson) {
  try {
    var data = JSON.parse(dataJson);
    var sh   = getOrCreateSheet(SHEET_STUDENTS);
    var id   = "S" + new Date().getTime();
    sh.appendRow([id, String(data.name), String(data.className),
                  Number(data.number), "active", "", new Date().toISOString()]);
    writeAudit("ADD_STUDENT",
      "เพิ่มนักเรียน: " + data.name + " ชั้น " + data.className + " เลขที่ " + data.number, id);
    return JSON.stringify({ success: true, id: id });
  } catch(e) { return JSON.stringify({ success: false, error: e.message }); }
}

function updateStudent(dataJson) {
  try {
    var data = JSON.parse(dataJson);
    var sh   = getOrCreateSheet(SHEET_STUDENTS);
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        var oldInfo = rows[i][1]+" ("+rows[i][2]+")";
        sh.getRange(i+1,2).setValue(String(data.name));
        sh.getRange(i+1,3).setValue(String(data.className));
        sh.getRange(i+1,4).setValue(Number(data.number));
        writeAudit("EDIT_STUDENT",
          "แก้ไข: [เดิม] "+oldInfo+" → [ใหม่] "+data.name+" ("+data.className+")", data.id);
        return JSON.stringify({ success: true });
      }
    }
    return JSON.stringify({ success: false, error: "ไม่พบนักเรียน" });
  } catch(e) { return JSON.stringify({ success: false, error: e.message }); }
}

function transferOutStudent(dataJson) {
  try {
    var data = JSON.parse(dataJson);
    var sh   = getOrCreateSheet(SHEET_STUDENTS);
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sh.getRange(i+1,5).setValue("transferred");
        sh.getRange(i+1,6).setValue(String(data.note||"ย้ายออก"));
        writeAudit("TRANSFER_OUT",
          "นักเรียนย้ายออก: "+rows[i][1]+" ("+rows[i][2]+") — "+data.note, data.id);
        return JSON.stringify({ success: true });
      }
    }
    return JSON.stringify({ success: false, error: "ไม่พบนักเรียน" });
  } catch(e) { return JSON.stringify({ success: false, error: e.message }); }
}

function deleteStudentPermanent(id) {
  try {
    var sid  = String(id);
    var sh   = getOrCreateSheet(SHEET_STUDENTS);
    var rows = sh.getDataRange().getValues();
    var name = "", cls = "";
    for (var i = rows.length-1; i >= 1; i--) {
      if (String(rows[i][0]) === sid) {
        name = rows[i][1]; cls = rows[i][2];
        sh.deleteRow(i+1); break;
      }
    }
    var txSh   = getOrCreateSheet(SHEET_TX);
    var txRows = txSh.getDataRange().getValues();
    var txCount = 0;
    for (var j = txRows.length-1; j >= 1; j--) {
      if (String(txRows[j][1]) === sid) { txSh.deleteRow(j+1); txCount++; }
    }
    writeAudit("DELETE_STUDENT",
      "ลบนักเรียนถาวร: "+name+" ("+cls+") พร้อม "+txCount+" รายการธุรกรรม", sid);
    return JSON.stringify({ success: true });
  } catch(e) { return JSON.stringify({ success: false, error: e.message }); }
}

// ════════════════════════════════════════════════════════════
// TRANSACTIONS
// ════════════════════════════════════════════════════════════
function _calcBalance(studentId) {
  var sh = getOrCreateSheet(SHEET_TX), rows = sh.getDataRange().getValues(), bal = 0;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(studentId)) {
      var amt = Number(rows[i][4])||0;
      bal += String(rows[i][2])==="DEPOSIT" ? amt : -amt;
    }
  }
  return bal;
}

function addTransaction(dataJson) {
  try {
    var data   = JSON.parse(dataJson);
    var type   = String(data.type||"DEPOSIT");
    var amount = Number(data.amount);
    if (!amount||amount<=0) return JSON.stringify({success:false,error:"จำนวนเงินต้องมากกว่า 0"});
    var currentBal = _calcBalance(data.studentId);
    if (type==="WITHDRAW"&&amount>currentBal)
      return JSON.stringify({success:false,error:"ยอดเงินไม่พอ (คงเหลือ "+currentBal.toLocaleString()+" บาท)"});
    var sh     = getOrCreateSheet(SHEET_TX);
    var id     = (type==="DEPOSIT"?"DEP":"WDR")+new Date().getTime();
    var newBal = type==="DEPOSIT" ? currentBal+amount : currentBal-amount;
    sh.appendRow([id, String(data.studentId), type, String(data.date),
                  amount, newBal, String(data.note||""), new Date().toISOString()]);

    // ดึงชื่อนักเรียนเพื่อ audit
    var stuSh = getOrCreateSheet(SHEET_STUDENTS), stuRows = stuSh.getDataRange().getValues();
    var stuName = data.studentId;
    for (var k=1;k<stuRows.length;k++) {
      if (String(stuRows[k][0])===String(data.studentId)) { stuName=stuRows[k][1]; break; }
    }
    var actionLabel = type==="DEPOSIT"?"เพิ่มรายการฝาก":"เพิ่มรายการถอน";
    writeAudit("ADD_TX",
      actionLabel+": "+stuName+" จำนวน ฿"+amount.toLocaleString()+" ("+data.date+") หมายเหตุ: "+(data.note||"-"), id);
    return JSON.stringify({ success: true, id: id, balance: newBal });
  } catch(e) { return JSON.stringify({ success: false, error: e.message }); }
}

function getTransactions() {
  try {
    var sh = getOrCreateSheet(SHEET_TX), rows = sh.getDataRange().getValues();
    if (rows.length<=1) return JSON.stringify({success:true,data:[]});
    var data=[];
    for (var i=1;i<rows.length;i++) {
      var r=rows[i]; if(!r[0]||r[0]==="") continue;
      data.push({id:String(r[0]),studentId:String(r[1]),type:String(r[2]),
        date:String(r[3]),amount:Number(r[4])||0,balance:Number(r[5])||0,
        note:String(r[6]||""),createdAt:r[7]?String(r[7]):""});
    }
    return JSON.stringify({success:true,data:data});
  } catch(e) { return JSON.stringify({success:false,error:e.message}); }
}

function deleteTransaction(id) {
  try {
    var tid=String(id), sh=getOrCreateSheet(SHEET_TX), rows=sh.getDataRange().getValues();
    var targetRow=-1, targetSid="", txInfo="";
    for (var i=1;i<rows.length;i++) {
      if (String(rows[i][0])===tid) {
        targetRow=i+1; targetSid=String(rows[i][1]);
        txInfo=rows[i][2]+" ฿"+rows[i][4]+" วันที่ "+rows[i][3]; break;
      }
    }
    if (targetRow<0) return JSON.stringify({success:false,error:"ไม่พบรายการ"});
    sh.deleteRow(targetRow);
    var newRows=sh.getDataRange().getValues(), run=0;
    for (var j=1;j<newRows.length;j++) {
      if (String(newRows[j][1])===targetSid) {
        var amt=Number(newRows[j][4])||0;
        run += String(newRows[j][2])==="DEPOSIT"?amt:-amt;
        sh.getRange(j+1,6).setValue(run);
      }
    }
    writeAudit("DELETE_TX","ลบรายการธุรกรรม: "+txInfo, tid);
    return JSON.stringify({success:true});
  } catch(e) { return JSON.stringify({success:false,error:e.message}); }
}

// ════════════════════════════════════════════════════════════
// PROMOTION
// ════════════════════════════════════════════════════════════
function promoteStudents(dataJson) {
  try {
    var data=JSON.parse(dataJson), sh=getOrCreateSheet(SHEET_STUDENTS);
    var histSh=getOrCreateSheet(SHEET_PROMO), rows=sh.getDataRange().getValues();
    var year=Number(data.academicYear)||(new Date().getFullYear()+543);
    var filter=data.classFilter||null, excl=data.excludeIds||[];
    var promoted=0,graduated=0,skipped=0,log=[];
    for (var i=1;i<rows.length;i++) {
      var id=String(rows[i][0]),name=String(rows[i][1]),curCls=String(rows[i][2]),status=String(rows[i][4]||"active");
      if(!id||id==="") continue;
      if(status!=="active"){skipped++;continue;}
      if(excl.indexOf(id)>=0){skipped++;continue;}
      if(filter&&curCls!==filter) continue;
      var nextCls=NEXT_CLASS[curCls];
      if(!nextCls){skipped++;continue;}
      if(nextCls==="จบการศึกษา"){
        sh.getRange(i+1,3).setValue("จบการศึกษา");
        sh.getRange(i+1,5).setValue("graduated");
        graduated++;
      } else {
        sh.getRange(i+1,3).setValue(nextCls);
        promoted++;
      }
      var hid="PROM"+new Date().getTime()+"_"+i;
      histSh.appendRow([hid,year,id,name,curCls,nextCls||"จบการศึกษา",new Date().toISOString()]);
      log.push({id:id,name:name,from:curCls,to:nextCls||"จบการศึกษา"});
    }
    writeAudit("PROMOTE",
      "เลื่อนชั้นปีการศึกษา "+year+": เลื่อน "+promoted+" คน, จบ "+graduated+" คน, ข้าม "+skipped+" คน","BATCH");
    return JSON.stringify({success:true,promoted:promoted,graduated:graduated,skipped:skipped,log:log});
  } catch(e) { return JSON.stringify({success:false,error:e.message}); }
}

function getPromotionHistory() {
  try {
    var sh=getOrCreateSheet(SHEET_PROMO), rows=sh.getDataRange().getValues();
    if(rows.length<=1) return JSON.stringify({success:true,data:[]});
    var data=[];
    for(var i=1;i<rows.length;i++){
      var r=rows[i]; if(!r[0]||r[0]==="") continue;
      data.push({id:String(r[0]),academicYear:Number(r[1])||0,studentId:String(r[2]),
        studentName:String(r[3]),fromClass:String(r[4]),toClass:String(r[5]),
        promotedAt:r[6]?String(r[6]):""});
    }
    return JSON.stringify({success:true,data:data});
  } catch(e) { return JSON.stringify({success:false,error:e.message}); }
}

// ════════════════════════════════════════════════════════════
// AUDIT LOG — อ่านได้, แก้ไข/ลบไม่ได้จาก frontend
// ════════════════════════════════════════════════════════════
function getAuditLog() {
  try {
    var sh=getOrCreateSheet(SHEET_AUDIT), rows=sh.getDataRange().getValues();
    if(rows.length<=1) return JSON.stringify({success:true,data:[]});
    var data=[];
    for(var i=1;i<rows.length;i++){
      var r=rows[i]; if(!r[0]||r[0]==="") continue;
      data.push({
        id:        String(r[0]),
        timestamp: r[1]?String(r[1]):"",
        action:    String(r[2]||""),
        detail:    String(r[3]||""),
        targetId:  String(r[4]||""),
        userEmail: String(r[5]||"")
      });
    }
    // คืนเรียงล่าสุดก่อน
    data.reverse();
    return JSON.stringify({success:true,data:data});
  } catch(e) { return JSON.stringify({success:false,error:e.message}); }
}

// ════════════════════════════════════════════════════════════
// LOAD ALL — single round-trip
// ════════════════════════════════════════════════════════════
function loadAllData() {
  try {
    var stuRes  = JSON.parse(getStudents());
    var txRes   = JSON.parse(getTransactions());
    return JSON.stringify({
      success:      true,
      students:     stuRes.success ? stuRes.data : [],
      transactions: txRes.success  ? txRes.data  : []
    });
  } catch(e) { return JSON.stringify({success:false,error:e.message}); }
}
function test() {
  Logger.log("Hello Apps Script");
}
function hello() {
  Logger.log("Push Test");
}