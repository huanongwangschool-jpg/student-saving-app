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
var SHEET_BANK_DEPOSITS = "MonthlyBankDeposits";
var SHEET_AUDIT_LOGS  = "AuditLogs";
var ADMIN_EMAIL       = "huanongwangschool@gmail.com";
var ADMIN_PASSWORD_SHA256 = "a612350f41b1944eb77c9c31078a48c7004542d2962c712ec3c08debd6547aa6";

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

function _sha256Hex(text) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8);
  return raw.map(function(b){
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? "0" + v : v;
  }).join("");
}

function _ensureAdminProps() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty("ADMIN_EMAIL")) props.setProperty("ADMIN_EMAIL", ADMIN_EMAIL);
  if (!props.getProperty("ADMIN_PASSWORD_SHA256")) props.setProperty("ADMIN_PASSWORD_SHA256", ADMIN_PASSWORD_SHA256);
  return props;
}

function checkAdminLogin(dataJson) {
  try {
    var data = JSON.parse(dataJson || "{}");
    var props = _ensureAdminProps();
    var email = String(data.email || "").trim().toLowerCase();
    var passHash = _sha256Hex(String(data.password || ""));
    if (email !== String(props.getProperty("ADMIN_EMAIL") || "").toLowerCase() ||
        passHash !== String(props.getProperty("ADMIN_PASSWORD_SHA256") || "")) {
      return JSON.stringify({success:false,error:"อีเมลหรือรหัสผ่านผู้ดูแลระบบไม่ถูกต้อง"});
    }
    var token = Utilities.getUuid();
    CacheService.getScriptCache().put("ADMIN_TOKEN_" + token, email, 21600);
    writeAuditLog("ADMIN_LOGIN","AdminSession",email,"","เข้าสู่ระบบผู้ดูแลระบบ",email,"");
    return JSON.stringify({success:true,role:"Admin",email:email,token:token});
  } catch(e) { return JSON.stringify({success:false,error:e.message}); }
}

function _requireAdmin(token) {
  var email = CacheService.getScriptCache().get("ADMIN_TOKEN_" + String(token || ""));
  if (!email) throw new Error("ต้องเข้าสู่ระบบผู้ดูแลระบบก่อน");
  return email;
}

function _auditLogHeaders() {
  return ["logId","actionType","targetType","targetId","oldValue","newValue","editedBy","editedAt","note"];
}

function writeAuditLog(actionType, targetType, targetId, oldValue, newValue, editedBy, note) {
  try {
    var sh = getOrCreateSheet(SHEET_AUDIT_LOGS, _auditLogHeaders());
    sh.appendRow([
      "AUD" + new Date().getTime(),
      String(actionType || ""),
      String(targetType || ""),
      String(targetId || ""),
      typeof oldValue === "string" ? oldValue : JSON.stringify(oldValue || ""),
      typeof newValue === "string" ? newValue : JSON.stringify(newValue || ""),
      String(editedBy || _getUserEmailSafe()),
      new Date().toISOString(),
      String(note || "")
    ]);
  } catch(e) {
    Logger.log("AuditLogs write failed: " + e.message);
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
  getOrCreateSheet(SHEET_BANK_DEPOSITS, [
    "depositId","month","year","closedAt","closedBy","totalAmount",
    "totalTransactions","classSummaryJson","note","status","reversedAt","reversedBy","createdAt"
  ]);
  getOrCreateSheet(SHEET_AUDIT_LOGS, [
    "logId","actionType","targetType","targetId","oldValue","newValue","editedBy","editedAt","note"
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
        writeAuditLog("EDIT_STUDENT","Student",data.id,
          {name:rows[i][1],className:rows[i][2],number:rows[i][3]},
          {name:data.name,className:data.className,number:data.number},
          _getUserEmailSafe(),"แก้ไขข้อมูลนักเรียน");
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
    var className = String(data.className || "");
    var teacherName = String(data.teacherName || "");
    if (!amount||amount<=0) return JSON.stringify({success:false,error:"จำนวนเงินต้องมากกว่า 0"});
    var currentBal = _calcBalance(data.studentId);
    if (type==="WITHDRAW"&&amount>currentBal)
      return JSON.stringify({success:false,error:"ยอดเงินไม่พอ (คงเหลือ "+currentBal.toLocaleString()+" บาท)"});
    var sh     = getOrCreateSheet(SHEET_TX);
    var id     = (type==="DEPOSIT"?"DEP":"WDR")+new Date().getTime();
    var newBal = type==="DEPOSIT" ? currentBal+amount : currentBal-amount;
    var note = String(data.note||"");
    if (type === "DEPOSIT") {
      var meta = [];
      if (className) meta.push("ชั้น: " + className);
      if (teacherName) meta.push("ครูผู้ทำรายการ: " + teacherName);
      if (meta.length) note = note ? note + " | " + meta.join(" | ") : meta.join(" | ");
    }
    var createdAt = new Date().toISOString();
    sh.appendRow([id, String(data.studentId), type, String(data.date),
                  amount, newBal, note, createdAt]);

    // ดึงชื่อนักเรียนเพื่อ audit
    var stuSh = getOrCreateSheet(SHEET_STUDENTS), stuRows = stuSh.getDataRange().getValues();
    var stuName = data.studentId;
    for (var k=1;k<stuRows.length;k++) {
      if (String(stuRows[k][0])===String(data.studentId)) { stuName=stuRows[k][1]; break; }
    }
    var actionLabel = type==="DEPOSIT"?"เพิ่มรายการฝาก":"เพิ่มรายการถอน";
    writeAudit("ADD_TX",
      actionLabel+": "+stuName+" จำนวน ฿"+amount.toLocaleString()+" ("+data.date+") หมายเหตุ: "+(data.note||"-"), id);
    writeAuditLog("ADD_TX","Transaction",id,"",
      {studentId:data.studentId,type:type,date:data.date,amount:amount,balance:newBal,className:className,teacherName:teacherName,createdAt:createdAt},
      teacherName || _getUserEmailSafe(), actionLabel);
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
    writeAuditLog("DELETE_TX","Transaction",tid,txInfo,"ลบรายการ",_getUserEmailSafe(),"ลบรายการธุรกรรม");
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
// MONTHLY BANK DEPOSITS — ปิดยอดสิ้นเดือนเพื่อนำเงินเข้าธนาคาร
// ════════════════════════════════════════════════════════════
function _bankHeaders() {
  return [
    "depositId","month","year","closedAt","closedBy","totalAmount",
    "totalTransactions","classSummaryJson","note","status","reversedAt","reversedBy","createdAt"
  ];
}

function _getBankSheet() {
  var sh = getOrCreateSheet(SHEET_BANK_DEPOSITS, _bankHeaders());
  var headers = _bankHeaders();
  var lastCol = sh.getLastColumn ? sh.getLastColumn() : 0;
  if (lastCol < headers.length) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  }
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var status = String(rows[i][9] || "");
    if (!status || status.indexOf("T") > 0 || status.indexOf("-") > 0) {
      var oldCreatedAt = rows[i][9] ? String(rows[i][9]) : "";
      sh.getRange(i+1,10).setValue("ACTIVE");
      if (!rows[i][12]) sh.getRange(i+1,13).setValue(oldCreatedAt || new Date().toISOString());
    }
  }
  return sh;
}

function _parseTxDate(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  var s = String(value || "").trim();
  if (!s) return null;
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) {
    var y = Number(m[1]);
    if (y > 2400) y -= 543;
    return new Date(y, Number(m[2]) - 1, Number(m[3]));
  }
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    var y2 = Number(m[3]);
    if (y2 < 100) y2 += 2000;
    if (y2 > 2400) y2 -= 543;
    return new Date(y2, Number(m[2]) - 1, Number(m[1]));
  }
  var d = new Date(s);
  return isNaN(d) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function _getUserEmailSafe() {
  try { return Session.getEffectiveUser().getEmail() || ""; } catch(e) { return ""; }
}

function _monthAlreadyClosed(month, year) {
  var sh = _getBankSheet();
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var status = String(rows[i][9] || "ACTIVE");
    if (String(rows[i][1]) === String(month) && Number(rows[i][2]) === Number(year) && status === "ACTIVE") {
      return {
        depositId: String(rows[i][0]),
        month: String(rows[i][1]),
        year: Number(rows[i][2]) || 0,
        closedAt: rows[i][3] ? String(rows[i][3]) : "",
        closedBy: String(rows[i][4] || ""),
        totalAmount: Number(rows[i][5]) || 0,
        totalTransactions: Number(rows[i][6]) || 0,
        classSummaryJson: String(rows[i][7] || "[]"),
        note: String(rows[i][8] || ""),
        status: status,
        reversedAt: rows[i][10] ? String(rows[i][10]) : "",
        reversedBy: String(rows[i][11] || ""),
        createdAt: rows[i][12] ? String(rows[i][12]) : ""
      };
    }
  }
  return null;
}

function getMonthlyBankDepositSummary(dataJson) {
  try {
    var data = JSON.parse(dataJson || "{}");
    var month = String(data.month || "").padStart(2, "0");
    var year = Number(data.year) || (new Date().getFullYear() + 543);
    if (!month || month === "00") return JSON.stringify({success:false,error:"กรุณาเลือกเดือน"});

    var closed = _monthAlreadyClosed(month, year);
    var gregYear = year > 2400 ? year - 543 : year;
    var stuRows = getOrCreateSheet(SHEET_STUDENTS).getDataRange().getValues();
    var studentClass = {};
    for (var s = 1; s < stuRows.length; s++) {
      if (stuRows[s][0]) studentClass[String(stuRows[s][0])] = String(stuRows[s][2] || "");
    }

    var txRows = getOrCreateSheet(SHEET_TX).getDataRange().getValues();
    var classMap = {}, total = 0, count = 0;
    for (var i = 1; i < txRows.length; i++) {
      var r = txRows[i];
      if (!r[0] || String(r[2]) !== "DEPOSIT") continue;
      var d = _parseTxDate(r[3]);
      if (!d) continue;
      if (d.getFullYear() !== gregYear || d.getMonth() + 1 !== Number(month)) continue;
      var amt = Number(r[4]) || 0;
      var cls = studentClass[String(r[1])] || "ไม่ระบุชั้น";
      if (!classMap[cls]) classMap[cls] = {className:cls,totalAmount:0,totalTransactions:0};
      classMap[cls].totalAmount += amt;
      classMap[cls].totalTransactions++;
      total += amt;
      count++;
    }
    var summary = [];
    for (var c in classMap) summary.push(classMap[c]);
    summary.sort(function(a,b){ return CLASSES.indexOf(a.className) - CLASSES.indexOf(b.className); });

    return JSON.stringify({
      success:true,
      month:month,
      year:year,
      closed:!!closed,
      existing:closed,
      closedAt:new Date().toISOString(),
      closedBy:_getUserEmailSafe(),
      totalAmount:total,
      totalTransactions:count,
      classSummary:summary
    });
  } catch(e) { return JSON.stringify({success:false,error:e.message}); }
}

function saveMonthlyBankDeposit(dataJson) {
  try {
    var data = JSON.parse(dataJson || "{}");
    var month = String(data.month || "").padStart(2, "0");
    var year = Number(data.year) || (new Date().getFullYear() + 543);
    var note = String(data.note || "");
    if (!month || month === "00") return JSON.stringify({success:false,error:"กรุณาเลือกเดือน"});
    var closed = _monthAlreadyClosed(month, year);
    if (closed) return JSON.stringify({success:false,duplicate:true,error:"เดือนนี้ถูกปิดยอดแล้ว",existing:closed});

    var summaryRes = JSON.parse(getMonthlyBankDepositSummary(JSON.stringify({month:month,year:year})));
    if (!summaryRes.success) return JSON.stringify(summaryRes);
    var closedBy = String(data.closedBy || "").trim();
    if (!closedBy) return JSON.stringify({success:false,error:"กรุณากรอกชื่อผู้ดำเนินการ"});
    var sh = _getBankSheet();
    var now = new Date().toISOString();
    var id = "BANK" + new Date().getTime();
    sh.appendRow([
      id, month, year, now, closedBy, Number(summaryRes.totalAmount) || 0,
      Number(summaryRes.totalTransactions) || 0,
      JSON.stringify(summaryRes.classSummary || []), note, "ACTIVE", "", "", now
    ]);
    writeAudit("MONTHLY_BANK_DEPOSIT",
      "ปิดยอดนำฝากธนาคาร เดือน "+month+"/"+year+" จำนวน ฿"+Number(summaryRes.totalAmount||0).toLocaleString()+" รายการ "+summaryRes.totalTransactions,
      id);
    writeAuditLog("MONTHLY_BANK_DEPOSIT","MonthlyBankDeposit",id,"",
      {month:month,year:year,totalAmount:summaryRes.totalAmount,totalTransactions:summaryRes.totalTransactions,status:"ACTIVE"},
      closedBy,note);
    return JSON.stringify({
      success:true,
      depositId:id,
      month:month,
      year:year,
      closedAt:now,
      closedBy:closedBy,
      totalAmount:summaryRes.totalAmount,
      totalTransactions:summaryRes.totalTransactions,
      classSummary:summaryRes.classSummary,
      note:note,
      status:"ACTIVE",
      reversedAt:"",
      reversedBy:"",
      createdAt:now
    });
  } catch(e) { return JSON.stringify({success:false,error:e.message}); }
}

function getMonthlyBankDeposits() {
  try {
    var sh = _getBankSheet();
    var rows = sh.getDataRange().getValues();
    var data = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i]; if (!r[0]) continue;
      var summary = [];
      try { summary = JSON.parse(String(r[7] || "[]")); } catch(e) {}
      data.push({
        depositId:String(r[0]), month:String(r[1]), year:Number(r[2])||0,
        closedAt:r[3]?String(r[3]):"", closedBy:String(r[4]||""),
        totalAmount:Number(r[5])||0, totalTransactions:Number(r[6])||0,
        classSummary:summary, classSummaryJson:String(r[7]||"[]"),
        note:String(r[8]||""), status:String(r[9]||"ACTIVE"),
        reversedAt:r[10]?String(r[10]):"", reversedBy:String(r[11]||""),
        createdAt:r[12]?String(r[12]):""
      });
    }
    data.reverse();
    return JSON.stringify({success:true,data:data});
  } catch(e) { return JSON.stringify({success:false,error:e.message}); }
}

function reverseMonthlyBankDeposit(dataJson) {
  try {
    var data = JSON.parse(dataJson || "{}");
    var adminEmail = _requireAdmin(data.adminToken);
    var id = String(data.depositId || "");
    var sh = _getBankSheet();
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === id) {
        var oldStatus = String(rows[i][9] || "ACTIVE");
        if (oldStatus === "REVERSED") return JSON.stringify({success:false,error:"รายการนี้ถูกย้อนสถานะแล้ว"});
        var now = new Date().toISOString();
        sh.getRange(i+1,10).setValue("REVERSED");
        sh.getRange(i+1,11).setValue(now);
        sh.getRange(i+1,12).setValue(adminEmail);
        writeAudit("REVERSE_MONTHLY_BANK_DEPOSIT","ย้อนสถานะปิดยอด "+id,id);
        writeAuditLog("REVERSE_MONTHLY_BANK_DEPOSIT","MonthlyBankDeposit",id,
          {status:oldStatus},{status:"REVERSED",reversedAt:now,reversedBy:adminEmail},
          adminEmail,String(data.note||"ย้อนสถานะปิดยอด"));
        return JSON.stringify({success:true,depositId:id,status:"REVERSED",reversedAt:now,reversedBy:adminEmail});
      }
    }
    return JSON.stringify({success:false,error:"ไม่พบรายการปิดยอด"});
  } catch(e) { return JSON.stringify({success:false,error:e.message}); }
}

function getAuditLogs(dataJson) {
  try {
    var data = JSON.parse(dataJson || "{}");
    _requireAdmin(data.adminToken);
    var sh = getOrCreateSheet(SHEET_AUDIT_LOGS, _auditLogHeaders());
    var rows = sh.getDataRange().getValues();
    var dataRows = [];
    for (var i=1;i<rows.length;i++) {
      var r=rows[i]; if(!r[0]) continue;
      dataRows.push({logId:String(r[0]),actionType:String(r[1]||""),targetType:String(r[2]||""),targetId:String(r[3]||""),oldValue:String(r[4]||""),newValue:String(r[5]||""),editedBy:String(r[6]||""),editedAt:r[7]?String(r[7]):"",note:String(r[8]||"")});
    }
    dataRows.reverse();
    return JSON.stringify({success:true,data:dataRows});
  } catch(e) { return JSON.stringify({success:false,error:e.message}); }
}

function getRealBalanceSummary() {
  try {
    var closed = {};
    var bank = 0;
    var bankRows = _getBankSheet().getDataRange().getValues();
    for (var b=1;b<bankRows.length;b++) {
      var br=bankRows[b]; if(!br[0]) continue;
      var st=String(br[9]||"ACTIVE");
      if(st!=="ACTIVE") continue;
      var key=String(br[2])+"-"+String(br[1]).padStart(2,"0");
      closed[key]=true;
      bank += Number(br[5])||0;
    }
    var unclosed = 0, unclosedTx = 0;
    var txRows = getOrCreateSheet(SHEET_TX).getDataRange().getValues();
    for (var i=1;i<txRows.length;i++) {
      var r=txRows[i]; if(!r[0] || String(r[2])!=="DEPOSIT") continue;
      var d=_parseTxDate(r[3]); if(!d) continue;
      var y=d.getFullYear()+543, m=String(d.getMonth()+1).padStart(2,"0");
      if(!closed[y+"-"+m]) { unclosed += Number(r[4])||0; unclosedTx++; }
    }
    return JSON.stringify({success:true,unclosedAmount:unclosed,bankAmount:bank,totalAmount:unclosed+bank,unclosedTransactions:unclosedTx});
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
