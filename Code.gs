// ===================================================
// Google Apps Script — พื้นที่นวัตกรรมการศึกษาจังหวัดเชียงใหม่ 2569
// ===================================================

var SHEET_ID   = '1i1u_ORW2WxcxOv7OIJ-BKpjXF5xKFJGM5UEhe_mwzuo';
var SHEET_NAME = 'ข้อมูลโครงการ';

// ===================================================
// GET — ทดสอบ + ดึงข้อมูลทั้งหมด (JSONP)
// ===================================================
function doGet(e) {
  var callback = (e && e.parameter) ? e.parameter.callback : null;
  var action   = (e && e.parameter) ? e.parameter.action   : null;

  var result;

  if (action === 'getAll') {
    // ดึงข้อมูลทั้งหมดจาก Sheet
    result = getAllRecords();
  } else {
    result = JSON.stringify({ status: 'ok', message: 'Web App พร้อมใช้งาน' });
  }

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + result + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(result)
    .setMimeType(ContentService.MimeType.JSON);
}

// ===================================================
// ดึงข้อมูลทุกแถวจาก Sheet → แปลงเป็น records array
// ===================================================
function getAllRecords() {
  try {
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) {
      return JSON.stringify({ status: 'success', data: [] });
    }

    var data = sheet.getDataRange().getValues();
    // แถวที่ 1 = header, แถว 2 เป็นต้นไป = ข้อมูล
    var records = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      // คอลัมน์ตาม header ที่กำหนดใน saveToSheet:
      // 0=ลำดับ,1=วันที่,2=ประเภท,3=ชื่อ,4=ผู้รับผิดชอบ,5=โทร,6=โครงการ
      // 7=งบรวม,8=แหล่งงบ,9=รายละเอียดงบ,10=ระยะเวลา,11=จำนวนยุทธ์
      // 12=ยุทธศาสตร์,13=กลุ่มเป้า,14=รายละเอียดกลุ่ม,15=จำนวนคน
      // 16=ผลลัพธ์,17=แผนขยาย,18=ข้อเสนอ

      // แปลง budgetBreakdown จาก string "แหล่ง: จำนวน | แหล่ง: จำนวน"
      var bdBreakdown = {};
      if (row[9]) {
        row[9].toString().split(' | ').forEach(function(part) {
          var m = part.match(/^(.+):\s*([\d,]+)\s*บาท$/);
          if (m) bdBreakdown[m[1].trim()] = parseInt(m[2].replace(/,/g,''));
        });
      }

      // แปลง targetBreakdown จาก string "กลุ่ม: จำนวน คน | ..."
      var tgtBreakdown = {};
      if (row[14]) {
        row[14].toString().split(' | ').forEach(function(part) {
          var m = part.match(/^(.+):\s*([\d,]+)\s*คน$/);
          if (m) tgtBreakdown[m[1].trim()] = parseInt(m[2].replace(/,/g,''));
        });
      }

      records.push({
        id:              i,
        ts:              row[1] ? row[1].toString() : '',
        orgType:         row[2] ? row[2].toString() : '',
        orgName:         row[3] ? row[3].toString() : '',
        contact:         row[4] ? row[4].toString() : '',
        phone:           row[5] ? row[5].toString() : '',
        project:         row[6] ? row[6].toString() : '',
        budget:          row[7] ? Number(row[7]) : 0,
        budgetSrcStr:    row[8] ? row[8].toString() : '',
        budgetBreakdown: bdBreakdown,
        period:          row[10] ? row[10].toString() : '',
        stratCount:      row[11] ? Number(row[11]) : 0,
        strategy:        row[12] ? row[12].toString() : '',
        targetStr:       row[13] ? row[13].toString() : '',
        targetBreakdown: tgtBreakdown,
        targetTotal:     row[15] ? Number(row[15]) : 0,
        outcome:         row[16] ? row[16].toString() : '',
        expand:          row[17] ? row[17].toString() : '',
        suggestion:      row[18] ? row[18].toString() : ''
      });
    }

    return JSON.stringify({ status: 'success', data: records });
  } catch(err) {
    Logger.log('getAllRecords error: ' + err.toString());
    return JSON.stringify({ status: 'error', message: err.toString(), data: [] });
  }
}

// ===================================================
// POST — รับข้อมูลจาก hidden form (field: payload)
// ===================================================
function doPost(e) {
  try {
    var raw = '';
    if (e && e.parameter && e.parameter.payload) {
      raw = e.parameter.payload;
    } else if (e && e.postData && e.postData.contents) {
      raw = e.postData.contents;
    }
    if (!raw) throw new Error('ไม่พบข้อมูล payload');

    var data = JSON.parse(raw);
    saveToSheet(data);

    return HtmlService.createHtmlOutput(
      '<html><body>' +
      '<p style="font-family:sans-serif;color:green;text-align:center;padding:20px">✅ บันทึกข้อมูลเรียบร้อย</p>' +
      '</body></html>'
    );
  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return HtmlService.createHtmlOutput(
      '<html><body><p style="color:red;font-family:sans-serif;padding:20px">Error: ' + err.toString() + '</p></body></html>'
    );
  }
}

// ===================================================
// บันทึกลง Google Sheets
// ===================================================
function saveToSheet(data) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    var headers = [
      'ลำดับ','วันที่-เวลาที่ส่ง','ประเภทหน่วยงาน','ชื่อหน่วยงาน/สถานศึกษา',
      'ชื่อผู้รับผิดชอบ','เบอร์โทรศัพท์','ชื่อโครงการ/กิจกรรม',
      'งบประมาณรวม (บาท)','แหล่งงบประมาณ','รายละเอียดงบแยกแหล่ง',
      'ระยะเวลาดำเนินการ','จำนวนยุทธศาสตร์','ยุทธศาสตร์/กลยุทธ์ที่เชื่อมโยง',
      'กลุ่มเป้าหมาย','รายละเอียดกลุ่มเป้าหมาย','จำนวนรวม (คน)',
      'ผลลัพธ์สำคัญ','แผนขยายผล','ข้อเสนอแนะ'
    ];
    sheet.appendRow(headers);
    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setBackground('#1a3557');
    hRange.setFontColor('#ffffff');
    hRange.setFontWeight('bold');
    hRange.setFontSize(11);
    hRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    var widths = [50,160,180,220,160,110,280,120,200,280,160,80,320,180,280,90,280,280,280];
    widths.forEach(function(w,i){ sheet.setColumnWidth(i+1, w); });
  }

  var nextRow = sheet.getLastRow() + 1;
  var rowNum  = nextRow - 1;
  var dateStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');

  var bdDetail = '';
  if (data.budgetBreakdown) {
    var bdParts = [];
    for (var src in data.budgetBreakdown) {
      bdParts.push(src + ': ' + Number(data.budgetBreakdown[src]).toLocaleString() + ' บาท');
    }
    bdDetail = bdParts.join(' | ');
  }

  var tgtDetail = '';
  if (data.targetBreakdown) {
    var tgtParts = [];
    for (var grp in data.targetBreakdown) {
      tgtParts.push(grp + ': ' + Number(data.targetBreakdown[grp]).toLocaleString() + ' คน');
    }
    tgtDetail = tgtParts.join(' | ');
  }

  var row = [
    rowNum, dateStr,
    data.orgType      || '',
    data.orgName      || '',
    data.contact      || '',
    data.phone        || '',
    data.project      || '',
    Number(data.budget)      || 0,
    data.budgetSrcStr || '',
    bdDetail,
    data.period       || '',
    Number(data.stratCount)  || 0,
    data.strategy     || '',
    data.targetStr    || '',
    tgtDetail,
    Number(data.targetTotal) || 0,
    data.outcome      || '',
    data.expand       || '',
    data.suggestion   || ''
  ];

  sheet.appendRow(row);

  var rowRange = sheet.getRange(nextRow, 1, 1, row.length);
  rowRange.setBackground(nextRow % 2 === 0 ? '#f0f6fb' : '#ffffff');
  rowRange.setWrap(true);
  rowRange.setVerticalAlignment('top');
  sheet.getRange(nextRow, 8).setNumberFormat('#,##0');
  sheet.getRange(nextRow, 16).setNumberFormat('#,##0');

  Logger.log('✅ บันทึกแถว ' + nextRow + ': ' + data.project);
}

// ===================================================
// ทดสอบ
// ===================================================
function testSave() {
  saveToSheet({
    orgType:'โรงเรียนนำร่อง', orgName:'โรงเรียนทดสอบ',
    contact:'นางสาวทดสอบ ระบบ', phone:'0812345678',
    project:'โครงการทดสอบการบันทึกข้อมูล', budget:50000,
    budgetSrcStr:'งบประมาณต้นสังกัด',
    budgetBreakdown:{'งบประมาณต้นสังกัด':50000},
    period:'ไตรมาส 1', stratCount:2,
    strategy:'1.1 พัฒนานวัตกรรมการศึกษา | 4.1 สนับสนุนเทคโนโลยี',
    targetStr:'ผู้เรียน, ครู',
    targetBreakdown:{'ผู้เรียน':100,'ครู':10},
    targetTotal:110,
    outcome:'ทดสอบผลลัพธ์', expand:'ทดสอบแผนขยาย', suggestion:'ทดสอบข้อเสนอแนะ'
  });
  Logger.log('testSave เสร็จสิ้น');
}

function testGetAll() {
  var result = getAllRecords();
  Logger.log(result);
}
