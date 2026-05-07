// ===== UPDATED GOOGLE APPS SCRIPT =====
// Supports JSONP for iPhone/Safari compatibility
// Paste in Extensions → Apps Script → Deploy as Web App

const SHEET_ID = "18Wxm5i-sIfiQ7RSUs9Q7nyI2_OtBnTJuToUAz4j1LSI";

const HEADERS = [
  "तारीख","समय","विद्यार्थी नाम","पिता का नाम","कक्षा","विद्यालय","ब्लॉक",
  "अंग्रेज़ी L1","अंग्रेज़ी L2","अंग्रेज़ी श्रेणी",
  "हिंदी L1","हिंदी L2","हिंदी श्रेणी",
  "गणित L1","गणित L2","गणित श्रेणी","समग्र श्रेणी"
];

function saveToLocal(studentData) {
  let students = JSON.parse(localStorage.getItem("students") || "[]");

  students.push(studentData);

  localStorage.setItem("students", JSON.stringify(students));
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName("डेटा");
  if (!sheet) {
    sheet = ss.insertSheet("डेटा");
    const hr = sheet.getRange(1,1,1,HEADERS.length);
    sheet.appendRow(HEADERS);
    hr.setBackground("#1a5276").setFontColor("#ffffff").setFontWeight("bold").setFontSize(11);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(3,150); sheet.setColumnWidth(4,150); sheet.setColumnWidth(6,200);
  }
  return sheet;
}

function colorCatCell(sheet, row, col, cat) {
  const cell = sheet.getRange(row, col);
  if (cat==="proficient") { cell.setBackground("#d5f5e3"); cell.setFontColor("#1e8449"); }
  else if (cat==="developing") { cell.setBackground("#fef9e7"); cell.setFontColor("#9a7d0a"); }
  else if (cat==="needs") { cell.setBackground("#fadbd8"); cell.setFontColor("#922b21"); }
}

// Handle POST — save student data
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();
    const now = new Date();
    const r = data.results || {};
    const row = [
      Utilities.formatDate(now,"Asia/Kolkata","dd/MM/yyyy"),
      Utilities.formatDate(now,"Asia/Kolkata","HH:mm"),
      data.name, data.father, data.cls, data.school, data.block,
      r.english?.l1score??"", r.english?.l2score??"", r.english?.category??"",
      r.hindi?.l1score??"",   r.hindi?.l2score??"",   r.hindi?.category??"",
      r.math?.l1score??"",    r.math?.l2score??"",    r.math?.category??"",
      data.overallCategory??""
    ];
    sheet.appendRow(row);
    const lr = sheet.getLastRow();
    colorCatCell(sheet,lr,10,r.english?.category);
    colorCatCell(sheet,lr,13,r.hindi?.category);
    colorCatCell(sheet,lr,16,r.math?.category);
    colorCatCell(sheet,lr,17,data.overallCategory);
    return ContentService.createTextOutput(JSON.stringify({status:"success"})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status:"error",msg:err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

// Handle GET — fetch data OR save (JSONP workaround for Safari/iPhone)
function doGet(e) {
  try {
    const callback = e.parameter.callback;
    const action = e.parameter.action;

    // JSONP save (iPhone workaround — data sent as GET param)
    if (action === "save") {
      const data = JSON.parse(decodeURIComponent(e.parameter.data || "{}"));
      const sheet = getOrCreateSheet();
      const now = new Date();
      const r = data.results || {};
      const row = [
        Utilities.formatDate(now,"Asia/Kolkata","dd/MM/yyyy"),
        Utilities.formatDate(now,"Asia/Kolkata","HH:mm"),
        data.name, data.father, data.cls, data.school, data.block,
        r.english?.l1score??"", r.english?.l2score??"", r.english?.category??"",
        r.hindi?.l1score??"",   r.hindi?.l2score??"",   r.hindi?.category??"",
        r.math?.l1score??"",    r.math?.l2score??"",    r.math?.category??"",
        data.overallCategory??""
      ];
      sheet.appendRow(row);
      const lr = sheet.getLastRow();
      colorCatCell(sheet,lr,10,r.english?.category);
      colorCatCell(sheet,lr,13,r.hindi?.category);
      colorCatCell(sheet,lr,16,r.math?.category);
      colorCatCell(sheet,lr,17,data.overallCategory);
      const json = JSON.stringify({status:"success"});
      if (callback) return ContentService.createTextOutput(callback+"("+json+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
      return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
    }

    // Normal fetch
    const scope = e.parameter.scope || "district";
    const sheet = getOrCreateSheet();
    let students = [];
    if (sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      students = data.slice(1).map(row => {
        const o = {};
        headers.forEach((h,i) => o[h] = row[i]);
        return {
          date: o["\u0924\u093e\u0930\u0940\u0916"], name: o["\u0935\u093f\u0926\u094d\u092f\u093e\u0930\u094d\u0925\u0940 \u0928\u093e\u092e"], father: o["\u092a\u093f\u0924\u093e \u0915\u093e \u0928\u093e\u092e"],
          cls: o["\u0915\u0915\u094d\u0937\u093e"], school: o["\u0935\u093f\u0926\u094d\u092f\u093e\u0932\u092f"], block: o["\u092c\u094d\u0932\u0949\u0915"],
          results: {
            english:{l1score:o["\u0905\u0902\u0917\u094d\u0930\u0947\u091c\u093c\u0940 L1"],l2score:o["\u0905\u0902\u0917\u094d\u0930\u0947\u091c\u093c\u0940 L2"],category:o["\u0905\u0902\u0917\u094d\u0930\u0947\u091c\u093c\u0940 \u0936\u094d\u0930\u0947\u0923\u0940"]},
            hindi:  {l1score:o["\u0939\u093f\u0902\u0926\u0940 L1"],   l2score:o["\u0939\u093f\u0902\u0926\u0940 L2"],  category:o["\u0939\u093f\u0902\u0926\u0940 \u0936\u094d\u0930\u0947\u0923\u0940"]},
            math:   {l1score:o["\u0917\u0923\u093f\u0924 L1"],    l2score:o["\u0917\u0923\u093f\u0924 L2"],   category:o["\u0917\u0923\u093f\u0924 \u0936\u094d\u0930\u0947\u0923\u0940"]}
          },
          overallCategory: o["\u0938\u092e\u0917\u094d\u0930 \u0936\u094d\u0930\u0947\u0923\u0940"]
        };
      });
      if (scope !== "district") students = students.filter(s => s.school === scope);
    }
    const json = JSON.stringify({status:"success", students});
    if (callback) return ContentService.createTextOutput(callback+"("+json+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    const json = JSON.stringify({status:"error",msg:err.toString(),students:[]});
    const callback = e.parameter.callback;
    if (callback) return ContentService.createTextOutput(callback+"("+json+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }
}
