// ==========================================
// 顧客カルテ管理 WebApp - Code.gs
// Developed by fukkey
// シート列構成:
//   A(0):タイムスタンプ  B(1):名前  C(2):電話番号
//   D(3):施術日          E(4):メニュー  F(5):施術内容
//   G(6):金額            H(7):指名の有無  I(8):髪質  J(9):メモ  K(10):写真URL
//   L(11):カラー種別     M(12):薬剤_根本  N(13):薬剤_中間
//   O(14):薬剤_毛先      P(15):薬剤_その他  Q(16):オキシ濃度
//   R(17):補色情報       S(18):仕上がり評価  T(19):ビフォー写真URL
//   U(20):アフター写真URL  V(21):担当者
// ==========================================

var SHEET_NAME   = "カルテデータ";
var CONFIG_SHEET = "設定";
var PHOTO_FOLDER = "カルテ写真";
var HEADERS = [
  "タイムスタンプ","名前","電話番号","施術日","メニュー","施術内容",
  "金額","指名の有無","髪質","メモ","写真URL",
  "カラー種別","薬剤_根本","薬剤_中間","薬剤_毛先","薬剤_その他",
  "オキシ濃度","補色情報","仕上がり評価","ビフォー写真URL","アフター写真URL","担当者"
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("顧客カルテ")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0");
}

function initSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else {
    // 既存シートにヘッダーが不足していれば補完
    var existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (existing.length < HEADERS.length) {
      for (var i = existing.length; i < HEADERS.length; i++) {
        sheet.getRange(1, i + 1).setValue(HEADERS[i]);
      }
    }
  }
  return sheet;
}

function searchCustomers(name, menu) {
  var sheet = initSheet();
  var data  = sheet.getDataRange().getValues();
  var results = [];
  name = (name || "").trim();
  menu = (menu || "").trim();
  var menuWords = menu ? menu.split(/\s+/).filter(function(w){ return w; }) : [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var rName    = String(r[1]  || "").trim();
    var rMenu    = String(r[4]  || "").trim();
    var rContent = String(r[5]  || "").trim();
    var rHair    = String(r[8]  || "").trim();
    var rColorType = String(r[11] || "").trim();
    var rDrug1   = String(r[12] || "").trim();
    var rDrug2   = String(r[13] || "").trim();
    var rDrug3   = String(r[14] || "").trim();
    var rDrug4   = String(r[15] || "").trim();
    var rEval    = String(r[18] || "").trim();
    var nameOk = !name || rName.includes(name);
    var menuOk = menuWords.length === 0 || menuWords.some(function(w){
      return rMenu.includes(w) || rContent.includes(w) || rHair.includes(w)
          || rColorType.includes(w) || rDrug1.includes(w) || rDrug2.includes(w)
          || rDrug3.includes(w) || rDrug4.includes(w) || rEval.includes(w);
    });
    if (nameOk && menuOk) results.push(toObj(r, i + 1));
  }
  results.sort(function(a, b) { return (b.施術日 || "").localeCompare(a.施術日 || ""); });
  return results;
}

function getCustomerList() {
  var sheet = initSheet();
  var data  = sheet.getDataRange().getValues();
  var map   = {};
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][1] || "").trim();
    var date = String(data[i][3] || "");
    if (!name) continue;
    if (!map[name]) map[name] = { count: 0, lastDate: "" };
    map[name].count++;
    if (date > map[name].lastDate) map[name].lastDate = date;
  }
  var list = [];
  for (var n in map) list.push({ name: n, count: map[n].count, lastDate: fmtDate(map[n].lastDate) });
  list.sort(function(a, b) { return a.name.localeCompare(b.name, "ja"); });
  return list;
}

function getHistory(name) {
  var sheet = initSheet();
  var data  = sheet.getDataRange().getValues();
  var list  = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1] || "").trim() === name) list.push(toObj(data[i], i + 1));
  }
  list.sort(function(a, b) { return (b.施術日 || "").localeCompare(a.施術日 || ""); });
  return list;
}

function addRecord(rec) {
  var sheet = initSheet();
  sheet.appendRow([
    new Date(),
    rec.名前        || "",
    rec.電話番号     || "",
    rec.施術日      || fmtDate(new Date()),
    rec.メニュー     || "",
    rec.施術内容     || "",
    rec.金額        || "",
    rec.指名の有無   || "",
    rec.髪質        || "",
    rec.メモ        || "",
    rec.写真URL     || "",
    rec.カラー種別   || "",
    rec.薬剤_根本   || "",
    rec.薬剤_中間   || "",
    rec.薬剤_毛先   || "",
    rec.薬剤_その他  || "",
    rec.オキシ濃度   || "",
    rec.補色情報     || "",
    rec.仕上がり評価  || "",
    rec.ビフォー写真URL || "",
    rec.アフター写真URL || "",
    rec.担当者       || ""
  ]);
  return "ok";
}

function updateRecord(rowNum, rec) {
  var sheet = initSheet();
  sheet.getRange(rowNum, 2, 1, 21).setValues([[
    rec.名前        || "",
    rec.電話番号     || "",
    rec.施術日      || "",
    rec.メニュー     || "",
    rec.施術内容     || "",
    rec.金額        || "",
    rec.指名の有無   || "",
    rec.髪質        || "",
    rec.メモ        || "",
    rec.写真URL     || "",
    rec.カラー種別   || "",
    rec.薬剤_根本   || "",
    rec.薬剤_中間   || "",
    rec.薬剤_毛先   || "",
    rec.薬剤_その他  || "",
    rec.オキシ濃度   || "",
    rec.補色情報     || "",
    rec.仕上がり評価  || "",
    rec.ビフォー写真URL || "",
    rec.アフター写真URL || "",
    rec.担当者       || ""
  ]]);
  return "ok";
}

function deleteRecord(rowNum) {
  var sheet = initSheet();
  sheet.deleteRow(rowNum);
  return "ok";
}

function savePhoto(base64Data, fileName) {
  return base64Data;
}

function getPhotoBase64(base64Data) {
  return base64Data ? "data:image/jpeg;base64," + base64Data : "";
}

function getAnalytics() {
  var sheet = initSheet();
  var data  = sheet.getDataRange().getValues();

  var customers = {};
  for (var i = 1; i < data.length; i++) {
    var r    = data[i];
    var name = String(r[1] || "").trim();
    if (!name) continue;
    if (!customers[name]) customers[name] = { visits: [], sales: [] };
    var date   = r[3] ? new Date(r[3]) : null;
    var amount = parseInt(r[6]) || 0;
    if (date) customers[name].visits.push(date);
    customers[name].sales.push(amount);
  }

  var royalList    = [];
  var avgPriceList = [];
  var cycleList    = [];

  for (var n in customers) {
    var c      = customers[n];
    var total  = c.sales.reduce(function(a,b){ return a+b; }, 0);
    var avg    = c.sales.length ? Math.round(total / c.sales.length) : 0;
    var visits = c.visits.length;
    royalList.push({ name: n, visits: visits, avg: avg, score: visits * avg, total: total });
    avgPriceList.push({ name: n, avg: avg, visits: visits });
    if (c.visits.length >= 2) {
      c.visits.sort(function(a,b){ return a-b; });
      var gaps = [];
      for (var j = 1; j < c.visits.length; j++) {
        gaps.push((c.visits[j] - c.visits[j-1]) / (1000*60*60*24));
      }
      var avgCycle = Math.round(gaps.reduce(function(a,b){ return a+b; },0) / gaps.length);
      cycleList.push({ name: n, cycle: avgCycle, visits: visits });
    }
  }

  royalList.sort(function(a,b){ return b.visits - a.visits; });
  avgPriceList.sort(function(a,b){ return b.avg - a.avg; });
  cycleList.sort(function(a,b){ return a.cycle - b.cycle; });

  return {
    royal:    royalList.slice(0, 20),
    avgPrice: avgPriceList.slice(0, 10),
    cycle:    cycleList.slice(0, 10)
  };
}

function getConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s  = ss.getSheetByName(CONFIG_SHEET);
  var def = defaultConfig();
  if (!s) return def;
  var v = s.getRange("A1").getValue();
  if (!v) return def;
  var saved = JSON.parse(v);
  def.menuOptions.forEach(function(m) {
    if (saved.menuOptions.indexOf(m) === -1) saved.menuOptions.push(m);
  });
  if (!saved.staffList) saved.staffList = def.staffList;
  return saved;
}

function saveConfig(cfg) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s  = ss.getSheetByName(CONFIG_SHEET) || ss.insertSheet(CONFIG_SHEET);
  s.getRange("A1").setValue(JSON.stringify(cfg));
  return "ok";
}

function defaultConfig() {
  return {
    salonName:   "カルテ管理",
    menuOptions: ["カット","カラー","ストレート","パーマ","トリートメント","ブリーチ","HM","ヘナ","スパ"],
    staffList:   []
  };
}

function toObj(r, rowNum) {
  return {
    _row:         rowNum,
    名前:         r[1],  電話番号:     r[2],  施術日:       fmtDate(r[3]),
    メニュー:     r[4],  施術内容:     r[5],  金額:         r[6],
    指名の有無:   r[7],  髪質:         r[8],  メモ:         r[9],
    写真URL:      r[10],
    カラー種別:   r[11], 薬剤_根本:   r[12], 薬剤_中間:   r[13],
    薬剤_毛先:   r[14], 薬剤_その他: r[15], オキシ濃度:   r[16],
    補色情報:     r[17], 仕上がり評価: r[18],
    ビフォー写真URL: r[19], アフター写真URL: r[20],
    担当者:       r[21]
  };
}

function fmtDate(d) {
  if (!d) return "";
  try { return Utilities.formatDate(new Date(d), "Asia/Tokyo", "yyyy/MM/dd"); }
  catch(e) { return String(d); }
}
