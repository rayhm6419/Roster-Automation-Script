/********* 配置区：只改这里 *********/
const CFG = {
  // 来源与目标
  SOURCE_SHEET: 'script',
  TARGET_SHEET: '司机表',

  // 表头所在行（你的 script 页日期与字段都在第 1 行）
  DATE_HEADER_ROW: 1,
  FIELD_HEADER_ROW: 1,

  // 字段表头（不区分大小写；给出候选以兼容中英文）
  COL_HEADER_NAME:  'Name SEA',
  COL_HEADER_CAR:   'Car Type',
  COL_HEADER_ROUTE_CANDIDATES:     ['Route', '路线'],
  COL_HEADER_AREA_CANDIDATES:      ['Route Area', '路线区域'],
  COL_HEADER_ROUTE_NO_CANDIDATES:  ['Route Number', 'Route No', '路线编号', '线路编号'],

  // 写入到《司机表》的位置
  DEST_START_ROW: 3,
  DEST_COL_NAME:      'M',   // 姓名
  DEST_COL_ROUTE_NO:  'N',   // 线路编号
  DEST_COL_CAR:       'O',   // 车型
  DEST_COL_ROUTE:     'P',   // 路线
  DEST_COL_AREA:      'R',   // 路线区域

  // 哪些值算不上班（空白=上班）
  OFF_WORDS: ['off', 'office', 'vacation', '请假', '休'],
  OFF_REGEXPS: [/^off\b/i],  // 如 "off?"、"off day" 也算不上班

  // 仅高亮表头单元格的颜色
  HIGHLIGHT_COLOR: '#FFF3CD',

  // 追加 DSP 行
  DSP_APPEND: {
    count: 13,               // 追加数量
    prefix: 'DSP Driver ',   // 名字前缀
    // 需要默认值可在这里填；留空则不写
    defaults: {
      routeNo: '',           // 比如 '-' 或 'DSP'
      car:     '',           // 比如 'Old MiniVan'
      route:   '',           // 比如 'DT Seattle'
      area:    ''            // 比如 'Bellevue线'
    },
    routeNoComments: ['101 Olympia', '103 Tacoma i5 left', '111', 'Mercer Island + Bellevue DT'],

  // ← 可选：这些注释的文字颜色（不需要就删掉此行）
    routeNoCommentColor: '#d93025'
  }
};
/********* 配置区结束 *********/


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Roster Automation')
    .addItem('从 script 抓取 → 生成今天', 'buildRosterFromScriptToday')
    .addItem('从 script 抓取 → 生成明天', 'buildRosterFromScriptTomorrow')
    .addToUi();
}

function buildRosterFromScriptToday(){ buildRosterFromScript_(0); }
function buildRosterFromScriptTomorrow(){ buildRosterFromScript_(1); }

/** 从 script 抓取（0=今天，1=明天），写入到《司机表》 */
function buildRosterFromScript_(dayOffset) {
  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(CFG.SOURCE_SHEET);
  const tgt = ss.getSheetByName(CFG.TARGET_SHEET);
  if (!src) throw new Error('找不到来源表：' + CFG.SOURCE_SHEET);
  if (!tgt) throw new Error('找不到目标表：' + CFG.TARGET_SHEET);

  const theDay = stripTime(addDays(new Date(), dayOffset));
  const lastCol = src.getLastColumn();
  const lastRow = src.getLastRow();
  if (lastCol < 1 || lastRow < CFG.FIELD_HEADER_ROW + 1) {
    SpreadsheetApp.getActiveSpreadsheet().toast('来源表为空', 'Roster Automation', 6);
    return;
  }

  // 表头
  const dateHeaders  = src.getRange(CFG.DATE_HEADER_ROW,  1, 1, lastCol).getValues()[0];
  const fieldHeaders = src.getRange(CFG.FIELD_HEADER_ROW, 1, 1, lastCol).getValues()[0];

  // 找当天列
  const cDay = findDateCol_(dateHeaders, theDay);
  if (!cDay) throw new Error('在 "'+CFG.SOURCE_SHEET+'" 第 '+CFG.DATE_HEADER_ROW+' 行找不到日期列：'+theDay);

  // 找字段列
  const cName    = findColByHeader_(fieldHeaders, CFG.COL_HEADER_NAME);
  const cCar     = findColByHeader_(fieldHeaders, CFG.COL_HEADER_CAR);
  const cRoute   = findFirstExistingHeaderCol_(fieldHeaders, CFG.COL_HEADER_ROUTE_CANDIDATES);
  const cArea    = findFirstExistingHeaderCol_(fieldHeaders, CFG.COL_HEADER_AREA_CANDIDATES);
  const cRouteNo = findFirstExistingHeaderCol_(fieldHeaders, CFG.COL_HEADER_ROUTE_NO_CANDIDATES);
  if (!cName) throw new Error('找不到姓名列（字段表头第 '+CFG.FIELD_HEADER_ROW+' 行）');

  // 数据区（表头下一行）
  const startRow = Math.max(CFG.FIELD_HEADER_ROW, CFG.DATE_HEADER_ROW) + 1;
  const nRows    = lastRow - startRow + 1;
  if (nRows <= 0) return;

  const names    = src.getRange(startRow, cName,    nRows, 1).getValues().flat();
  const marks    = src.getRange(startRow, cDay,     nRows, 1).getValues().flat();
  const cars     = cCar     ? src.getRange(startRow, cCar,     nRows, 1).getValues().flat() : [];
  const routes   = cRoute   ? src.getRange(startRow, cRoute,   nRows, 1).getValues().flat() : [];
  const areas    = cArea    ? src.getRange(startRow, cArea,    nRows, 1).getValues().flat() : [];
  const routeNos = cRouteNo ? src.getRange(startRow, cRouteNo, nRows, 1).getValues().flat() : [];

  const offSet = new Set(CFG.OFF_WORDS.map(function(s){ return String(s).toLowerCase().trim(); }));
  var out = [];

  for (var i = 0; i < nRows; i++) {
    var name = String(names[i] || '').trim();
    if (!name) continue;

    var raw  = String(marks[i] || '').trim();
    var mark = raw.toLowerCase();

    var isOff = false;
    if (raw === '') isOff = false;                         // 空白=上班
    else if (offSet.has(mark)) isOff = true;               // 明确词
    else if (CFG.OFF_REGEXPS.some(function(re){ return re.test(raw); })) isOff = true; // 模糊匹配

    if (!isOff) {
    var routeNoVal = cRouteNo ? String(routeNos[i] || '') : '';

    // ★ 特殊规则：Dalin Sun 在 周一/周三 强制写 '104 Belingham'
    var dow = theDay.getDay(); // 0=Sun,1=Mon,2=Tue,3=Wed,...
    if (name.toLowerCase() === 'dalin sun' && (dow === 1 || dow === 3)) {
      routeNoVal = '104 Belingham';
    }
    out.push({
      name:    name,
      car:     cCar     ? String(cars[i]     || '') : '',
      route:   cRoute   ? String(routes[i]   || '') : '',
      area:    cArea    ? String(areas[i]    || '') : '',
      routeNo: routeNoVal
    });
  }
}

  // 写入目标表（先清空后写入）
  writeColumn_(tgt, CFG.DEST_COL_NAME,      CFG.DEST_START_ROW, out.map(function(r){ return [r.name]; }));
  if (CFG.DEST_COL_ROUTE_NO) writeColumn_(tgt, CFG.DEST_COL_ROUTE_NO, CFG.DEST_START_ROW, out.map(function(r){ return [r.routeNo]; }));
  if (CFG.DEST_COL_CAR)      writeColumn_(tgt, CFG.DEST_COL_CAR,      CFG.DEST_START_ROW, out.map(function(r){ return [r.car]; }));
  if (CFG.DEST_COL_ROUTE)    writeColumn_(tgt, CFG.DEST_COL_ROUTE,    CFG.DEST_START_ROW, out.map(function(r){ return [r.route]; }));
  if (CFG.DEST_COL_AREA)     writeColumn_(tgt, CFG.DEST_COL_AREA,     CFG.DEST_START_ROW, out.map(function(r){ return [r.area]; }));

  // 在最后一行后追加 DSP Driver
  appendDSPRows_(tgt, CFG.DEST_START_ROW + out.length);

  // 仅高亮来源表表头当天单元格（不动原有颜色）
  var props = PropertiesService.getDocumentProperties();
  var key   = 'last_highlight_col_' + CFG.SOURCE_SHEET;
  var prev  = parseInt(props.getProperty(key) || '0', 10);
  if (prev && prev !== cDay) src.getRange(CFG.DATE_HEADER_ROW, prev).setBackground(null);
  src.getRange(CFG.DATE_HEADER_ROW, cDay).setBackground(CFG.HIGHLIGHT_COLOR);
  props.setProperty(key, String(cDay));

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '已从 "'+CFG.SOURCE_SHEET+'" 抓取 ' + (dayOffset===0?'今天':'明天') + '：' + out.length + ' 人',
    'Roster Automation', 6
  );
}


/********* 工具函数 *********/
function findColByHeader_(headers, label) {
  if (!label) return null;
  var t = String(label).toLowerCase().trim();
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]).toLowerCase().trim() === t) return c + 1;
  }
  return null;
}
function findFirstExistingHeaderCol_(headers, candidates) {
  if (!candidates) return null;
  for (var i = 0; i < candidates.length; i++) {
    var col = findColByHeader_(headers, candidates[i]);
    if (col) return col;
  }
  return null;
}
function findDateCol_(headers, targetDate) {
  for (var c = 0; c < headers.length; c++) {
    var d = toDate_(headers[c]);
    if (d && sameDay_(d, targetDate)) return c + 1;
  }
  return null;
}
function toDate_(v) {
  if (v instanceof Date && !isNaN(v)) return stripTime(v);
  var s = String(v || '').replace(/,/g,' ').trim();
  if (!s) return null;
  var d = new Date(s);
  return isNaN(d) ? null : stripTime(d);
}
function stripTime(d){ return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function addDays(d,n){ var x=new Date(d); x.setDate(x.getDate()+n); return x; }
function sameDay_(a,b){ return a.getFullYear()===b.getFullYear() && a.getMonth()===b.getMonth() && a.getDate()===b.getDate(); }

function writeColumn_(sheet, colA1, startRow, values2D){
  var col = sheet.getRange(colA1 + startRow).getColumn();
  var clearRows = sheet.getMaxRows() - startRow + 1;
  if (clearRows > 0) sheet.getRange(startRow, col, clearRows, 1).clearContent();
  if (values2D.length) sheet.getRange(startRow, col, values2D.length, 1).setValues(values2D);
}

// 追加 DSP 行（写名字；如配置了默认值与前 N 项覆盖值则一并写入）
function appendDSPRows_(tgt, startRowForDSP){
  var cfg = CFG.DSP_APPEND || {};
  var count  = cfg.count || 0;
  if (count <= 0) return;

  var nameCol = tgt.getRange(CFG.DEST_COL_NAME + CFG.DEST_START_ROW).getColumn();

  // 1) 写名字
  var names = [];
  for (var i = 1; i <= count; i++) {
    names.push([ (cfg.prefix || 'DSP Driver ') + i ]);
  }
  tgt.getRange(startRowForDSP, nameCol, names.length, 1).setValues(names);

  // 2) 写默认值（整列）
  var d = cfg.defaults || {};
  var col;

  // Route Number：先准备默认值，再用前 N 项覆盖
  if (CFG.DEST_COL_ROUTE_NO) {
    col = tgt.getRange(CFG.DEST_COL_ROUTE_NO + startRowForDSP).getColumn();

    // 先填充默认值（'' 或配置的默认值）
    var routeNoVals = names.map(function(){ return [ (d.routeNo !== undefined ? d.routeNo : '') ]; });

    // 再用 routeNoComments 覆盖前 N 行
    var comments = cfg.routeNoComments || [];
    for (var r = 0; r < Math.min(comments.length, names.length); r++) {
      routeNoVals[r][0] = String(comments[r]);
    }

    // 一次性写入
    tgt.getRange(startRowForDSP, col, routeNoVals.length, 1).setValues(routeNoVals);

    // 可选：给这些注释着色
    if (cfg.routeNoCommentColor && comments.length) {
      var color = cfg.routeNoCommentColor;
      for (var k = 0; k < Math.min(comments.length, names.length); k++) {
        tgt.getRange(startRowForDSP + k, col).setFontColor(color);
      }
    }
  }

  if (CFG.DEST_COL_CAR && d.car !== undefined) {
    col = tgt.getRange(CFG.DEST_COL_CAR + startRowForDSP).getColumn();
    tgt.getRange(startRowForDSP, col, names.length, 1).setValue(d.car);
  }
  if (CFG.DEST_COL_ROUTE && d.route !== undefined) {
    col = tgt.getRange(CFG.DEST_COL_ROUTE + startRowForDSP).getColumn();
    tgt.getRange(startRowForDSP, col, names.length, 1).setValue(d.route);
  }
  if (CFG.DEST_COL_AREA && d.area !== undefined) {
    col = tgt.getRange(CFG.DEST_COL_AREA + startRowForDSP).getColumn();
    tgt.getRange(startRowForDSP, col, names.length, 1).setValue(d.area);
  }
}

