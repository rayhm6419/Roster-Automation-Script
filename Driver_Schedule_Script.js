/********* 配置区：按需修改 *********/
const CFG = {
  // 来源与目标
  SOURCE_SHEET: 'script',
  TARGET_SHEET: '司机表',

  // 表头所在行（你的 script 页日期与字段都在第 1 行）
  DATE_HEADER_ROW: 1,
  FIELD_HEADER_ROW: 1,

  // 字段表头（大小写不敏感；提供候选以兼容中英文）
  COL_HEADER_NAME:  'Name SEA',
  COL_HEADER_CAR:   'Car Type',
  COL_HEADER_ROUTE_CANDIDATES:    ['Route', '路线'],
  COL_HEADER_AREA_CANDIDATES:     ['Route Area', '路线区域'],
  COL_HEADER_ROUTE_NO_CANDIDATES: ['Route Number', 'Route No', '路线编号', '线路编号'],

  // 写入到《司机表》的位置（列字母 + 起始行）
  DEST_COL_NAME:     'M',
  DEST_COL_ROUTE_NO: 'N',
  DEST_COL_CAR:      'O',
  DEST_COL_ROUTE:    'P',
  DEST_COL_AREA:     'R',
  DEST_START_ROW: 3,

  // 不上班关键字（空白=上班）
  OFF_WORDS: ['off', 'office', 'vacation', '请假', '休'],
  OFF_REGEXPS: [/^off\b/i],  // 如 "off?"、"off day"

  // 仅高亮来源页“当天表头”单元格的颜色（不改整列）
  HIGHLIGHT_COLOR: '#FFF3CD',

  // 追加 DSP 信息
  DSP_APPEND: {
    count: 13,
    prefix: 'DSP Driver ',
    defaults: { routeNo: '', car: '', route: '', area: '' },
    // 前 N 个 DSP 的 Route Number 覆盖（按顺序对应 DSP1..）
    routeNoComments: ['101 Olympia', '103 Tacoma i5 left', '111', 'Mercer Island + Bellevue DT'],
    routeNoCommentColor: '#d93025'   // 想用表格默认颜色就删掉这一行
  }
};
/********* 配置区结束 *********/


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Roster Automation')
    .addItem('从 script 抓取 → 生成今天', 'buildRosterFromScriptToday')
    .addItem('从 script 抓取 → 生成明天', 'buildRosterFromScriptTomorrow')
    .addSeparator()
    .addItem('forecast：选择日期生成司机表', 'forecast') // 日历弹窗
    .addToUi();
}

/** 今天/明天（便捷） */
function buildRosterFromScriptToday(){ buildRosterFromScriptByOffset_(0); }
function buildRosterFromScriptTomorrow(){ buildRosterFromScriptByOffset_(1); }

/** —— 入口：forecast（日历弹窗） —— */
function forecast() {
  const b = getScriptHeaderBounds_(); // 计算 script 表头的最小/最大日期与默认建议日期
  const t = HtmlService.createTemplateFromFile('forecast_dialog');
  t.minISO     = b.minISO;
  t.maxISO     = b.maxISO;
  t.defaultISO = b.suggestISO; // 默认给“明天”（若不在范围内则用最小值）
  const html = t.evaluate().setWidth(360).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, '选择日期生成司机表');
}

/** 被 HTML 调用：按 ISO 字符串(yyyy-MM-dd)生成司机表 */
function doForecastForDateString(iso) {
  if (!iso) throw new Error('未选择日期');
  const d = new Date(iso + 'T00:00:00'); // 用表格时区即可
  if (isNaN(d)) throw new Error('无法解析日期：' + iso);
  buildRosterFromScriptOnDate_(stripTime(d));
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return Utilities.formatDate(stripTime(d), tz, 'yyyy-MM-dd');
}

/** 计算 script 表头日期范围（min/max/suggest） */
function getScriptHeaderBounds_() {
  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(CFG.SOURCE_SHEET);
  if (!src) throw new Error('找不到来源表：' + CFG.SOURCE_SHEET);

  const lastCol = src.getLastColumn();
  const headers = src.getRange(CFG.DATE_HEADER_ROW, 1, 1, lastCol).getValues()[0];

  const dates = headers
    .map(h => toDate_(h))
    .filter(d => d && !isNaN(d.getTime()))
    .sort((a,b) => a - b);

  if (!dates.length) throw new Error('script 表头没有可解析的日期');

  const tz = ss.getSpreadsheetTimeZone();
  const toISO = d => Utilities.formatDate(d, tz, 'yyyy-MM-dd');

  const min = dates[0];
  const max = dates[dates.length - 1];

  // 建议默认值：明天（若不在范围内则回退到 min）
  const tmw = stripTime(addDays(new Date(), 1));
  const suggest = (tmw >= min && tmw <= max) ? tmw : min;

  return { minISO: toISO(min), maxISO: toISO(max), suggestISO: toISO(suggest) };
}


/** —— 内部：按“天数偏移”生成（今天/明天） —— */
function buildRosterFromScriptByOffset_(dayOffset) {
  const d = stripTime(addDays(new Date(), dayOffset));
  buildRosterFromScriptOnDate_(d);
}

/** —— 核心：按“绝对日期”生成 —— */
function buildRosterFromScriptOnDate_(theDay) {
  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(CFG.SOURCE_SHEET);
  const tgt = ss.getSheetByName(CFG.TARGET_SHEET);
  if (!src) throw new Error('找不到来源表：' + CFG.SOURCE_SHEET);
  if (!tgt) throw new Error('找不到目标表：' + CFG.TARGET_SHEET);

  const lastCol = src.getLastColumn();
  const lastRow = src.getLastRow();
  if (lastCol < 1 || lastRow < CFG.FIELD_HEADER_ROW + 1) {
    SpreadsheetApp.getActiveSpreadsheet().toast('来源表为空', 'Roster Automation', 6);
    return;
  }

  // 表头
  const dateHeaders  = src.getRange(CFG.DATE_HEADER_ROW,  1, 1, lastCol).getValues()[0];
  const fieldHeaders = src.getRange(CFG.FIELD_HEADER_ROW, 1, 1, lastCol).getValues()[0];

  // 找目标日期所在列
  const cDay = findDateCol_(dateHeaders, theDay);
  if (!cDay) throw new Error('在 "'+CFG.SOURCE_SHEET+'" 第 '+CFG.DATE_HEADER_ROW+' 行找不到该日期列：'+theDay);

  // 找字段列
  const cName    = findColByHeader_(fieldHeaders, CFG.COL_HEADER_NAME);
  const cCar     = findColByHeader_(fieldHeaders, CFG.COL_HEADER_CAR);
  const cRoute   = findFirstExistingHeaderCol_(fieldHeaders, CFG.COL_HEADER_ROUTE_CANDIDATES);
  const cArea    = findFirstExistingHeaderCol_(fieldHeaders, CFG.COL_HEADER_AREA_CANDIDATES);
  const cRouteNo = findFirstExistingHeaderCol_(fieldHeaders, CFG.COL_HEADER_ROUTE_NO_CANDIDATES);
  if (!cName) throw new Error('找不到姓名列（字段表头第 '+CFG.FIELD_HEADER_ROW+' 行）');

  // 数据区
  const startRow = Math.max(CFG.FIELD_HEADER_ROW, CFG.DATE_HEADER_ROW) + 1;
  const nRows    = lastRow - startRow + 1;
  if (nRows <= 0) return;

  const names    = src.getRange(startRow, cName,    nRows, 1).getValues().flat();
  const marks    = src.getRange(startRow, cDay,     nRows, 1).getValues().flat();
  const cars     = cCar     ? src.getRange(startRow, cCar,     nRows, 1).getValues().flat() : [];
  const routes   = cRoute   ? src.getRange(startRow, cRoute,   nRows, 1).getValues().flat() : [];
  const areas    = cArea    ? src.getRange(startRow, cArea,    nRows, 1).getValues().flat() : [];
  const routeNos = cRouteNo ? src.getRange(startRow, cRouteNo, nRows, 1).getValues().flat() : [];

  const offSet = new Set(CFG.OFF_WORDS.map(s => String(s).toLowerCase().trim()));
  const out = [];

  for (let i = 0; i < nRows; i++) {
    const name = String(names[i] || '').trim();
    if (!name) continue;

    const raw  = String(marks[i] || '').trim();
    const mark = raw.toLowerCase();

    let isOff = false;
    if (raw === '') isOff = false;                           // 空白=上班
    else if (offSet.has(mark)) isOff = true;                 // 明确词
    else if (CFG.OFF_REGEXPS.some(re => re.test(raw))) isOff = true; // 模糊匹配

    if (!isOff) {
      // 先取原始 Route No
      let routeNoVal = cRouteNo ? String(routeNos[i] || '') : '';

      // ★ 规则：Dalin Sun 在 周一/周三 强制 '104 Belingham'
      const dow = theDay.getDay(); // 0=Sun,1=Mon,2=Tue,3=Wed,...
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

  // 写入目标表（只清内容，不动格式）
  writeColumn_(tgt, CFG.DEST_COL_NAME,      CFG.DEST_START_ROW, out.map(r => [r.name]));
  if (CFG.DEST_COL_ROUTE_NO) writeColumn_(tgt, CFG.DEST_COL_ROUTE_NO, CFG.DEST_START_ROW, out.map(r => [r.routeNo]));
  if (CFG.DEST_COL_CAR)      writeColumn_(tgt, CFG.DEST_COL_CAR,      CFG.DEST_START_ROW, out.map(r => [r.car]));
  if (CFG.DEST_COL_ROUTE)    writeColumn_(tgt, CFG.DEST_COL_ROUTE,    CFG.DEST_START_ROW, out.map(r => [r.route]));
  if (CFG.DEST_COL_AREA)     writeColumn_(tgt, CFG.DEST_COL_AREA,     CFG.DEST_START_ROW, out.map(r => [r.area]));

  // 末尾追加 DSP
  appendDSPRows_(tgt, CFG.DEST_START_ROW + out.length);

  // 仅高亮来源表“表头当天单元格”
  const props = PropertiesService.getDocumentProperties();
  const key   = 'last_highlight_col_' + CFG.SOURCE_SHEET;
  const prev  = parseInt(props.getProperty(key) || '0', 10);
  if (prev && prev !== cDay) src.getRange(CFG.DATE_HEADER_ROW, prev).setBackground(null);
  src.getRange(CFG.DATE_HEADER_ROW, cDay).setBackground(CFG.HIGHLIGHT_COLOR);
  props.setProperty(key, String(cDay));

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '已从 "'+CFG.SOURCE_SHEET+'" 抓取：' +
    Utilities.formatDate(theDay, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') +
    '，上班 ' + out.length + ' 人',
    'Roster Automation', 6
  );
}


/********* 工具函数 *********/
function findColByHeader_(headers, label) {
  if (!label) return null;
  const t = String(label).toLowerCase().trim();
  for (let c = 0; c < headers.length; c++) {
    if (String(headers[c]).toLowerCase().trim() === t) return c + 1;
  }
  return null;
}
function findFirstExistingHeaderCol_(headers, candidates) {
  if (!candidates) return null;
  for (const lab of candidates) {
    const col = findColByHeader_(headers, lab);
    if (col) return col;
  }
  return null;
}
function findDateCol_(headers, targetDate) {
  for (let c = 0; c < headers.length; c++) {
    const d = toDate_(headers[c]);
    if (d && sameDay_(d, targetDate)) return c + 1;
  }
  return null;
}
function toDate_(v) {
  if (v instanceof Date && !isNaN(v)) return stripTime(v);
  const s = String(v || '').replace(/,/g,' ').trim();
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d) ? null : stripTime(d);
}
function stripTime(d){ return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function addDays(d,n){ const x=new Date(d); x.setDate(x.getDate()+n); return x; }
function sameDay_(a,b){ return a.getFullYear()===b.getFullYear() && a.getMonth()===b.getMonth() && a.getDate()===b.getDate(); }

function writeColumn_(sheet, colA1, startRow, values2D){
  const col = sheet.getRange(colA1 + startRow).getColumn();
  const clearRows = sheet.getMaxRows() - startRow + 1;
  if (clearRows > 0) sheet.getRange(startRow, col, clearRows, 1).clearContent(); // 只清内容
  if (values2D.length) sheet.getRange(startRow, col, values2D.length, 1).setValues(values2D);
}

/** 末尾追加 DSP（名字 + 默认值 + 前4项 Route No 备注/着色） */
function appendDSPRows_(tgt, startRowForDSP){
  const cfg = CFG.DSP_APPEND || {};
  const count  = cfg.count || 0;
  if (count <= 0) return;

  const nameCol = tgt.getRange(CFG.DEST_COL_NAME + CFG.DEST_START_ROW).getColumn();
  const names = [];
  for (let i = 1; i <= count; i++) names.push([ (cfg.prefix || 'DSP Driver ') + i ]);
  tgt.getRange(startRowForDSP, nameCol, names.length, 1).setValues(names);

  const d = cfg.defaults || {};
  let col;

  if (CFG.DEST_COL_ROUTE_NO) {
    col = tgt.getRange(CFG.DEST_COL_ROUTE_NO + startRowForDSP).getColumn();
    const routeNoVals = names.map(() => [ (d.routeNo !== undefined ? d.routeNo : '') ]);
    const comments = cfg.routeNoComments || [];
    for (let r = 0; r < Math.min(comments.length, names.length); r++) {
      routeNoVals[r][0] = String(comments[r]);
    }
    tgt.getRange(startRowForDSP, col, routeNoVals.length, 1).setValues(routeNoVals);
    if (cfg.routeNoCommentColor && comments.length) {
      for (let k = 0; k < Math.min(comments.length, names.length); k++) {
        tgt.getRange(startRowForDSP + k, col).setFontColor(cfg.routeNoCommentColor);
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
