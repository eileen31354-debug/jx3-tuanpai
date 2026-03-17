// ── 剑网三团牌记录 · Google Sheets 后端 ────────────────────────────────────
// 使用方式：
//   1. 新建 Google 表格，点击「扩展」→「Apps Script」
//   2. 将本文件全部内容粘贴进去，保存
//   3. 点击「部署」→「新建部署」→ 类型选「网页应用」
//      · 执行身份：我（your Google 账户）
//      · 访问权限：所有人
//   4. 复制生成的网页应用 URL，填入 index.html 中的 SHEET_URL
// ─────────────────────────────────────────────────────────────────────────

const SHEET_NAME = 'records';
const HEADERS    = ['id','tuanpai','tuanzhang','qufu','pingjia','xingxing','combos','beizhu','qita','time'];
const API_KEY    = '97tuanpai==key';

// ── 入口（所有操作均通过 GET 参数路由，避免 CORS 预检问题）─────────────────
function doGet(e) {
  try {
    if (!e || e.parameter.key !== API_KEY) {
      return respond({ error: 'unauthorized' });
    }
    const action = (e && e.parameter && e.parameter.action) || 'getAll';

    if (action === 'getAll') {
      return respond(getAllRecords());
    }
    if (action === 'upsert') {
      const record = JSON.parse(e.parameter.data);
      upsertRecord(record);
      return respond({ ok: true });
    }
    if (action === 'delete') {
      deleteById(e.parameter.id);
      return respond({ ok: true });
    }
    if (action === 'sync') {
      const records = JSON.parse(e.parameter.data);
      syncAll(records);
      return respond({ ok: true });
    }
    if (action === 'upsertBatch') {
      const batch = JSON.parse(e.parameter.data);
      batch.forEach(r => upsertRecord(r));
      return respond({ ok: true });
    }
    if (action === 'export') {
      const url = createWeeklyExport();
      return respond({ ok: true, url });
    }
    return respond({ error: 'unknown action: ' + action });
  } catch (err) {
    return respond({ error: err.toString() });
  }
}

// ── 读取全部记录 ───────────────────────────────────────────────────────────
function getAllRecords() {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let v = row[i];
      if (h === 'combos') {
        if (typeof v === 'string') { try { v = JSON.parse(v); } catch { v = []; } }
        else { v = []; }
      } else if (h === 'xingxing') {
        v = v === true || v === 'TRUE' || v === 1 || v === 'true';
      } else {
        // Google Sheets 可能将数字/日期列返回为非字符串类型，统一转为字符串
        v = (v === null || v === undefined) ? '' : String(v);
      }
      obj[h] = v;
    });
    return obj;
  });
}

// ── 新增 / 更新单条记录 ────────────────────────────────────────────────────
function upsertRecord(record) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getOrCreateSheet();
    const data  = sheet.getDataRange().getValues();
    const idCol = data[0].indexOf('id');
    const row   = buildRow(record);

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(record.id)) {
        sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
        return;
      }
    }
    sheet.appendRow(row);
  } finally {
    lock.releaseLock();
  }
}

// ── 删除单条记录 ───────────────────────────────────────────────────────────
function deleteById(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getOrCreateSheet();
    const data  = sheet.getDataRange().getValues();
    const idCol = data[0].indexOf('id');
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][idCol]) === String(id)) {
        sheet.deleteRow(i + 1);
        return;
      }
    }
  } finally {
    lock.releaseLock();
  }
}

// ── 全量同步（覆盖 Sheet 内所有数据）──────────────────────────────────────
// 用途：首次迁移本地数据到云端，或在同步失败后手动修复
function syncAll(records) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getOrCreateSheet();
    sheet.clearContents();
    sheet.appendRow(HEADERS);
    if (records && records.length) {
      const rows = records.map(r => buildRow(r));
      sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
    }
  } finally {
    lock.releaseLock();
  }
}

// ── 工具函数 ───────────────────────────────────────────────────────────────
function getOrCreateSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function buildRow(record) {
  return HEADERS.map(h => {
    if (h === 'combos')   return JSON.stringify(record[h] || []);
    if (h === 'xingxing') return !!record[h];
    return record[h] !== undefined ? record[h] : '';
  });
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 导出存档（每次生成一个新的 Google Sheet 文件）────────────────────────
const EXPORT_COL_HEADERS = ['评价', '信任', '团牌', '团长ID', '区服', '参战记录', '具体评价', '其他', '日期'];

function recordToExportRow(r) {
  const combosStr = (r.combos || [])
    .filter(c => c.fuben || c.xinfa)
    .map(c => (c.fuben && c.xinfa) ? `${c.fuben}/${c.xinfa}` : (c.fuben || c.xinfa))
    .join('、');
  return [r.pingjia, r.xingxing ? '★' : '', r.tuanpai, r.tuanzhang,
          r.qufu, combosStr, r.beizhu, r.qita, r.time];
}

// 生成一次导出（手动或定时触发器调用均可）
function createWeeklyExport() {
  const records   = getAllRecords();
  const dateStr   = Utilities.formatDate(new Date(), 'Asia/Shanghai', 'yyyy-MM-dd');
  const fileName  = `团牌记录存档_${dateStr}`;

  // 在 Google Drive 根目录创建新 Spreadsheet
  const newSS     = SpreadsheetApp.create(fileName);
  const sheet     = newSS.getActiveSheet();
  sheet.setName('记录');

  // 写表头
  sheet.appendRow(EXPORT_COL_HEADERS);
  const headerRange = sheet.getRange(1, 1, 1, EXPORT_COL_HEADERS.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#F7F4EF');

  // 写数据
  if (records.length) {
    const rows = records.map(recordToExportRow);
    sheet.getRange(2, 1, rows.length, EXPORT_COL_HEADERS.length).setValues(rows);
  }

  // 自动调整列宽
  sheet.autoResizeColumns(1, EXPORT_COL_HEADERS.length);

  // 在原始 Sheet 中记录导出日志
  logExport(fileName, newSS.getUrl(), records.length);

  return newSS.getUrl();
}

// 在主 Sheet 里记录每次导出（方便追溯）
function logExport(fileName, url, count) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName('导出日志');
  if (!sheet) {
    sheet = ss.insertSheet('导出日志');
    sheet.appendRow(['导出时间', '文件名', '记录数', '链接']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }
  const timeStr = Utilities.formatDate(new Date(), 'Asia/Shanghai', 'yyyy-MM-dd HH:mm');
  sheet.appendRow([timeStr, fileName, count, url]);
}

// ── 定时触发器设置（在 Apps Script 编辑器手动运行一次即可）──────────────
// 注意：运行前请先在「项目设置」中将时区改为 Asia/Shanghai（中国标准时间）
function setupWeeklyTrigger() {
  // 删除已有的同名触发器，防止重复
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'createWeeklyExport')
    .forEach(t => ScriptApp.deleteTrigger(t));

  // 每周日 12:00（北京时间，需项目时区设为 Asia/Shanghai）
  ScriptApp.newTrigger('createWeeklyExport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(12)
    .create();

  Logger.log('✓ 已设置每周日 12:00（北京时间）自动导出触发器');
}
