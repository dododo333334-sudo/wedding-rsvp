/**
 * 于睿&子瑜 婚宴出席回覆｜Google 試算表寫入服務
 * 請從要收資料的試算表開啟「擴充功能 > Apps Script」，貼上本檔後部署為網頁應用程式。
 * 不發布前端、不寄信、不修改既有 Google 表單。
 */
const SHEET_NAME = '婚宴回覆';
const HEADERS = [
  '時間戳記',
  '請求編號',
  '大名',
  '親友',
  '聯絡電話',
  '是否出席',
  '出席人數',
  '葷食份數',
  '素食份數',
  '兒童座椅',
  '特殊飲食或過敏',
  '其他需求',
  '喜帖類型',
  '電子郵件',
  '喜帖寄送地址',
  '祝福'
];

function doGet() {
  return ContentService
    .createTextOutput('于睿&子瑜婚宴回覆服務已啟用')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = parsePayload(e);
    const error = validatePayload(data);
    if (error) {
      return jsonOutput({ result: 'error', error: error });
    }
    const sheet = getSheet();
    if (findExistingRequest(sheet, data.requestId)) {
      return jsonOutput({ result: 'success', duplicate: true });
    }
    sheet.appendRow([
      new Date(),
      data.requestId,
      data.name,
      data.relation,
      data.phone,
      data.attendance,
      data.attendance === '會出席' ? data.guestCount : '',
      data.attendance === '會出席' ? data.meatCount : '',
      data.attendance === '會出席' ? data.vegetarianCount : '',
      data.attendance === '會出席' ? data.childSeats : '',
      data.attendance === '會出席' ? data.dietNotes : '',
      data.attendance === '會出席' ? data.otherNeeds : '',
      data.invitationType,
      needsEmail(data.invitationType) ? data.email : '',
      needsAddress(data.invitationType) ? data.address : '',
      data.blessing || ''
    ]);
    return jsonOutput({ result: 'success' });
  } catch (err) {
    return jsonOutput({ result: 'error', error: '伺服器暫時無法寫入，請稍後再試。' });
  } finally {
    lock.releaseLock();
  }
}

function parsePayload(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('缺少內容');
  }
  return JSON.parse(e.postData.contents);
}

function needsEmail(type) {
  return type === '電子喜帖' || type === '都需要';
}

function needsAddress(type) {
  return type === '實體喜帖' || type === '都需要';
}

function isEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || ''));
}

function validatePayload(data) {
  if (!data || typeof data !== 'object') return '資料格式不正確';
  if (!String(data.requestId || '').trim()) return '缺少請求編號';
  if (!String(data.name || '').trim()) return '請填寫大名';
  if (data.relation !== '新郎' && data.relation !== '新娘') return '請選擇親友關係';
  if (!String(data.phone || '').trim()) return '請填寫聯絡電話';
  if (data.attendance !== '會出席' && data.attendance !== '不克參加') return '請選擇是否出席';
  const types = ['電子喜帖', '實體喜帖', '都需要', '都不需要'];
  if (types.indexOf(data.invitationType) === -1) return '請選擇喜帖類型';
  if (data.attendance === '會出席') {
    const count = Number(data.guestCount);
    const veg = Number(data.vegetarianCount);
    const seats = Number(data.childSeats);
    const meat = Number(data.meatCount);
    if (!Number.isInteger(count) || count < 1 || count > 5) return '出席人數需為 1 至 5 人';
    if (!Number.isInteger(veg) || veg < 0 || veg > count) return '素食份數不可超過出席人數';
    if (!Number.isInteger(seats) || seats < 0 || seats > count) return '兒童座椅數不可超過出席人數';
    if (meat !== count - veg) return '葷素份數計算不一致';
  }
  if (needsEmail(data.invitationType) && !isEmail(data.email)) return '請填寫有效的電子郵件';
  if (needsAddress(data.invitationType) && !String(data.address || '').trim()) return '請填寫喜帖寄送地址';
  return '';
}

function getSpreadsheet() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('找不到試算表');
  return ss;
}

function getSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findExistingRequest(sheet, requestId) {
  const last = sheet.getLastRow();
  if (last < 2) return false;
  const values = sheet.getRange(2, 2, last - 1, 1).getValues();
  for (let i = 0; i < values.length; i += 1) {
    if (String(values[i][0]) === String(requestId)) return true;
  }
  return false;
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
