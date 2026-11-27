/**
 * 金日煎餃 KIN — 線上訂單後端（Google Apps Script）
 *
 * 功能：
 *  - 接收網站結帳送出的訂單（doPost），驗證後寫入 Google 試算表
 *  - 以「商品」工作表的價格重新計算金額，避免前端被竄改價格
 *  - 產生每日流水訂單編號：KIN-YYYYMMDD-001
 *  - 新訂單寄信通知店家；顧客有填 Email 則寄確認信
 *  - 試算表選單：初始化、今日營收摘要、每日營收報表排程
 *
 * 安裝步驟請見 gas/README.md。
 */

// ===== 設定 =====
const CONFIG = {
  SHOP_NAME: '金日煎餃 KIN',
  TIMEZONE: 'Asia/Taipei',
  // 新訂單通知信收件人；留空則寄給執行此腳本的 Google 帳號
  OWNER_EMAIL: '',
  SEND_OWNER_EMAIL: true,
  SEND_CUSTOMER_EMAIL: true,
  // 每日營收報表寄送時間（24 小時制）
  DAILY_REPORT_HOUR: 21,
  STATUSES: ['新訂單', '製作中', '可取餐', '已出貨', '已完成', '已取消'],
  FULFILLMENT: { dine: '內用', take: '外帶', ship: '冷凍宅配' },
  STORES: ['台北旗艦店', '信義店', '中壢店'],
  MAX_QTY: 99,
};

const SHEETS = {
  ORDERS: '訂單',
  ITEMS: '訂單明細',
  PRODUCTS: '商品',
};

const ORDER_HEADERS = ['下單時間', '訂單編號', '狀態', '取餐方式', '取餐門市', '宅配地址', '姓名', '電話', 'Email', '品項', '總數量', '金額', '備註', '來源網址'];
const ITEM_HEADERS = ['訂單編號', '下單時間', '商品代碼', '品名', '單價', '數量', '小計'];
const PRODUCT_HEADERS = ['商品代碼', '品名', '類別', '單價', '上架'];

// 與網站 index.html 的 products 一致；初始化時寫入「商品」工作表，之後以試算表為準
const DEFAULT_PRODUCTS = [
  ['p1', '招牌鮮肉鍋貼', '招牌鍋貼', 85],
  ['p2', '韓式微辣鍋貼', '招牌鍋貼', 95],
  ['p3', '韭菜鮮肉鍋貼', '招牌鍋貼', 90],
  ['p4', '高麗菜鮮肉水餃', '精選水餃', 80],
  ['p5', '鮮蝦水餃', '精選水餃', 120],
  ['p6', '韭菜水餃', '精選水餃', 85],
  ['p7', '暖心蔬菜湯', '湯品／飲料', 45],
  ['p8', '日式焙茶', '湯品／飲料', 40],
  ['p9', '酸辣湯', '湯品／飲料', 50],
  ['p10', '香醋涼拌小黃瓜', '經典小菜', 45],
  ['p11', '季節時蔬', '經典小菜', 55],
  ['g1', '招牌招財鍋貼禮盒 50 入', '冷凍禮盒', 420],
  ['g2', '雙味家庭分享組 80 入', '冷凍禮盒', 650],
];

// ===== Web App 入口 =====

/** 健康檢查：在瀏覽器開啟 Web App 網址會看到這段 JSON。 */
function doGet() {
  return json_({ ok: true, service: CONFIG.SHOP_NAME + ' 訂單 API', time: now_() });
}

/** 網站結帳送出訂單。前端以 text/plain 傳送 JSON，避免 CORS 預檢。 */
function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const result = createOrder_(data);
    return json_(result);
  } catch (err) {
    if (!err.userMessage) console.error(err);
    return json_({ ok: false, error: err.userMessage || '系統忙碌中，請稍後再試或改以電話訂購。' });
  }
}

// ===== 訂單處理 =====

function createOrder_(data) {
  // 機器人防護：隱藏欄位有值就當作成功但不寫入
  if (data.website) return { ok: true, orderId: 'KIN-0', total: 0 };

  const order = validate_(data);
  const ss = SpreadsheetApp.getActive();
  ensureSheets_(ss);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const time = new Date();
    const orderId = nextOrderId_(time);
    const timeText = Utilities.formatDate(time, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    const summary = order.lines.map(l => l.name + ' ×' + l.qty).join('、');
    const qty = order.lines.reduce((n, l) => n + l.qty, 0);

    ss.getSheetByName(SHEETS.ORDERS).appendRow([
      timeText, orderId, CONFIG.STATUSES[0], CONFIG.FULFILLMENT[order.fulfillment],
      order.fulfillment === 'ship' ? '' : order.pickup, order.address,
      order.name, "'" + order.phone, order.email, summary, qty, order.total, order.note, order.page,
    ]);
    const itemRows = order.lines.map(l => [orderId, timeText, l.id, l.name, l.price, l.qty, l.price * l.qty]);
    const items = ss.getSheetByName(SHEETS.ITEMS);
    items.getRange(items.getLastRow() + 1, 1, itemRows.length, ITEM_HEADERS.length).setValues(itemRows);
    SpreadsheetApp.flush();

    const full = Object.assign({ orderId, timeText, summary, qty }, order);
    notify_(full);
    return { ok: true, orderId, total: order.total };
  } finally {
    lock.releaseLock();
  }
}

function validate_(d) {
  const fail = msg => { const e = new Error(msg); e.userMessage = msg; throw e; };
  const str = (v, max) => String(v == null ? '' : v).trim().slice(0, max);

  const name = str(d.name, 60);
  const phone = str(d.phone, 20);
  const email = str(d.email, 120);
  const fulfillment = str(d.fulfillment, 10);
  const pickup = str(d.pickup, 40);
  const address = str(d.address, 200);
  const note = str(d.note, 300);
  const page = str(d.page, 200);

  if (!name) fail('請填寫姓名。');
  if (!/^[0-9+() -]{7,20}$/.test(phone)) fail('聯絡電話格式不正確。');
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) fail('電子郵件格式不正確。');
  if (!CONFIG.FULFILLMENT[fulfillment]) fail('請選擇取餐方式。');
  if (fulfillment === 'ship' && !address) fail('請輸入宅配地址。');
  if (fulfillment !== 'ship' && CONFIG.STORES.indexOf(pickup) < 0) fail('請選擇取餐門市。');
  if (!Array.isArray(d.items) || !d.items.length) fail('購物車沒有商品。');

  const catalog = products_();
  const merged = {};
  d.items.forEach(it => {
    const id = str(it && it.id, 10);
    const qty = Number(it && it.qty);
    if (!catalog[id]) fail('有商品已下架或不存在，請重新整理頁面後再下單。');
    if (!Number.isInteger(qty) || qty < 1 || qty > CONFIG.MAX_QTY) fail('商品數量不正確。');
    merged[id] = Math.min(CONFIG.MAX_QTY, (merged[id] || 0) + qty);
  });
  const lines = Object.keys(merged).map(id => ({ id, name: catalog[id].name, cat: catalog[id].cat, price: catalog[id].price, qty: merged[id] }));
  if (fulfillment === 'ship' && lines.some(l => l.cat !== '冷凍禮盒')) fail('宅配僅限冷凍禮盒，請分開下單。');

  const total = lines.reduce((n, l) => n + l.price * l.qty, 0);
  return { name, phone, email, fulfillment, pickup, address, note, page, lines, total };
}

/** 讀取「商品」工作表（只取上架中的商品）。 */
function products_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.PRODUCTS);
  const map = {};
  if (!sh || sh.getLastRow() < 2) {
    DEFAULT_PRODUCTS.forEach(([id, name, cat, price]) => { map[id] = { name, cat, price }; });
    return map;
  }
  sh.getRange(2, 1, sh.getLastRow() - 1, PRODUCT_HEADERS.length).getValues().forEach(([id, name, cat, price, active]) => {
    if (id && active !== false && Number(price) >= 0) map[String(id).trim()] = { name: String(name), cat: String(cat), price: Number(price) };
  });
  return map;
}

/** 每日流水號，存在 Script Properties。 */
function nextOrderId_(time) {
  const day = Utilities.formatDate(time, CONFIG.TIMEZONE, 'yyyyMMdd');
  const props = PropertiesService.getScriptProperties();
  const key = 'seq_' + day;
  const n = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(n));
  return 'KIN-' + day + '-' + String(n).padStart(3, '0');
}

// ===== 通知信 =====

function notify_(o) {
  const money = n => 'NT$ ' + Number(n).toLocaleString('en-US');
  const lines = o.lines.map(l => `・${l.name} × ${l.qty}　${money(l.price * l.qty)}`).join('\n');
  const where = o.fulfillment === 'ship' ? '宅配地址：' + o.address : '取餐門市：' + o.pickup;
  const body = [
    '訂單編號：' + o.orderId,
    '下單時間：' + o.timeText,
    '取餐方式：' + CONFIG.FULFILLMENT[o.fulfillment],
    where,
    '',
    lines,
    '',
    '合計：' + money(o.total),
  ].join('\n');

  try {
    if (CONFIG.SEND_OWNER_EMAIL) {
      const to = CONFIG.OWNER_EMAIL || Session.getEffectiveUser().getEmail();
      if (to) MailApp.sendEmail(to, `【新訂單】${o.orderId}｜${o.name}｜${money(o.total)}`,
        body + `\n\n顧客：${o.name}　電話：${o.phone}${o.email ? '　Email：' + o.email : ''}\n\n試算表：${SpreadsheetApp.getActive().getUrl()}`);
    }
    if (CONFIG.SEND_CUSTOMER_EMAIL && o.email) {
      MailApp.sendEmail(o.email, `${CONFIG.SHOP_NAME} 訂單確認 ${o.orderId}`,
        `${o.name} 您好，\n\n感謝您的訂購！我們已收到您的訂單：\n\n${body}\n\n付款與取餐細節將由門市與您聯繫。\n\n${CONFIG.SHOP_NAME}`,
        { name: CONFIG.SHOP_NAME });
    }
  } catch (err) {
    // 寄信失敗（例如超過每日配額）不影響訂單成立
    console.warn('寄信失敗', err);
  }
}

// ===== 試算表選單與工具 =====

function onOpen() {
  SpreadsheetApp.getUi().createMenu('金日訂單')
    .addItem('① 初始化工作表', 'setup')
    .addItem('今日營收摘要', 'showTodaySummary')
    .addSeparator()
    .addItem('開啟每日營收報表（每晚寄信）', 'installDailyReport')
    .addItem('關閉每日營收報表', 'removeDailyReport')
    .addToUi();
}

/** 建立工作表、標題列、狀態下拉選單與顏色。可重複執行。 */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);
  ensureSheets_(ss);
  const sh = ss.getSheetByName(SHEETS.PRODUCTS);
  if (sh.getLastRow() < 2) {
    sh.getRange(2, 1, DEFAULT_PRODUCTS.length, PRODUCT_HEADERS.length)
      .setValues(DEFAULT_PRODUCTS.map(p => p.concat([true])));
    sh.getRange(2, 5, DEFAULT_PRODUCTS.length, 1).insertCheckboxes();
  }
  try { SpreadsheetApp.getUi().alert('初始化完成！接著請依 README 部署為網頁應用程式。'); } catch (e) {}
}

function ensureSheets_(ss) {
  const make = (name, headers, widths) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#30251e').setFontColor('#faf6ee');
      sh.setFrozenRows(1);
      (widths || []).forEach((w, i) => w && sh.setColumnWidth(i + 1, w));
    }
    return sh;
  };
  const orders = make(SHEETS.ORDERS, ORDER_HEADERS, [150, 150, 80, 80, 100, 200, 90, 110, 170, 320, 60, 80, 160, 160]);
  make(SHEETS.ITEMS, ITEM_HEADERS, [150, 150, 80, 200, 60, 60, 80]);
  make(SHEETS.PRODUCTS, PRODUCT_HEADERS, [80, 200, 100, 70, 60]);

  // 狀態欄：下拉選單 + 條件式顏色（只設定一次）
  const statusRange = orders.getRange('C2:C');
  if (!statusRange.getDataValidation()) {
    statusRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.STATUSES, true).setAllowInvalid(false).build());
    const colors = { '新訂單': '#fde2cf', '製作中': '#fff1c2', '可取餐': '#d9f0d3', '已出貨': '#d6e6f5', '已完成': '#e6e6e6', '已取消': '#f4cccc' };
    const rules = Object.keys(colors).map(s => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s).setBackground(colors[s]).setRanges([statusRange]).build());
    orders.setConditionalFormatRules(orders.getConditionalFormatRules().concat(rules));
    orders.getRange('L2:L').setNumberFormat('"NT$" #,##0');
  }
}

function todayStats_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.ORDERS);
  const today = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const stats = { date: today, count: 0, total: 0, cancelled: 0, byStatus: {} };
  if (!sh || sh.getLastRow() < 2) return stats;
  sh.getRange(2, 1, sh.getLastRow() - 1, ORDER_HEADERS.length).getDisplayValues().forEach(r => {
    if (String(r[0]).slice(0, 10) !== today) return;
    const status = r[2];
    stats.byStatus[status] = (stats.byStatus[status] || 0) + 1;
    if (status === '已取消') { stats.cancelled++; return; }
    stats.count++;
    stats.total += Number(String(r[11]).replace(/[^0-9.]/g, '')) || 0;
  });
  return stats;
}

function summaryText_(s) {
  const st = Object.keys(s.byStatus).map(k => `${k}：${s.byStatus[k]}`).join('、') || '無';
  return `${s.date} 營收摘要\n\n有效訂單：${s.count} 筆\n營業額：NT$ ${s.total.toLocaleString('en-US')}\n取消：${s.cancelled} 筆\n狀態分布：${st}`;
}

function showTodaySummary() {
  SpreadsheetApp.getUi().alert(summaryText_(todayStats_()));
}

function sendDailyReport() {
  const to = CONFIG.OWNER_EMAIL || Session.getEffectiveUser().getEmail();
  const s = todayStats_();
  MailApp.sendEmail(to, `${CONFIG.SHOP_NAME}｜${s.date} 營收 NT$ ${s.total.toLocaleString('en-US')}（${s.count} 筆）`,
    summaryText_(s) + '\n\n試算表：' + SpreadsheetApp.getActive().getUrl());
}

function installDailyReport() {
  removeDailyReport(true);
  ScriptApp.newTrigger('sendDailyReport').timeBased().everyDays(1).atHour(CONFIG.DAILY_REPORT_HOUR).inTimezone(CONFIG.TIMEZONE).create();
  try { SpreadsheetApp.getUi().alert(`已開啟：每天約 ${CONFIG.DAILY_REPORT_HOUR}:00 寄送營收報表。`); } catch (e) {}
}

function removeDailyReport(silent) {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sendDailyReport').forEach(t => ScriptApp.deleteTrigger(t));
  if (silent !== true) try { SpreadsheetApp.getUi().alert('已關閉每日營收報表。'); } catch (e) {}
}

/** 在 Apps Script 編輯器手動執行，用來測試寫入是否正常。 */
function testOrder() {
  const res = createOrder_({
    name: '測試顧客', phone: '0912345678', email: '', fulfillment: 'take', pickup: CONFIG.STORES[0],
    items: [{ id: 'p1', qty: 2 }, { id: 'p8', qty: 1 }], note: '這是測試訂單', page: 'Apps Script 測試',
  });
  console.log(res);
}

// ===== 共用 =====

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function now_() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}
