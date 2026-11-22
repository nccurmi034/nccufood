/**
 * 金酥鍋貼｜訂單系統後端（Google Apps Script）
 *
 * 功能：
 *  1. doPost：接收網站結帳資料 → 驗證 → 以「菜單」分頁的價格重算金額 → 寫入「訂單」與「訂單明細」
 *  2. 產生訂單編號（JS + 日期 + 當日流水號）
 *  3. 寄通知信給店家與顧客
 *  4. 在「訂單」分頁把「狀態」改成「可取餐」「已出貨」時，自動寄信通知顧客
 *
 * 第一次使用：在編輯器選 setup 函式 → 執行 → 授權。詳見 README.md。
 */

const CFG = {
  SHOP_NAME: '金酥鍋貼',
  SHEET_ORDERS: '訂單',
  SHEET_ITEMS: '訂單明細',
  SHEET_MENU: '菜單',
  OWNER_EMAIL: '',          // 店家收通知信的信箱；留空 = 使用部署這支程式的帳號
  NOTIFY_OWNER: true,       // 有新訂單時寄信給店家
  NOTIFY_CUSTOMER: true,    // 寄訂單確認信給顧客
  FREE_SHIP: 800,           // 宅配滿額免運
  SHIP_FEE: 120,            // 宅配運費
  METHODS: ['內用', '外帶', '宅配'],
  STATUSES: ['新訂單', '製作中', '可取餐', '已出貨', '已完成', '已取消'],
  NOTIFY_STATUSES: ['可取餐', '已出貨']   // 改成這些狀態時，自動寄信給顧客
};

const ORDER_HEADERS = ['訂單編號', '下單時間', '取餐方式', '姓名', '電話', 'Email', '地址', '備註', '品項', '商品小計', '運費', '應付金額', '狀態'];
const ITEM_HEADERS = ['訂單編號', '品項ID', '品名', '單價', '數量', '小計'];
const MENU_HEADERS = ['ID', '品名', '價格', '上架'];

/** 初始菜單：與網站 index.html 的 products / gifts 一致。之後直接改「菜單」分頁即可。 */
const MENU_SEED = [
  [1, '招牌豬肉鍋貼（10入）', 85, true], [2, '韭菜豬肉鍋貼（10入）', 90, true],
  [3, '麻辣牛肉鍋貼（10入）', 110, true], [4, '鮮蝦玉米鍋貼（10入）', 120, true],
  [5, '手工豬肉水餃（12顆）', 95, true], [6, '蝦仁水餃（12顆）', 135, true],
  [7, '紅油抄手（8顆）', 100, true], [8, '酸辣湯', 55, true],
  [9, '玉米濃湯', 45, true], [10, '冰釀梅子綠茶', 40, true],
  [11, '涼拌小黃瓜', 40, true], [12, '滷味拼盤', 70, true],
  [101, '招牌招財鍋貼禮盒（50入）', 520, true], [102, '節慶伴手雙拼禮盒（60入）', 780, true]
];

/* ========== 一次性設定 ========== */

/** 建立分頁、表頭、狀態下拉選單、菜單資料與「狀態變更」觸發器。可重複執行。 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const orders = getOrCreate_(ss, CFG.SHEET_ORDERS, ORDER_HEADERS);
  const statusCol = ORDER_HEADERS.indexOf('狀態') + 1;
  orders.getRange(2, statusCol, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(CFG.STATUSES, true).setAllowInvalid(false).build());
  orders.setColumnWidth(ORDER_HEADERS.indexOf('品項') + 1, 320);

  getOrCreate_(ss, CFG.SHEET_ITEMS, ITEM_HEADERS);

  const menu = getOrCreate_(ss, CFG.SHEET_MENU, MENU_HEADERS);
  if (menu.getLastRow() < 2) menu.getRange(2, 1, MENU_SEED.length, 4).setValues(MENU_SEED);

  const hasTrigger = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'onStatusChange');
  if (!hasTrigger) ScriptApp.newTrigger('onStatusChange').forSpreadsheet(ss).onEdit().create();

  SpreadsheetApp.getUi().alert('設定完成！\n下一步：部署 → 新增部署 → 類型選「網頁應用程式」。');
}

function getOrCreate_(ss, name, headers) {
  let sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#3e2a1d').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ========== Web App ========== */

/** 健康檢查：用瀏覽器開網址，看到 ok:true 代表部署成功 */
function doGet() {
  return json_({ ok: true, service: CFG.SHOP_NAME + ' 訂單系統', time: new Date().toISOString() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    lock.waitLock(20000); locked = true;
    if (!e || !e.postData || !e.postData.contents) throw new Error('沒有收到資料');
    const data = JSON.parse(e.postData.contents);

    if (data.website) return json_({ ok: true, orderNo: 'JS-OK' });        // honeypot：機器人填了隱藏欄位，假裝成功

    const order = validate_(data);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const priced = price_(ss, order);
    const orderNo = nextOrderNo_();
    const now = new Date();

    const itemsText = priced.lines.map(l => l.name + ' ×' + l.qty).join('、');
    ss.getSheetByName(CFG.SHEET_ORDERS).appendRow([
      orderNo, now, order.method, safe_(order.name), safe_(order.phone), safe_(order.email),
      safe_(order.address), safe_(order.note), itemsText, priced.subtotal, priced.shipping, priced.total, '新訂單'
    ]);
    const itemRows = priced.lines.map(l => [orderNo, l.id, l.name, l.price, l.qty, l.price * l.qty]);
    const itemSheet = ss.getSheetByName(CFG.SHEET_ITEMS);
    itemSheet.getRange(itemSheet.getLastRow() + 1, 1, itemRows.length, ITEM_HEADERS.length).setValues(itemRows);
    SpreadsheetApp.flush();

    try { notify_(orderNo, order, priced, now); } catch (mailErr) { console.error('寄信失敗：' + mailErr); }   // 寄信失敗不影響訂單

    return json_({ ok: true, orderNo: orderNo, subtotal: priced.subtotal, shipping: priced.shipping, total: priced.total });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    if (locked) lock.releaseLock();
  }
}

/* ========== 驗證與計價 ========== */

function validate_(d) {
  const str = (v, max) => String(v == null ? '' : v).trim().slice(0, max);
  const o = {
    method: str(d.method, 10), name: str(d.name, 40), phone: str(d.phone, 20),
    email: str(d.email, 100), address: str(d.address, 200), note: str(d.note, 300), items: d.items
  };
  if (CFG.METHODS.indexOf(o.method) < 0) throw new Error('取餐方式不正確');
  if (!o.name) throw new Error('請填寫姓名');
  if (!/^[0-9+\-\s]{8,15}$/.test(o.phone)) throw new Error('電話格式不正確');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(o.email)) throw new Error('Email 格式不正確');
  if (o.method === '宅配' && !o.address) throw new Error('宅配請填寫地址');
  if (!Array.isArray(o.items) || o.items.length < 1 || o.items.length > 50) throw new Error('購物車是空的');
  return o;
}

/** 以「菜單」分頁的價格重新計算，不信任前端傳來的金額 */
function price_(ss, order) {
  const rows = ss.getSheetByName(CFG.SHEET_MENU).getDataRange().getValues().slice(1);
  const menu = {};
  rows.forEach(r => { if (r[0] !== '') menu[String(r[0])] = { name: String(r[1]), price: Number(r[2]), on: r[3] === true || String(r[3]).toUpperCase() === 'TRUE' }; });

  const merged = {};
  order.items.forEach(it => {
    const id = String(it.id), qty = Math.floor(Number(it.qty));
    if (!menu[id] || !menu[id].on) throw new Error('品項不存在或已下架：' + id);
    if (!(qty >= 1 && qty <= 99)) throw new Error('數量不正確');
    merged[id] = (merged[id] || 0) + qty;
  });
  const lines = Object.keys(merged).map(id => ({ id: id, name: menu[id].name, price: menu[id].price, qty: merged[id] }));
  const subtotal = lines.reduce((s, l) => s + l.price * l.qty, 0);
  const shipping = order.method === '宅配' && subtotal < CFG.FREE_SHIP ? CFG.SHIP_FEE : 0;
  return { lines: lines, subtotal: subtotal, shipping: shipping, total: subtotal + shipping };
}

/** 訂單編號：JS + yyMMdd + - + 當日 3 位流水號（在鎖定期間執行，不會重複） */
function nextOrderNo_() {
  const props = PropertiesService.getScriptProperties();
  const day = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd');
  const key = 'seq_' + day;
  const n = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(n));
  return 'JS' + day + '-' + ('00' + n).slice(-3);
}

/** 避免 Excel / Sheets 公式注入：以 = + - @ 開頭的文字前面加上單引號 */
function safe_(v) {
  const s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ========== 通知信 ========== */

function notify_(orderNo, o, p, when) {
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const rows = p.lines.map(l => '<tr><td>' + esc(l.name) + '</td><td align="right">' + l.qty + '</td><td align="right">NT$' + (l.price * l.qty) + '</td></tr>').join('');
  const table = '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:14px">' +
    '<tr style="background:#f3e6d0"><th>品項</th><th>數量</th><th>小計</th></tr>' + rows +
    '<tr><td colspan="2" align="right">運費</td><td align="right">NT$' + p.shipping + '</td></tr>' +
    '<tr><td colspan="2" align="right"><b>應付金額</b></td><td align="right"><b>NT$' + p.total + '</b></td></tr></table>';
  const info = '<p>訂單編號：<b>' + orderNo + '</b><br>取餐方式：' + esc(o.method) +
    (o.address ? '<br>地址：' + esc(o.address) : '') + (o.note ? '<br>備註：' + esc(o.note) : '') + '</p>';

  if (CFG.NOTIFY_OWNER) {
    const to = CFG.OWNER_EMAIL || Session.getEffectiveUser().getEmail();
    MailApp.sendEmail({ to: to, subject: '【新訂單】' + orderNo + '｜' + o.method + '｜NT$' + p.total,
      htmlBody: info + '<p>顧客：' + esc(o.name) + '／' + esc(o.phone) + '／' + esc(o.email) + '</p>' + table });
  }
  if (CFG.NOTIFY_CUSTOMER) {
    MailApp.sendEmail({ to: o.email, name: CFG.SHOP_NAME, subject: '【' + CFG.SHOP_NAME + '】訂單確認 ' + orderNo,
      htmlBody: '<p>' + esc(o.name) + ' 您好，感謝您的訂購！我們已收到您的訂單：</p>' + info + table + '<p>有任何問題請直接回覆此信。</p>' });
  }
}

/** 試算表「訂單」分頁的「狀態」被改成「可取餐」「已出貨」時，自動寄信給顧客（由 setup 建立觸發器） */
function onStatusChange(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== CFG.SHEET_ORDERS || e.range.getRow() < 2) return;
    if (e.range.getColumn() !== ORDER_HEADERS.indexOf('狀態') + 1) return;
    const status = String(e.value || '');
    if (CFG.NOTIFY_STATUSES.indexOf(status) < 0 || !CFG.NOTIFY_CUSTOMER) return;
    const row = sh.getRange(e.range.getRow(), 1, 1, ORDER_HEADERS.length).getValues()[0];
    const orderNo = row[0], name = row[3], email = row[5];
    if (!email) return;
    const msg = status === '可取餐' ? '您的餐點已準備好，歡迎前來取餐！' : '您的訂單已出貨，請留意收件。';
    MailApp.sendEmail({ to: email, name: CFG.SHOP_NAME, subject: '【' + CFG.SHOP_NAME + '】訂單 ' + orderNo + '：' + status,
      htmlBody: '<p>' + name + ' 您好，</p><p>' + msg + '</p><p>訂單編號：<b>' + orderNo + '</b></p>' });
  } catch (err) { console.error(err); }
}
