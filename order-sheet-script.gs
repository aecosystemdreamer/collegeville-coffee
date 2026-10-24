// Collegeville Coffee Roasters: orders, roast-day signups and the admin dashboard
//
// One Google Sheet, one Apps Script, one URL. The sheet is the record (and the
// backup); admin.html reads and updates it through this script.
//
// Updating from version 1 (about 5 minutes):
// 1. Open the CCR Orders sheet > Extensions > Apps Script.
// 2. Select all, paste this whole file over it, Save.
// 3. Project Settings (gear, left) > Script properties > Add script property:
//      Property: ADMIN_KEY    Value: a long password only you know
//    This is what you type into the admin dashboard to unlock it.
// 4. Pick setup() in the toolbar and click Run once. It adds the new columns
//    (Street, City, State, ZIP, Roast date, Packed) without touching your data.
// 5. Deploy > Manage deployments > pencil > Version: New version > Deploy.
//    The URL stays the same, so the website keeps working.
//
// Gmail caps MailApp at 100 emails/day on a free account (1,500 on Workspace).
// Each order sends 2 (customer + you), so ~50 orders/day before you hit it.

const SHOP = {
  name: 'Collegeville Coffee Roasters',
  alertEmail: '',            // where new-order alerts go; blank = the account running the script
  pickupNote: 'Pickup is in Collegeville, PA. We will email you the address and a time window when your order is ready.',
  paymentNote: 'Within one business day we will email you a secure payment link. We start roasting once payment is in.',
};

// The order a status moves through on the dashboard.
const STATUSES = ['New', 'Paid', 'Roasted', 'Packed', 'Done'];

// Columns are found by header name, so adding or moving columns in the sheet
// doesn't break anything.
const TABS = {
  Orders: ['Received', 'Order #', 'Status', 'Name', 'Email', 'Phone', 'Delivery', 'Ship to', 'Items', 'Subtotal', 'Shipping', 'Total', 'Customer notes', 'Payment link sent', 'Paid', 'Roasted', 'Shipped / picked up', 'Street', 'City', 'State', 'ZIP', 'Roast date', 'Packed'],
  'Order Items': ['Received', 'Order #', 'SKU', 'Item', 'Detail', 'Qty', 'Unit price', 'Line total', 'Gift note'],
  Signups: ['Received', 'Email'],
};

// Run once from the editor. Creates missing tabs and adds missing columns.
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABS).forEach((name) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    const have = sheet.getLastColumn() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
    const missing = TABS[name].filter((h) => !have.includes(h));
    if (missing.length) {
      const start = have.length + 1;
      sheet.getRange(1, start, 1, missing.length).setValues([missing]);
    }
    const width = sheet.getLastColumn();
    sheet.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#2a1a12').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.type === 'signup') { handleSignup(data); return json({ status: 'ok' }); }
    if (data.type === 'order') { handleOrder(data); return json({ status: 'ok' }); }
    if (data.type === 'admin') return json(handleAdmin(data));
    return json({ status: 'error', error: 'unknown request' });
  } catch (err) {
    console.error(err);
    return json({ status: 'error', error: 'server error' });
  }
}

// ============ public: signups and orders ============
function handleSignup(data) {
  const email = clean(data.email, 200);
  if (!isEmail(email)) return;
  const sheet = tab('Signups');
  const existing = sheet.getLastRow() > 1 ? sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat() : [];
  if (existing.map(String).map((s) => s.toLowerCase()).includes(email.toLowerCase())) return;
  sheet.appendRow([new Date(), safe(email)]);
}

function handleOrder(data) {
  const order = {
    orderNo: clean(data.orderNo, 20),
    name: clean(data.name, 100),
    email: clean(data.email, 200),
    phone: clean(data.phone, 40),
    fulfillment: data.fulfillment === 'pickup' ? 'pickup' : 'ship',
    address: clean(data.address, 300),
    street: clean(data.street, 200),
    city: clean(data.city, 100),
    state: clean(data.state, 2).toUpperCase(),
    zip: clean(data.zip, 10),
    roastDate: clean(data.roastDate, 10),
    notes: clean(data.notes, 1000),
    items: (Array.isArray(data.items) ? data.items : []).slice(0, 50).map((i) => ({
      sku: clean(i.sku, 40),
      name: clean(i.name, 100),
      detail: clean(i.detail, 100),
      giftNote: clean(i.giftNote, 300),
      qty: Math.max(1, Math.min(99, parseInt(i.qty, 10) || 1)),
      unitPrice: Number(i.unitPrice) || 0,
    })),
    subtotal: Number(data.subtotal) || 0,
    shipping: Number(data.shipping) || 0,
    total: Number(data.total) || 0,
  };
  if (!order.orderNo || !isEmail(order.email) || !order.items.length) return;

  // Prices come from the browser, so treat totals as a request, not a bill:
  // check them against your price list before sending the payment link.
  const now = new Date();
  const itemsText = order.items.map((i) => `${i.qty} x ${i.name} (${i.detail})`).join('\n');

  appendByHeader(tab('Orders'), {
    'Received': now,
    'Order #': order.orderNo,
    'Status': 'New',
    'Name': safe(order.name),
    'Email': safe(order.email),
    'Phone': safe(order.phone),
    'Delivery': order.fulfillment === 'pickup' ? 'Pickup' : 'Ship',
    'Ship to': safe(order.address),
    'Items': safe(itemsText),
    'Subtotal': order.subtotal,
    'Shipping': order.shipping,
    'Total': order.total,
    'Customer notes': safe(order.notes),
    'Street': safe(order.street),
    'City': safe(order.city),
    'State': safe(order.state),
    'ZIP': safe(order.zip),
    'Roast date': order.roastDate,
  });

  const items = tab('Order Items');
  order.items.forEach((i) => appendByHeader(items, {
    'Received': now,
    'Order #': order.orderNo,
    'SKU': safe(i.sku),
    'Item': safe(i.name),
    'Detail': safe(i.detail),
    'Qty': i.qty,
    'Unit price': i.unitPrice,
    'Line total': i.qty * i.unitPrice,
    'Gift note': safe(i.giftNote),
  }));

  sendCustomerEmail(order);
  sendShopAlert(order);
}

// ============ admin: dashboard reads and status updates ============
function handleAdmin(data) {
  const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!key || String(data.key || '') !== key) {
    Utilities.sleep(800); // slow down guessing
    return { status: 'error', error: 'wrong key' };
  }
  if (data.action === 'list') return { status: 'ok', orders: listOrders() };
  if (data.action === 'setStatus') return setStatus(clean(data.orderNo, 20), clean(data.value, 20));
  return { status: 'error', error: 'unknown action' };
}

function listOrders() {
  const orders = readByHeader(tab('Orders'));
  const items = readByHeader(tab('Order Items'));
  const byOrder = {};
  items.forEach((i) => {
    const no = String(i['Order #']);
    (byOrder[no] = byOrder[no] || []).push({
      sku: String(i['SKU'] || ''),
      name: String(i['Item'] || ''),
      detail: String(i['Detail'] || ''),
      qty: Number(i['Qty']) || 1,
      unitPrice: Number(i['Unit price']) || 0,
      note: String(i['Gift note'] || ''),
    });
  });
  return orders.slice(-300).reverse().map((o) => ({
    orderNo: String(o['Order #']),
    received: o['Received'] instanceof Date ? o['Received'].toISOString() : String(o['Received'] || ''),
    status: STATUSES.includes(o['Status']) ? o['Status'] : 'New',
    name: String(o['Name'] || ''),
    email: String(o['Email'] || ''),
    phone: String(o['Phone'] || ''),
    delivery: String(o['Delivery'] || ''),
    address: String(o['Ship to'] || ''),
    street: String(o['Street'] || ''),
    city: String(o['City'] || ''),
    state: String(o['State'] || ''),
    zip: String(o['ZIP'] || ''),
    roastDate: o['Roast date'] instanceof Date ? Utilities.formatDate(o['Roast date'], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(o['Roast date'] || ''),
    customerNotes: String(o['Customer notes'] || ''),
    subtotal: Number(o['Subtotal']) || 0,
    shipping: Number(o['Shipping']) || 0,
    total: Number(o['Total']) || 0,
    items: byOrder[String(o['Order #'])] || [],
  }));
}

function setStatus(orderNo, value) {
  if (!STATUSES.includes(value)) return { status: 'error', error: 'unknown status' };
  const sheet = tab('Orders');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = (h) => headers.indexOf(h) + 1;
  const last = sheet.getLastRow();
  if (last < 2) return { status: 'error', error: 'order not found' };
  const nos = sheet.getRange(2, col('Order #'), last - 1, 1).getValues().flat().map(String);
  const idx = nos.lastIndexOf(orderNo);
  if (idx < 0) return { status: 'error', error: 'order not found' };
  const row = idx + 2;
  sheet.getRange(row, col('Status')).setValue(value);
  // stamp the matching date column the first time an order reaches a stage
  const stamp = { Paid: 'Paid', Roasted: 'Roasted', Packed: 'Packed', Done: 'Shipped / picked up' }[value];
  if (stamp && col(stamp) > 0) {
    const cell = sheet.getRange(row, col(stamp));
    if (!cell.getValue()) cell.setValue(new Date());
  }
  return { status: 'ok', orderNo: orderNo, value: value };
}

// ============ emails ============
function sendCustomerEmail(o) {
  const first = o.name.split(' ')[0] || 'there';
  const lines = o.items.map((i) =>
    `  ${i.qty} x ${i.name}, ${i.detail}  ${usd(i.qty * i.unitPrice)}` + (i.giftNote ? `\n      Note on the tag: "${i.giftNote}"` : '')).join('\n');
  const body = [
    `Hi ${first},`,
    '',
    `Thanks for your order! It's in the roasting queue.`,
    '',
    `Order ${o.orderNo}`,
    lines,
    '',
    `  Subtotal  ${usd(o.subtotal)}`,
    `  ${o.fulfillment === 'pickup' ? 'Pickup' : 'Shipping'}  ${o.shipping ? usd(o.shipping) : 'Free'}`,
    `  Total     ${usd(o.total)}`,
    '',
    o.fulfillment === 'pickup' ? SHOP.pickupNote : `Shipping to: ${o.address}`,
    '',
    SHOP.paymentNote,
    '',
    'Questions? Just reply to this email.',
    '',
    SHOP.name,
  ].join('\n');

  MailApp.sendEmail({
    to: o.email,
    subject: `Order ${o.orderNo} received | ${SHOP.name}`,
    body,
    name: SHOP.name,
  });
}

function sendShopAlert(o) {
  const to = SHOP.alertEmail || Session.getEffectiveUser().getEmail();
  if (!to) return;
  const lines = o.items.map((i) =>
    `${i.qty} x ${i.name} | ${i.detail} | ${i.sku} | ${usd(i.qty * i.unitPrice)}` + (i.giftNote ? `\n   Note on the tag: "${i.giftNote}"` : '')).join('\n');
  const body = [
    `New order ${o.orderNo}: ${usd(o.total)}`,
    '',
    `${o.name} <${o.email}>${o.phone ? ' ' + o.phone : ''}`,
    o.fulfillment === 'pickup' ? 'LOCAL PICKUP' : `SHIP TO: ${o.address}`,
    o.roastDate ? `Roast day: ${o.roastDate}` : '',
    '',
    lines,
    '',
    `Subtotal ${usd(o.subtotal)} | Shipping ${usd(o.shipping)} | Total ${usd(o.total)}`,
    o.notes ? `\nCustomer notes: ${o.notes}` : '',
    '',
    'Next: check the total against your prices, then send a payment request for the total to the email above.',
    `Sheet: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`,
  ].join('\n');

  MailApp.sendEmail({
    to,
    replyTo: o.email,
    subject: `New order ${o.orderNo}: ${usd(o.total)} (${o.fulfillment === 'pickup' ? 'pickup' : 'ship'})`,
    body,
    name: SHOP.name + ' Orders',
  });
}

// ============ helpers ============
function tab(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error(`Missing tab "${name}". Run setup() once from the editor.`);
  return sheet;
}

function appendByHeader(sheet, values) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.appendRow(headers.map((h) => (h in values ? values[h] : '')));
}

function readByHeader(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(1, 1, last, sheet.getLastColumn()).getValues();
  const headers = values.shift();
  return values.map((r) => Object.fromEntries(headers.map((h, i) => [h, r[i]])));
}

function clean(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u0009\u000B-\u001F]/g, '').trim().slice(0, max);
}

// Stops a cell starting with = + - @ from running as a spreadsheet formula.
function safe(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function isEmail(v) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v);
}

function usd(n) {
  return '$' + Number(n).toFixed(2);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
