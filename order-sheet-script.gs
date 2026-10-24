// Collegeville Coffee Roasters: order + roast-day signup backend
//
// One Google Sheet, one Apps Script, one URL for the whole site.
//
// Setup (about 10 minutes):
// 1. Create a new Google Sheet named "CCR Orders".
// 2. Extensions > Apps Script. Delete what's there, paste this whole file, Save.
// 3. Fill in SHOP below (your email for new-order alerts, pickup details).
// 4. In the Apps Script toolbar pick setup() and click Run once. It creates the
//    three tabs (Orders, Order Items, Signups) with headers. Google will ask you
//    to authorize "edit spreadsheets" and "send email as you". That's expected.
// 5. Deploy > New deployment > type "Web app" > Execute as "Me" >
//    Who has access "Anyone". Copy the Web app URL.
// 6. Paste the URL into index.html as SHEET_WEB_APP_URL.
//
// If you edit this script later: Deploy > Manage deployments > edit (pencil) >
// Version "New version" > Deploy. The URL stays the same.
//
// Gmail caps MailApp at 100 emails/day on a free account (1,500 on Workspace).
// Each order sends 2 (customer + you), so ~50 orders/day before you hit it.

const SHOP = {
  name: 'Collegeville Coffee Roasters',
  alertEmail: '',            // where new-order alerts go; blank = the account running the script
  pickupNote: 'Pickup is in Collegeville, PA. We will email you the address and a time window when your order is ready.',
  paymentNote: 'Within one business day we will email you a secure payment link. We start roasting once payment is in.',
};

const TABS = {
  orders: ['Received', 'Order #', 'Status', 'Name', 'Email', 'Phone', 'Delivery', 'Ship to', 'Items', 'Subtotal', 'Shipping', 'Total', 'Customer notes', 'Payment link sent', 'Paid', 'Roasted', 'Shipped / picked up'],
  items: ['Received', 'Order #', 'SKU', 'Item', 'Detail', 'Qty', 'Unit price', 'Line total', 'Gift note'],
  signups: ['Received', 'Email'],
};

// Run once from the editor to create the tabs.
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = { orders: 'Orders', items: 'Order Items', signups: 'Signups' };
  Object.keys(TABS).forEach((key) => {
    let sheet = ss.getSheetByName(names[key]);
    if (!sheet) sheet = ss.insertSheet(names[key]);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(TABS[key]);
      sheet.getRange(1, 1, 1, TABS[key].length).setFontWeight('bold').setBackground('#3b2117').setFontColor('#f0e2c8');
      sheet.setFrozenRows(1);
    }
  });
  const blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1 && blank.getLastRow() === 0) ss.deleteSheet(blank);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.type === 'signup') handleSignup(data);
    else if (data.type === 'order') handleOrder(data);
    return json({ status: 'ok' });
  } catch (err) {
    console.error(err);
    return json({ status: 'error' });
  }
}

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

  tab('Orders').appendRow([
    now, order.orderNo, 'New', safe(order.name), safe(order.email), safe(order.phone),
    order.fulfillment === 'pickup' ? 'Pickup' : 'Ship', safe(order.address), safe(itemsText),
    order.subtotal, order.shipping, order.total, safe(order.notes), '', '', '', '',
  ]);

  const itemRows = order.items.map((i) => [
    now, order.orderNo, safe(i.sku), safe(i.name), safe(i.detail), i.qty, i.unitPrice, i.qty * i.unitPrice, safe(i.giftNote),
  ]);
  const items = tab('Order Items');
  items.getRange(items.getLastRow() + 1, 1, itemRows.length, itemRows[0].length).setValues(itemRows);

  sendCustomerEmail(order);
  sendShopAlert(order);
}

function sendCustomerEmail(o) {
  const first = o.name.split(' ')[0] || 'there';
  const lines = o.items.map((i) =>
    `  ${i.qty} x ${i.name}, ${i.detail}  ${usd(i.qty * i.unitPrice)}` + (i.giftNote ? `\n      Gift note: "${i.giftNote}"` : '')).join('\n');
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
    `${i.qty} x ${i.name} | ${i.detail} | ${i.sku} | ${usd(i.qty * i.unitPrice)}` + (i.giftNote ? `\n   Gift note: "${i.giftNote}"` : '')).join('\n');
  const body = [
    `New order ${o.orderNo}: ${usd(o.total)}`,
    '',
    `${o.name} <${o.email}>${o.phone ? ' ' + o.phone : ''}`,
    o.fulfillment === 'pickup' ? 'LOCAL PICKUP' : `SHIP TO: ${o.address}`,
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
