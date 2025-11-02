/* Code.gs for Sajha Sangeet - Inventory, Customers, Sales, Billing
   Place this content in Code.gs in your Apps Script project.
*/

const INVENTORY_SHEET = "Inventory";
const CUSTOMERS_SHEET = "Customers";
const SALES_SHEET = "Sales_Log";

/**
 * onOpen - build custom menu
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sajha Sangeet')
    .addItem('Open App', 'showSidebar')
    .addItem('Initialize Sheets (one-time)', 'initSheets')
    .addToUi();
}

/**
 * showSidebar - show the application UI
 */
function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Sajha Sangeet — Shop Manager')
    .setWidth(900);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * doGet - optional: allow deployment as web app
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('Sajha Sangeet');
}

/**
 * initSheets - create sheets with headers if they don't exist and sample data
 */
function initSheets() {
  const ss = SpreadsheetApp.getActive();
  // Inventory
  let inv = ss.getSheetByName(INVENTORY_SHEET);
  if (!inv) {
    inv = ss.insertSheet(INVENTORY_SHEET);
    inv.appendRow(["ItemID","ItemName","Category","BuyPrice","SellPrice","StockQty","ReorderLevel","Notes"]);
    inv.appendRow([generateId("I"),"Acoustic Guitar (Yamaha FG)","Guitars",10000,15000,10,2,"Good condition"]);
    inv.appendRow([generateId("I"),"Electric Guitar (Squier)","Guitars",15000,22000,5,1,"With case"]);
  }
  // Customers
  let cust = ss.getSheetByName(CUSTOMERS_SHEET);
  if (!cust) {
    cust = ss.insertSheet(CUSTOMERS_SHEET);
    cust.appendRow(["CustomerID","Name","Phone","Email","Address","Notes"]);
    cust.appendRow([generateId("C"),"Ram Shrestha","+977-98xxxxxxx","ram@example.com","Dhangadhi","Regular customer"]);
  }
  // Sales Log
  let sales = ss.getSheetByName(SALES_SHEET);
  if (!sales) {
    sales = ss.insertSheet(SALES_SHEET);
    sales.appendRow(["SaleID","Date","CustomerID","CustomerName","ItemsJSON","TotalAmount","PaymentMethod","Notes"]);
  }
  SpreadsheetApp.getUi().alert('Sheets initialized (or already existed).');
}

/* -----------------------------
   Utilities & CRUD server functions
   -----------------------------*/

/**
 * generateId - simple id generator
 */
function generateId(prefix) {
  const ts = new Date().getTime().toString().slice(-6);
  return prefix + ts + Math.floor(Math.random()*90+10);
}

/**
 * readSheetData - returns array of objects
 */
function readSheetData(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const headers = data.shift().map(h=>String(h));
  return data.map(r=>{
    const obj = {};
    for (let i=0;i<headers.length;i++) obj[headers[i]] = r[i];
    return obj;
  });
}

/* Inventory CRUD */
function getInventory() { return readSheetData(INVENTORY_SHEET); }

function addInventory(item) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(INVENTORY_SHEET);
  if (!sh) throw new Error("Inventory sheet not found. Run Initialize Sheets.");
  const id = generateId("I");
  const row = [id, item.ItemName || "", item.Category || "", Number(item.BuyPrice||0), Number(item.SellPrice||0), Number(item.StockQty||0), Number(item.ReorderLevel||0), item.Notes || ""];
  sh.appendRow(row);
  return { success:true, id };
}

function updateInventory(itemId, updates) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(INVENTORY_SHEET);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  for (let r=1;r<data.length;r++){
    if (String(data[r][0]) === String(itemId)) {
      // apply updates by header
      for (let c=0;c<headers.length;c++){
        const key = headers[c];
        if (updates.hasOwnProperty(key)) {
          sh.getRange(r+1, c+1).setValue(updates[key]);
        }
      }
      return {success:true};
    }
  }
  return {success:false, message:"Item not found"};
}

function deleteInventory(itemId) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(INVENTORY_SHEET);
  const data = sh.getDataRange().getValues();
  for (let r=1;r<data.length;r++){
    if (String(data[r][0]) === String(itemId)) {
      sh.deleteRow(r+1);
      return {success:true};
    }
  }
  return {success:false, message:"Item not found"};
}

/* Customers CRUD */
function getCustomers() { return readSheetData(CUSTOMERS_SHEET); }

function addCustomer(cust) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CUSTOMERS_SHEET);
  const id = generateId("C");
  const row = [id, cust.Name||"", cust.Phone||"", cust.Email||"", cust.Address||"", cust.Notes||""];
  sh.appendRow(row);
  return {success:true, id};
}

function updateCustomer(customerId, updates) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CUSTOMERS_SHEET);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  for (let r=1;r<data.length;r++){
    if (String(data[r][0]) === String(customerId)) {
      for (let c=0;c<headers.length;c++){
        const key = headers[c];
        if (updates.hasOwnProperty(key)) {
          sh.getRange(r+1, c+1).setValue(updates[key]);
        }
      }
      return {success:true};
    }
  }
  return {success:false, message:"Customer not found"};
}

function deleteCustomer(customerId) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CUSTOMERS_SHEET);
  const data = sh.getDataRange().getValues();
  for (let r=1;r<data.length;r++){
    if (String(data[r][0]) === String(customerId)) {
      sh.deleteRow(r+1);
      return {success:true};
    }
  }
  return {success:false, message:"Customer not found"};
}

/* Sales and billing */
function getSales() { return readSheetData(SALES_SHEET); }

/**
 * createSale(payload)
 * payload = {CustomerID, CustomerName, Items: [{ItemID, ItemName, UnitPrice, Qty}], PaymentMethod, Notes}
 */
function createSale(payload) {
  const ss = SpreadsheetApp.getActive();
  const invSh = ss.getSheetByName(INVENTORY_SHEET);
  if (!invSh) throw new Error("Inventory sheet missing.");
  const salesSh = ss.getSheetByName(SALES_SHEET);
  if (!salesSh) throw new Error("Sales sheet missing.");
  const saleId = generateId("S");
  const date = new Date();
  // compute totals and check stock
  let total = 0;
  const items = payload.Items.map(it=>{
    const subtotal = Number(it.UnitPrice||0) * Number(it.Qty||0);
    total += subtotal;
    return {
      ItemID: it.ItemID,
      ItemName: it.ItemName,
      UnitPrice: Number(it.UnitPrice||0),
      Qty: Number(it.Qty||0),
      Subtotal: subtotal
    };
  });
  // update inventory stock (reduce), fail if insufficient
  const invData = invSh.getDataRange().getValues(); // include headers
  const invHeader = invData[0];
  // make map row index by ItemID
  const idToRow = {};
  for (let r=1;r<invData.length;r++){
    idToRow[String(invData[r][0])] = r+1; // sheet row index
  }
  // check stock
  for (let i=0;i<items.length;i++){
    const it = items[i];
    const rowNum = idToRow[String(it.ItemID)];
    if (!rowNum) throw new Error("Inventory item not found: " + it.ItemID);
    const stockQty = Number(invSh.getRange(rowNum, invHeader.indexOf("StockQty")+1).getValue());
    if (stockQty < it.Qty) {
      throw new Error("Insufficient stock for " + it.ItemName + ". Available: " + stockQty);
    }
  }
  // reduce stock now
  for (let i=0;i<items.length;i++){
    const it = items[i];
    const rowNum = idToRow[String(it.ItemID)];
    const stockCell = invSh.getRange(rowNum, invHeader.indexOf("StockQty")+1);
    const current = Number(stockCell.getValue());
    stockCell.setValue(current - it.Qty);
  }
  // save sale
  const itemsJson = JSON.stringify(items);
  salesSh.appendRow([saleId, date, payload.CustomerID || "", payload.CustomerName || "", itemsJson, total, payload.PaymentMethod || "", payload.Notes || ""]);
  // Return printable bill html via helper
  const billHtml = buildBillHtml({SaleID: saleId, Date: date, CustomerID: payload.CustomerID, CustomerName: payload.CustomerName, Items: items, TotalAmount: total, PaymentMethod: payload.PaymentMethod, Notes: payload.Notes});
  return {success:true, saleId, billHtml};
}

/* helper build bill html */
function buildBillHtml(sale) {
  const shopName = "Sajha Sangeet";
  let rows = "";
  sale.Items.forEach(it=>{
    rows += `<tr>
      <td>${escapeHtml(it.ItemName)}</td>
      <td style="text-align:center">${it.Qty}</td>
      <td style="text-align:right">${formatNumber(it.UnitPrice)}</td>
      <td style="text-align:right">${formatNumber(it.Subtotal)}</td>
    </tr>`;
  });
  const html = `
    <html><head><meta charset="utf-8"><title>Bill - ${sale.SaleID}</title>
    <style>
      body{font-family:Arial, Helvetica, sans-serif; padding:20px;}
      .header{ text-align:center; }
      table{width:100%; border-collapse:collapse; margin-top:10px;}
      td,th{border-bottom:1px solid #ddd; padding:6px;}
      .right{text-align:right;}
    </style>
    </head><body>
      <div class="header">
        <h2>${shopName}</h2>
        <div>Sale ID: ${sale.SaleID}</div>
        <div>Date: ${Utilities.formatDate(new Date(sale.Date), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")}</div>
        <div>Customer: ${escapeHtml(sale.CustomerName || "")}</div>
      </div>
      <table>
        <thead><tr><th>Item</th><th style="text-align:center">Qty</th><th style="text-align:right">Unit</th><th style="text-align:right">Subtotal</th></tr></thead>
        <tbody>
          ${rows}
        </tbody>
        <tfoot>
          <tr><td colspan="3" style="text-align:right"><strong>Total</strong></td><td style="text-align:right"><strong>${formatNumber(sale.TotalAmount)}</strong></td></tr>
        </tfoot>
      </table>
      <div style="margin-top:12px;">Payment: ${escapeHtml(sale.PaymentMethod||"")}</div>
      <div style="margin-top:20px; text-align:center;">Thank you for shopping at ${shopName}!</div>
    </body></html>
  `;
  return html;
}

/* Utility formatting/escaping */
function formatNumber(n) {
  return Number(n).toLocaleString();
}
function escapeHtml(text) {
  if (text === null || text === undefined) return "";
  return String(text).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

/* Summary / chart data */
function getSummaryData() {
  const sales = getSales();
  // produce total sales per day (simple)
  const map = {};
  sales.forEach(s=>{
    const d = new Date(s.Date);
    const key = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
    map[key] = (map[key] || 0) + Number(s.TotalAmount||0);
  });
  const rows = Object.keys(map).sort().map(k=>[k, map[k]]);
  // inventory low stock
  const inv = getInventory();
  const low = inv.filter(i => Number(i.StockQty) <= Number(i.ReorderLevel||0));
  return {salesByDate: rows, lowStock: low};
}

/* Expose HTML files if needed */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
