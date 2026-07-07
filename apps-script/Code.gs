// iLemeMade — Google Apps Script API v2
// Deploy as: Execute as "Me", Access "Anyone (anonymous)".

const SPREADSHEET_ID = '1oRLuRspyOfbo-_x9uZJzSGhn_DKa7CrDnyZOlg1PKL8';

const SHEETS = {
  PRODUCTS:          'Products',
  SUPPLIES:          'Supplies',
  PURCHASES:         'Purchases',
  PRODUCT_MATERIALS: 'ProductMaterials',
  INVENTORY:         'Inventory',
  LABOR:             'Labor',
  SALES:             'Sales',
  ADJUSTMENTS:       'Adjustments',
  DASHBOARD:         'Dashboard',
};

// ─── Spreadsheet menu (runs automatically when the sheet is opened) ───────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('iLemeMade')
    .addItem('📦 Record a Sale',          'showSalesForm')
    .addItem('🔧 Inventory Adjustment',   'showAdjustmentForm')
    .addSeparator()
    .addItem('📊 Refresh Dashboard',      'refreshDashboard')
    .addToUi();
}

function showSalesForm() {
  const html = HtmlService.createTemplateFromFile('SalesForm')
    .evaluate()
    .setTitle('Record a Sale')
    .setWidth(500);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showAdjustmentForm() {
  const html = HtmlService.createTemplateFromFile('AdjustmentForm')
    .evaluate()
    .setTitle('Inventory Adjustment')
    .setWidth(500);
  SpreadsheetApp.getUi().showSidebar(html);
}

function refreshDashboard() {
  getDashboard();
  SpreadsheetApp.getUi().alert('Dashboard refreshed!');
}

// ─── Router ───────────────────────────────────────────────────────────────────

function doGet(e) {
  // Serve HTML forms when ?page= is present
  if (e && e.parameter && e.parameter.page) {
    const page = e.parameter.page;
    if (page === 'sales') {
      return HtmlService.createTemplateFromFile('SalesForm')
        .evaluate()
        .setTitle('Record a Sale — iLemeMade')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    if (page === 'adjust') {
      return HtmlService.createTemplateFromFile('AdjustmentForm')
        .evaluate()
        .setTitle('Inventory Adjustment — iLemeMade')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }

  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'ping';
  let result;
  try {
    switch (action) {
      case 'ping':                result = { status: 'ok', message: 'iLemeMade API v2 running!' }; break;
      case 'getProducts':         result = getProducts(); break;
      case 'getSupplies':         result = getSupplies(); break;
      case 'getSupplyLevels':     result = getSupplyLevels(); break;
      case 'getPurchases':        result = getPurchases(e.parameter); break;
      case 'getProductMaterials': result = getProductMaterials(e.parameter); break;
      case 'getInventory':        result = getInventory(e.parameter); break;
      case 'getItem':             result = getItem(e.parameter.itemId); break;
      case 'getSales':            result = getSales(e.parameter); break;
      case 'getAdjustments':      result = getAdjustments(e.parameter); break;
      case 'getLabor':            result = getLabor(e.parameter.itemId); break;
      case 'getDashboard':        result = getDashboard(); break;
      case 'setup':               result = setupSheets(); break;
      case 'seedData':            result = seedSampleData(); break;
      default:                    result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message, stack: err.stack };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (_) { return jsonOut({ error: 'Invalid JSON body' }); }
  let result;
  try {
    switch (body.action) {
      case 'addPurchase':        result = addPurchase(body.data); break;
      case 'addProductMaterial': result = addProductMaterial(body.data); break;
      case 'addInventoryItem':   result = addInventoryItem(body.data); break;
      case 'addSale':            result = addSale(body.data); break;
      case 'addLabor':           result = addLabor(body.data); break;
      case 'updateItemStatus':   result = updateItemStatus(body.itemId, body.status); break;
      case 'addAdjustment':      result = addAdjustment(body.data); break;
      case 'decrementStock':     result = decrementStock(body.productId, body.qty); break;
      default:                   result = { error: 'Unknown POST action: ' + body.action };
    }
  } catch (err) { result = { error: err.message }; }
  return jsonOut(result);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function getSheet(name) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet not found: ' + name);
  return sheet;
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1)
    .filter(row => row.some(cell => cell !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) return val.toISOString().split('T')[0];
  return String(val);
}

function round2(n) { return Math.round(n * 100) / 100; }

// ─── READ: Catalog ────────────────────────────────────────────────────────────

function getProducts() {
  return sheetToObjects(getSheet(SHEETS.PRODUCTS)).map(r => ({
    productId:           String(r.ProductID    || ''),
    productName:         String(r.ProductName  || ''),
    category:            String(r.Category     || ''),
    description:         String(r.Description  || ''),
    trackingType:        String(r.TrackingType || 'Individual'), // 'Individual' | 'Quantity'
    defaultLaborMinutes: Number(r.DefaultLaborMinutes) || 0,
    defaultSalePrice:    Number(r.DefaultSalePrice)    || 0,
    stockQty:            r.StockQty !== '' ? Number(r.StockQty) : null,
  }));
}

function getSupplies() {
  return sheetToObjects(getSheet(SHEETS.SUPPLIES)).map(r => ({
    supplyId:   String(r.SupplyID   || ''),
    supplyName: String(r.SupplyName || ''),
    category:   String(r.Category   || ''),
    unit:       String(r.Unit       || ''),
    notes:      String(r.Notes      || ''),
  }));
}

function getPurchases(params) {
  let rows = sheetToObjects(getSheet(SHEETS.PURCHASES)).map(r => ({
    purchaseId:     String(r.PurchaseID      || ''),
    supplyId:       String(r.SupplyID        || ''),
    purchaseDate:   formatDate(r.PurchaseDate),
    supplier:       String(r.Supplier        || ''),
    productName:    String(r.ProductName     || ''),
    quantityBought: Number(r.QuantityBought) || 0,
    totalPaid:      Number(r.TotalPaid)      || 0,
    receiptUrl:     String(r.ReceiptURL      || ''),
    notes:          String(r.Notes           || ''),
    costPerUnit:    round2((Number(r.TotalPaid) || 0) / (Number(r.QuantityBought) || 1)),
  }));
  if (params && params.supplyId) rows = rows.filter(r => r.supplyId === params.supplyId);
  return rows;
}

function getProductMaterials(params) {
  let rows = sheetToObjects(getSheet(SHEETS.PRODUCT_MATERIALS)).map(r => ({
    productId:  String(r.ProductID  || ''),
    supplyId:   String(r.SupplyID   || ''),
    amountUsed: Number(r.AmountUsed) || 0,
    notes:      String(r.Notes      || ''),
  }));
  if (params && params.productId) rows = rows.filter(r => r.productId === params.productId);
  return rows;
}

// ─── READ: Inventory / Labor / Sales ─────────────────────────────────────────

function getInventory(params) {
  let items = sheetToObjects(getSheet(SHEETS.INVENTORY)).map(r => ({
    itemId:      String(r.ItemID    || ''),
    productId:   String(r.ProductID || ''),
    status:      String(r.Status    || ''),
    color:       String(r.Color     || ''),
    dateCreated: formatDate(r.DateCreated),
    photoUrl:    String(r.PhotoURL  || ''),
    story:       String(r.Story     || ''),
  }));
  if (params && params.status)    items = items.filter(i => i.status.toLowerCase() === params.status.toLowerCase());
  if (params && params.productId) items = items.filter(i => i.productId === params.productId);
  return items;
}

function getLabor(itemId) {
  const rows = sheetToObjects(getSheet(SHEETS.LABOR));
  const row  = rows.find(r => String(r.ItemID) === String(itemId));
  if (!row) return { itemId, minutesWorked: 0, hourlyRate: 0, laborCost: 0 };
  const minutesWorked = Number(row.MinutesWorked) || 0;
  const hourlyRate    = Number(row.HourlyRate)    || 0;
  return { itemId, minutesWorked, hourlyRate, laborCost: round2((minutesWorked / 60) * hourlyRate) };
}

function getSales(params) {
  let sales = sheetToObjects(getSheet(SHEETS.SALES)).map(r => ({
    date:         formatDate(r.Date),
    productId:    String(r.ProductID     || ''),  // Quantity items
    itemId:       String(r.ItemID        || ''),  // Individual items
    qty:          Number(r.Qty)          || 1,
    priceEach:    Number(r.PriceEach)    || 0,
    subTotal:     Number(r.SubTotal)     || 0,
    taxCollected: Number(r.TaxCollected) || 0,   // Etsy collects — not our money
    ourRevenue:   Number(r.OurRevenue)   || 0,   // What we actually keep
    channel:      String(r.Channel       || ''),
    salesOrder:   String(r.SalesOrder    || ''),
    notes:        String(r.Notes         || ''),
  }));
  if (params && params.itemId)    sales = sales.filter(s => s.itemId    === params.itemId);
  if (params && params.productId) sales = sales.filter(s => s.productId === params.productId);
  if (params && params.channel)   sales = sales.filter(s => s.channel.toLowerCase() === params.channel.toLowerCase());
  return sales;
}

function getAdjustments(params) {
  let rows = sheetToObjects(getSheet(SHEETS.ADJUSTMENTS)).map(r => ({
    date:           formatDate(r.Date),
    productId:      String(r.ProductID      || ''),
    itemId:         String(r.ItemID         || ''),
    adjustmentType: String(r.AdjustmentType || ''),
    qty:            Number(r.Qty)           || 0,
    notes:          String(r.Notes          || ''),
  }));
  if (params && params.productId) rows = rows.filter(r => r.productId === params.productId);
  if (params && params.itemId)    rows = rows.filter(r => r.itemId    === params.itemId);
  return rows;
}

// ─── READ: Calculated material cost (recipe × weighted avg purchase price) ────

function getItemMaterialCost(productId) {
  const materials = getProductMaterials({ productId });
  const purchases = getPurchases({});

  const bySupply = {};
  purchases.forEach(p => {
    if (!bySupply[p.supplyId]) bySupply[p.supplyId] = { totalPaid: 0, totalQty: 0 };
    bySupply[p.supplyId].totalPaid += p.totalPaid;
    bySupply[p.supplyId].totalQty  += p.quantityBought;
  });
  const avgCost = {};
  Object.keys(bySupply).forEach(id => {
    const b = bySupply[id];
    avgCost[id] = b.totalQty > 0 ? b.totalPaid / b.totalQty : 0;
  });

  const lineItems = materials.map(m => ({
    supplyId:    m.supplyId,
    amountUsed:  m.amountUsed,
    costPerUnit: round2(avgCost[m.supplyId] || 0),
    lineCost:    round2(m.amountUsed * (avgCost[m.supplyId] || 0)),
    notes:       m.notes,
  }));

  return { productId, total: round2(lineItems.reduce((s, l) => s + l.lineCost, 0)), lineItems };
}

// ─── READ: Supply levels (purchased − used across all made items) ─────────────

function getSupplyLevels() {
  const supplies  = getSupplies();
  const purchases = getPurchases({});
  const materials = getProductMaterials({});
  const inventory = getInventory({});
  const products  = getProducts();
  const sales     = getSales({});

  // Individual items: count every row in Inventory (supplies consumed when made, not when sold)
  const madeByProduct = {};
  inventory.forEach(item => {
    madeByProduct[item.productId] = (madeByProduct[item.productId] || 0) + 1;
  });

  // Quantity items: made = current StockQty + units sold + units removed without sale
  // (damaged/gifted/promo items still consumed supplies when made)
  const soldQtyByProduct = {};
  sales.filter(s => s.productId).forEach(s => {
    soldQtyByProduct[s.productId] = (soldQtyByProduct[s.productId] || 0) + s.qty;
  });

  const adjustments = getAdjustments({});
  const removedQtyByProduct = {};
  adjustments.filter(a => a.productId && a.qty < 0).forEach(a => {
    removedQtyByProduct[a.productId] = (removedQtyByProduct[a.productId] || 0) + Math.abs(a.qty);
  });

  products.filter(p => p.trackingType === 'Quantity').forEach(p => {
    madeByProduct[p.productId] = (p.stockQty || 0)
      + (soldQtyByProduct[p.productId]   || 0)
      + (removedQtyByProduct[p.productId]|| 0);
  });

  const totalPurchased = {};
  const totalPaid      = {};
  purchases.forEach(p => {
    totalPurchased[p.supplyId] = (totalPurchased[p.supplyId] || 0) + p.quantityBought;
    totalPaid[p.supplyId]      = (totalPaid[p.supplyId]      || 0) + p.totalPaid;
  });

  const totalUsed = {};
  materials.forEach(m => {
    const made = madeByProduct[m.productId] || 0;
    totalUsed[m.supplyId] = (totalUsed[m.supplyId] || 0) + (m.amountUsed * made);
  });

  return supplies.map(s => {
    const purchased  = round2(totalPurchased[s.supplyId] || 0);
    const used       = round2(totalUsed[s.supplyId]      || 0);
    const remaining  = round2(purchased - used);
    const paid       = round2(totalPaid[s.supplyId]      || 0);
    return {
      supplyId:       s.supplyId,
      supplyName:     s.supplyName,
      category:       s.category,
      unit:           s.unit,
      purchased,
      used,
      remaining,
      totalPaid:      paid,
      avgCostPerUnit: purchased > 0 ? round2(paid / purchased) : 0,
      low:            remaining < (purchased * 0.20),
    };
  });
}

// ─── READ: Full detail for one Individual item ────────────────────────────────

function getItem(itemId) {
  if (!itemId) return { error: 'itemId required' };
  const inventory = sheetToObjects(getSheet(SHEETS.INVENTORY));
  const row       = inventory.find(r => String(r.ItemID) === itemId);
  if (!row) return { error: 'Item not found: ' + itemId };

  const labor      = getLabor(itemId);
  const matCost    = getItemMaterialCost(String(row.ProductID || ''));
  const saleList   = getSales({ itemId });
  const saleRecord = saleList[0] || null;
  const ourRevenue = saleRecord ? saleRecord.ourRevenue : 0;
  const totalCost  = round2(matCost.total + labor.laborCost);

  return {
    itemId:      String(row.ItemID    || ''),
    productId:   String(row.ProductID || ''),
    status:      String(row.Status    || ''),
    color:       String(row.Color     || ''),
    dateCreated: formatDate(row.DateCreated),
    photoUrl:    String(row.PhotoURL  || ''),
    story:       String(row.Story     || ''),
    materialCost: matCost,
    labor,
    sale:        saleRecord,
    financials:  { materialCost: matCost.total, laborCost: labor.laborCost, totalCost, ourRevenue, profit: round2(ourRevenue - totalCost) },
  };
}

// ─── READ: Dashboard ──────────────────────────────────────────────────────────

function getDashboard() {
  const sales     = getSales({});
  const inventory = getInventory({});
  const products  = getProducts();

  const revenue = round2(sales.reduce((sum, s) => sum + s.ourRevenue, 0));

  // Expenses for Individual sold items
  let totalExpenses = 0;
  const soldItemIds = [...new Set(sales.filter(s => s.itemId).map(s => s.itemId))];
  soldItemIds.forEach(itemId => {
    const item = inventory.find(i => i.itemId === itemId);
    if (!item) return;
    totalExpenses += getItemMaterialCost(item.productId).total + getLabor(itemId).laborCost;
  });

  // Expenses for Quantity sold items (material + default labor estimate)
  const soldQtyByProduct = {};
  sales.filter(s => s.productId).forEach(s => {
    soldQtyByProduct[s.productId] = (soldQtyByProduct[s.productId] || 0) + s.qty;
  });
  Object.entries(soldQtyByProduct).forEach(([pid, qty]) => {
    const mat     = getItemMaterialCost(pid);
    const product = products.find(p => p.productId === pid);
    const labor   = product ? (product.defaultLaborMinutes / 60) * 5 : 0; // $5/hr default
    totalExpenses += (mat.total + labor) * qty;
  });
  totalExpenses = round2(totalExpenses);

  const profit       = round2(revenue - totalExpenses);
  const savings      = round2(profit * 0.50);
  const reinvestment = round2(profit * 0.25);
  const spending     = round2(profit * 0.25);

  // Best seller by OurRevenue
  const revByProduct = {};
  sales.forEach(s => {
    const pid = s.productId || (inventory.find(i => i.itemId === s.itemId) || {}).productId || 'Unknown';
    revByProduct[pid] = (revByProduct[pid] || 0) + s.ourRevenue;
  });
  const bestPid    = Object.keys(revByProduct).sort((a, b) => revByProduct[b] - revByProduct[a])[0] || '—';
  const bestProd   = products.find(p => p.productId === bestPid);
  const bestSeller = bestProd ? bestProd.productName : bestPid;

  // Inventory count: individual Available + quantity stock
  const individualAvail = inventory.filter(i => i.status === 'Available').length;
  const quantityStock   = products.filter(p => p.trackingType === 'Quantity')
                                  .reduce((sum, p) => sum + (p.stockQty || 0), 0);

  const supplyLevels = getSupplyLevels();

  const metrics = {
    revenue,
    expenses:            totalExpenses,
    profit,
    inventoryCount:      individualAvail + quantityStock,
    unitsSold:           sales.reduce((sum, s) => sum + s.qty, 0),
    bestSeller,
    savingsBalance:      savings,
    reinvestmentBalance: reinvestment,
    spendingBalance:     spending,
    lowSupplies:         supplyLevels.filter(s => s.low).map(s => s.supplyName),
    generatedAt:         new Date().toISOString(),
  };

  try { writeDashboardSheet(metrics, supplyLevels, products); } catch (_) {}
  return metrics;
}

function writeDashboardSheet(metrics, supplyLevels, products) {
  const sheet = getSheet(SHEETS.DASHBOARD);
  sheet.clearContents();

  const rows = [
    ['── Financials ──',         ''],
    ['Revenue (our take)',        '$' + metrics.revenue],
    ['Expenses',                  '$' + metrics.expenses],
    ['Profit',                    '$' + metrics.profit],
    ['Units Sold',                 metrics.unitsSold],
    ['Best Seller',                metrics.bestSeller],
    [''],
    ['── Savings Buckets ──',    ''],
    ['Savings (50%)',             '$' + metrics.savingsBalance],
    ['Reinvestment (25%)',        '$' + metrics.reinvestmentBalance],
    ['Spending (25%)',            '$' + metrics.spendingBalance],
    [''],
    ['── Inventory ──',          ''],
    ['Total in Stock',             metrics.inventoryCount],
  ];

  // Quantity products stock breakdown
  products.filter(p => p.trackingType === 'Quantity').forEach(p => {
    rows.push(['  ' + p.productName, (p.stockQty || 0) + ' in stock']);
  });

  rows.push(['']);
  rows.push(['── Supply Levels ──', '']);
  supplyLevels.forEach(s => {
    rows.push([s.supplyName + ' (' + s.unit + ')', s.remaining + ' remaining' + (s.low ? ' ⚠️ LOW' : '')]);
  });
  rows.push(['']);
  rows.push(['Last Updated', metrics.generatedAt]);

  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 200);
}

// ─── WRITE endpoints ──────────────────────────────────────────────────────────

function addPurchase(data) {
  const costPerUnit = round2((Number(data.totalPaid) || 0) / (Number(data.quantityBought) || 1));
  getSheet(SHEETS.PURCHASES).appendRow([
    data.purchaseId   || 'PUR-' + Date.now(),
    data.supplyId,
    data.purchaseDate || new Date().toISOString().split('T')[0],
    data.supplier     || '',
    data.productName  || '',
    Number(data.quantityBought) || 0,
    Number(data.totalPaid)      || 0,
    costPerUnit,
    data.receiptUrl   || '',
    data.notes        || '',
  ]);
  return { success: true, costPerUnit };
}

function addProductMaterial(data) {
  getSheet(SHEETS.PRODUCT_MATERIALS).appendRow([data.productId, data.supplyId, Number(data.amountUsed) || 0, data.notes || '']);
  return { success: true };
}

function addInventoryItem(data) {
  getSheet(SHEETS.INVENTORY).appendRow([
    data.itemId, data.productId, data.status || 'Available', data.color || '',
    data.dateCreated || new Date().toISOString().split('T')[0], data.photoUrl || '', data.story || '',
  ]);
  return { success: true, itemId: data.itemId };
}

function addSale(data) {
  const qty       = Number(data.qty)          || 1;
  const priceEach = Number(data.priceEach)    || Number(data.salePrice) || 0;
  const subTotal  = round2(Number(data.subTotal)     || priceEach * qty);
  const taxColl   = round2(Number(data.taxCollected) || 0);
  const ourRev    = round2(Number(data.ourRevenue)   || subTotal);

  getSheet(SHEETS.SALES).appendRow([
    data.date       || new Date().toISOString().split('T')[0],
    data.productId  || '',
    data.itemId     || '',
    qty, priceEach, subTotal, taxColl, ourRev,
    data.channel    || '',
    data.salesOrder || '',
    data.notes      || '',
  ]);

  if (data.productId && !data.itemId) {
    decrementStock(data.productId, qty);
  } else if (data.itemId) {
    try { updateItemStatus(data.itemId, 'Sold'); } catch (_) {}
  }
  return { success: true };
}

function decrementStock(productId, qty) {
  const sheet  = getSheet(SHEETS.PRODUCTS);
  const data   = sheet.getDataRange().getValues();
  const idCol  = data[0].indexOf('ProductID');
  const qtyCol = data[0].indexOf('StockQty');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(productId)) {
      const current = Number(data[i][qtyCol]) || 0;
      const newQty  = Math.max(0, current - qty);
      sheet.getRange(i + 1, qtyCol + 1).setValue(newQty);
      return { success: true, newStock: newQty };
    }
  }
  return { error: 'Product not found: ' + productId };
}

function incrementStock(productId, qty) {
  // decrementStock with a negative qty adds to stock (current - (-n) = current + n)
  return decrementStock(productId, -Math.abs(qty));
}

function addAdjustment(data) {
  // AdjustmentType: Production (+), Gift (-), Promo (-), Damaged (-), Other (±)
  const qty  = Number(data.qty) || 0;
  const type = String(data.adjustmentType || data.AdjustmentType || '');

  // Positive types add stock; negative types remove it
  const addTypes = ['Production'];
  const removeTypes = ['Gift', 'Promo', 'Damaged'];

  let delta = qty; // qty should already be signed (+/-) from the form
  if (addTypes.includes(type)    && qty < 0) delta = Math.abs(qty); // force positive
  if (removeTypes.includes(type) && qty > 0) delta = -qty;           // force negative

  getSheet(SHEETS.ADJUSTMENTS).appendRow([
    data.date           || new Date().toISOString().split('T')[0],
    data.productId      || '',
    data.itemId         || '',
    type,
    delta,
    data.notes          || '',
  ]);

  // Update StockQty for Quantity products
  if (data.productId) {
    if (delta > 0) {
      incrementStock(data.productId, delta);
    } else if (delta < 0) {
      decrementStock(data.productId, Math.abs(delta));
    }
  }

  return { success: true, delta };
}

function addLabor(data) {
  getSheet(SHEETS.LABOR).appendRow([data.itemId, Number(data.minutesWorked) || 0, Number(data.hourlyRate) || 0]);
  return { success: true };
}

function updateItemStatus(itemId, status) {
  const sheet = getSheet(SHEETS.INVENTORY);
  const data  = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('ItemID');
  const stCol = data[0].indexOf('Status');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(itemId)) {
      sheet.getRange(i + 1, stCol + 1).setValue(status);
      return { success: true };
    }
  }
  return { error: 'Item not found: ' + itemId };
}

// ─── Form handlers (called via google.script.run from HTML forms) ─────────────

function getProductsForForm() {
  return getProducts().map(p => ({
    productId:    p.productId,
    productName:  p.productName,
    trackingType: p.trackingType,
    stockQty:     p.stockQty,
    defaultPrice: p.defaultSalePrice,
  }));
}

function submitSale(data) {
  try {
    const result = addSale(data);
    return { success: true, result };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function submitAdjustment(data) {
  try {
    const result = addAdjustment(data);
    return { success: true, result };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ─── SETUP ────────────────────────────────────────────────────────────────────

function setupSheets() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const log = [];

  const schema = {
    // Products: TrackingType = 'Individual' | 'Quantity'. StockQty only used for Quantity.
    [SHEETS.PRODUCTS]:          ['ProductID','ProductName','Category','Description','TrackingType','DefaultLaborMinutes','DefaultSalePrice','StockQty'],
    [SHEETS.SUPPLIES]:          ['SupplyID','SupplyName','Category','Unit','Notes'],
    [SHEETS.PURCHASES]:         ['PurchaseID','SupplyID','PurchaseDate','Supplier','ProductName','QuantityBought','TotalPaid','CostPerUnit','ReceiptURL','Notes'],
    [SHEETS.PRODUCT_MATERIALS]: ['ProductID','SupplyID','AmountUsed','Notes'],
    [SHEETS.INVENTORY]:         ['ItemID','ProductID','Status','Color','DateCreated','PhotoURL','Story'],
    [SHEETS.LABOR]:             ['ItemID','MinutesWorked','HourlyRate'],
    // Sales: ProductID for Quantity items, ItemID for Individual items (one will be blank per row).
    // TaxCollected = Etsy-collected tax (not our money). OurRevenue = what we keep.
    [SHEETS.SALES]:             ['Date','ProductID','ItemID','Qty','PriceEach','SubTotal','TaxCollected','OurRevenue','Channel','SalesOrder','Notes'],
    [SHEETS.ADJUSTMENTS]:       ['Date','ProductID','ItemID','AdjustmentType','Qty','Notes'],
    [SHEETS.DASHBOARD]:         ['Metric','Value'],
  };

  Object.entries(schema).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) { sheet = ss.insertSheet(name); log.push('Created: ' + name); }
    else { log.push('Exists:  ' + name); }
    if (!sheet.getRange(1, 1).getValue()) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      styleHeader(sheet, headers.length);
    }
  });

  return { success: true, log };
}

function styleHeader(sheet, numCols) {
  const range = sheet.getRange(1, 1, 1, numCols);
  range.setBackground('#2a1f3d');
  range.setFontColor('#ffffff');
  range.setFontWeight('bold');
  sheet.setFrozenRows(1);
}

// ─── DEMO seed data ───────────────────────────────────────────────────────────

function seedSampleData() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const log = [];

  const prodSheet = ss.getSheetByName(SHEETS.PRODUCTS);
  if (prodSheet.getLastRow() < 2) {
    prodSheet.getRange(2, 1, 9, 8).setValues([
      ['DEMO-BC-001','[DEMO] BearCat Keychain',       'Keychain', '[DEMO] Gummy bear cat resin keychain', 'Individual',45,12.00,''],
      ['DEMO-HK-001','[DEMO] Hello Kitty Resin',      'Resin Art','[DEMO] Hello Kitty in clear resin',    'Individual',60,18.00,''],
      ['DEMO-CB-001','[DEMO] Chicken Banana',         'Figurine', '[DEMO] Chicken riding a banana',       'Individual',30,10.00,''],
      ['DEMO-DR-001','[DEMO] 3D Printed Dragon',      '3D Print', '[DEMO] Articulated dragon',            'Individual',20, 8.00,''],
      ['DEMO-RL-001','[DEMO] Resin Lemon',            'Resin Art','[DEMO] Lemon slice resin art',         'Individual',75,22.00,''],
      ['DEMO-SH-001','[DEMO] Trail Shoe Art',         'Resin Art','[DEMO] Trail shoe resin casting',      'Individual',90,28.00,''],
      ['DEMO-RF-001','[DEMO] Recycled Filament Cast', '3D Print', '[DEMO] Recycled filament art',         'Individual',15, 6.00,''],
      ['DEMO-BR-001','[DEMO] Handmade Bracelet',      'Jewelry',  '[DEMO] Beaded charm bracelet',         'Individual',30, 8.00,''],
      ['DEMO-RK-001','[DEMO] Hand Painted Rock',      'Painted',  '[DEMO] Hand painted river rock',       'Individual',20, 5.00,''],
    ]);
    log.push('Seeded Products: 9 rows');
  } else { log.push('Products skipped'); }

  const supSheet = ss.getSheetByName(SHEETS.SUPPLIES);
  if (supSheet.getLastRow() < 2) {
    supSheet.getRange(2, 1, 10, 5).setValues([
      ['DEMO-RES-001','[DEMO] Clear Casting Resin',      'Resin',    'oz',  '[DEMO] 2-part A+B clear casting resin'],
      ['DEMO-RES-002','[DEMO] UV Resin',                 'Resin',    'oz',  '[DEMO] Single-part UV-cure resin'],
      ['DEMO-FIL-001','[DEMO] PLA Filament — Black',     'Filament', 'g',   '[DEMO] Standard black PLA'],
      ['DEMO-FIL-002','[DEMO] PLA Filament — Blue',      'Filament', 'g',   '[DEMO] Ocean blue PLA'],
      ['DEMO-PAI-001','[DEMO] Acrylic Paint Set',        'Paint',    'oz',  '[DEMO] Multi-color acrylic set'],
      ['DEMO-HDW-001','[DEMO] Keychain Hook + Ring Set', 'Hardware', 'each','[DEMO] Lobster clasp + jump ring, silver'],
      ['DEMO-HDW-002','[DEMO] LED Light Base',           'Hardware', 'each','[DEMO] USB LED display stand, 7-color'],
      ['DEMO-PKG-001','[DEMO] Clear Poly Bag 3×4"',      'Packaging','each','[DEMO] Resealable display bags'],
      ['DEMO-PKG-002','[DEMO] Kraft Hang Tag',           'Packaging','each','[DEMO] Pre-punched price/story tags'],
      ['DEMO-GIT-001','[DEMO] Gold Holographic Glitter', 'Other',    'oz',  '[DEMO] Fine holographic gold glitter'],
    ]);
    log.push('Seeded Supplies: 10 rows');
  } else { log.push('Supplies skipped'); }

  const purSheet = ss.getSheetByName(SHEETS.PURCHASES);
  if (purSheet.getLastRow() < 2) {
    purSheet.getRange(2, 1, 11, 10).setValues([
      ['DEMO-PUR-001','DEMO-RES-001','2025-02-15','Amazon',      'Alumilite Clear 32oz Kit',     32,24.99,0.781,'','[DEMO] First resin purchase'],
      ['DEMO-PUR-002','DEMO-HDW-001','2025-02-15','Hobby Lobby', 'Lobster Clasp + Ring 50pk',    50, 6.99,0.140,'','[DEMO] 50-pack silver hardware'],
      ['DEMO-PUR-003','DEMO-FIL-001','2025-02-20','Amazon',      'Hatchbox PLA Black 1kg',     1000,19.99,0.020,'','[DEMO] 1kg spool'],
      ['DEMO-PUR-004','DEMO-PAI-001','2025-02-20',"Michael's",   'Apple Barrel Acrylic 12-pack', 12, 8.99,0.749,'','[DEMO] 12 × 1oz bottles'],
      ['DEMO-PUR-005','DEMO-PKG-001','2025-02-20','Amazon',      'Clear Poly Bag 3x4 100pk',    100, 5.99,0.060,'','[DEMO] 100-pack bags'],
      ['DEMO-PUR-006','DEMO-PKG-002','2025-02-20','Amazon',      'Kraft Hang Tags 50pk',         50, 3.99,0.080,'','[DEMO] 50-pack tags'],
      ['DEMO-PUR-007','DEMO-GIT-001','2025-03-01','Hobby Lobby', 'Gold Holo Glitter 2oz',         2, 3.49,1.745,'','[DEMO] Fine gold holographic'],
      ['DEMO-PUR-008','DEMO-HDW-002','2025-03-05','Amazon',      'LED Display Base 5pk',          5, 9.99,1.998,'','[DEMO] 7-color USB bases'],
      ['DEMO-PUR-009','DEMO-FIL-002','2025-03-10','Amazon',      'Hatchbox PLA Blue 1kg',      1000,21.99,0.022,'','[DEMO] Ocean blue spool'],
      ['DEMO-PUR-010','DEMO-RES-001','2025-05-01','Amazon',      'Alumilite Clear 32oz Kit',     32,26.99,0.844,'','[DEMO] Reorder'],
      ['DEMO-PUR-011','DEMO-RES-002','2025-04-10','Amazon',      "LET'S RESIN UV Resin 17oz",    17,18.99,1.117,'','[DEMO] UV resin'],
    ]);
    log.push('Seeded Purchases: 11 rows');
  } else { log.push('Purchases skipped'); }

  const matSheet = ss.getSheetByName(SHEETS.PRODUCT_MATERIALS);
  if (matSheet.getLastRow() < 2) {
    matSheet.getRange(2, 1, 26, 4).setValues([
      ['DEMO-BC-001','DEMO-RES-001',0.15,'[DEMO] 0.15 oz by weight'],
      ['DEMO-BC-001','DEMO-HDW-001',1,   '[DEMO] 1 hook + ring'],
      ['DEMO-BC-001','DEMO-PKG-001',1,   '[DEMO] 1 bag'],
      ['DEMO-BC-001','DEMO-PKG-002',1,   '[DEMO] 1 hang tag'],
      ['DEMO-HK-001','DEMO-RES-001',0.20,'[DEMO] 0.20 oz larger mold'],
      ['DEMO-HK-001','DEMO-GIT-001',0.01,'[DEMO] Pinch of glitter'],
      ['DEMO-HK-001','DEMO-PKG-001',1,   '[DEMO] 1 bag'],
      ['DEMO-HK-001','DEMO-PKG-002',1,   '[DEMO] 1 hang tag'],
      ['DEMO-CB-001','DEMO-RES-001',0.18,'[DEMO] 0.18 oz figurine mold'],
      ['DEMO-CB-001','DEMO-PAI-001',0.02,'[DEMO] Detail painting'],
      ['DEMO-CB-001','DEMO-PKG-001',1,   '[DEMO] 1 bag'],
      ['DEMO-DR-001','DEMO-FIL-001',8,   '[DEMO] ~8g at 15% infill'],
      ['DEMO-DR-001','DEMO-PKG-001',1,   '[DEMO] 1 bag'],
      ['DEMO-DR-001','DEMO-PKG-002',1,   '[DEMO] 1 hang tag'],
      ['DEMO-RL-001','DEMO-RES-001',0.25,'[DEMO] 0.25 oz lemon mold'],
      ['DEMO-RL-001','DEMO-GIT-001',0.01,'[DEMO] Gold glitter'],
      ['DEMO-RL-001','DEMO-PKG-001',1,   '[DEMO] 1 bag'],
      ['DEMO-RL-001','DEMO-PKG-002',1,   '[DEMO] 1 hang tag'],
      ['DEMO-SH-001','DEMO-RES-001',0.30,'[DEMO] 0.30 oz shoe mold'],
      ['DEMO-SH-001','DEMO-PAI-001',0.03,'[DEMO] Sole + lace painting'],
      ['DEMO-SH-001','DEMO-GIT-001',0.01,'[DEMO] Glitter accent'],
      ['DEMO-SH-001','DEMO-PKG-001',1,   '[DEMO] 1 bag'],
      ['DEMO-SH-001','DEMO-PKG-002',1,   '[DEMO] 1 hang tag'],
      ['DEMO-RK-001','DEMO-PAI-001',0.05,'[DEMO] ~0.05 oz paint'],
      ['DEMO-RK-001','DEMO-PKG-001',1,   '[DEMO] 1 bag'],
      ['DEMO-RK-001','DEMO-PKG-002',1,   '[DEMO] 1 hang tag'],
    ]);
    log.push('Seeded ProductMaterials: 26 rows');
  } else { log.push('ProductMaterials skipped'); }

  const invSheet = ss.getSheetByName(SHEETS.INVENTORY);
  if (invSheet.getLastRow() < 2) {
    invSheet.getRange(2, 1, 12, 7).setValues([
      ['DEMO-HK-001-001','DEMO-HK-001','Sold',     'Pink/White', '2025-03-01','','[DEMO] My first Hello Kitty! Pink glitter.'],
      ['DEMO-HK-001-002','DEMO-HK-001','Available','Blue/Silver','2025-03-15','','[DEMO] Sparkly blue with silver stars.'],
      ['DEMO-HK-001-003','DEMO-HK-001','Available','Rainbow',    '2025-04-01','','[DEMO] Rainbow swirl — my favorite!'],
      ['DEMO-BC-001-001','DEMO-BC-001','Sold',     'Clear/Green','2025-02-20','','[DEMO] Green gummy cat, first keychain sold!'],
      ['DEMO-BC-001-002','DEMO-BC-001','Available','Purple',     '2025-04-10','','[DEMO] Purple sparkle BearCat.'],
      ['DEMO-BC-001-003','DEMO-BC-001','Available','Orange',     '2025-04-20','','[DEMO] Halloween orange BearCat.'],
      ['DEMO-DR-001-001','DEMO-DR-001','Available','Blue',       '2025-05-01','','[DEMO] Ocean blue articulated dragon.'],
      ['DEMO-DR-001-002','DEMO-DR-001','Sold',     'Red',        '2025-04-05','','[DEMO] Red dragon sold at school fair!'],
      ['DEMO-RL-001-001','DEMO-RL-001','Available','Yellow',     '2025-05-10','','[DEMO] Classic lemon slice with gold flake.'],
      ['DEMO-RK-001-001','DEMO-RK-001','Sold',     'Multi',      '2025-03-10','','[DEMO] Ladybug rock, sold to grandma.'],
      ['DEMO-RK-001-002','DEMO-RK-001','Available','Multi',      '2025-05-20','','[DEMO] Rainbow paw print rock.'],
      ['DEMO-BR-001-001','DEMO-BR-001','Sold',     'Pink/Gold',  '2025-04-15','','[DEMO] Pink bracelet with gold butterfly charm.'],
    ]);
    log.push('Seeded Inventory: 12 rows');
  } else { log.push('Inventory skipped'); }

  const laborSheet = ss.getSheetByName(SHEETS.LABOR);
  if (laborSheet.getLastRow() < 2) {
    laborSheet.getRange(2, 1, 12, 3).setValues([
      ['DEMO-HK-001-001',60,5],['DEMO-HK-001-002',45,5],['DEMO-HK-001-003',75,5],
      ['DEMO-BC-001-001',40,5],['DEMO-BC-001-002',40,5],['DEMO-BC-001-003',40,5],
      ['DEMO-DR-001-001',20,5],['DEMO-DR-001-002',20,5],['DEMO-RL-001-001',90,5],
      ['DEMO-RK-001-001',20,5],['DEMO-RK-001-002',20,5],['DEMO-BR-001-001',30,5],
    ]);
    log.push('Seeded Labor: 12 rows');
  } else { log.push('Labor skipped'); }

  // Sales: Date, ProductID, ItemID, Qty, PriceEach, SubTotal, TaxCollected, OurRevenue, Channel, SalesOrder, Notes
  const salesSheet = ss.getSheetByName(SHEETS.SALES);
  if (salesSheet.getLastRow() < 2) {
    salesSheet.getRange(2, 1, 5, 11).setValues([
      ['2025-03-05','','DEMO-HK-001-001',1,18,18,0,18,'In Person','','[DEMO] Spring craft fair'],
      ['2025-03-12','','DEMO-BC-001-001',1,12,12,0,12,'In Person','','[DEMO] School fundraiser'],
      ['2025-04-07','','DEMO-DR-001-002',1, 8, 8,0, 8,'Etsy',    '','[DEMO] First Etsy sale!'],
      ['2025-04-18','','DEMO-RK-001-001',1, 5, 5,0, 5,'In Person','','[DEMO] Sold to grandma'],
      ['2025-04-22','','DEMO-BR-001-001',1, 8, 8,0, 8,'In Person','','[DEMO] Mom\'s coworker'],
    ]);
    log.push('Seeded Sales: 5 rows');
  } else { log.push('Sales skipped'); }

  return { success: true, log };
}

// ─── Utilities ────────────────────────────────────────────────────────────────

function clearAllSheets() {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const temp = ss.insertSheet('_temp');
  ['Products','Supplies','Purchases','ProductMaterials','Inventory','Labor','Sales','Adjustments','Dashboard','Costs'].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) { ss.deleteSheet(sheet); Logger.log('Deleted: ' + name); }
  });
  ss.deleteSheet(temp);
  Logger.log('✅ Cleared. Now run firstTimeSetup.');
}

function firstTimeSetup() {
  setupSheets();
  seedSampleData();
  Logger.log('✅ Setup complete. Check the spreadsheet.');
}

// ─── One-time: add real GummyBearCat products ─────────────────────────────────

function addRealGummyCats() {
  const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  const prodSheet = ss.getSheetByName(SHEETS.PRODUCTS);
  const existingIds = sheetToObjects(prodSheet).map(r => String(r.ProductID));

  // ProductID, ProductName, Category, Description, TrackingType, DefaultLaborMinutes, DefaultSalePrice, StockQty
  const products = [
    ['BC-001','BearCat Keychain — Patriotic','Keychain','Patriotic colorway gummy bear cat resin keychain','Quantity',45,8.00,11],
    ['BC-002','BearCat Keychain — Pink',     'Keychain','Pink colorway gummy bear cat resin keychain',     'Quantity',45,8.00, 5],
    ['BC-003','BearCat Keychain — Fuchsia',  'Keychain','Fuchsia colorway gummy bear cat resin keychain',  'Quantity',45,8.00, 6],
    ['BC-004','BearCat Keychain — Red Swirl','Keychain','Red swirl colorway gummy bear cat resin keychain','Quantity',45,5.00, 5],
  ];

  const toAdd = products.filter(p => !existingIds.includes(p[0]));
  if (toAdd.length > 0) {
    prodSheet.getRange(prodSheet.getLastRow() + 1, 1, toAdd.length, 8).setValues(toAdd);
    Logger.log('Added ' + toAdd.length + ' products: BC-001 through BC-004');
  } else {
    Logger.log('BC-001–BC-004 already exist — skipped');
  }
  Logger.log('BC-001 Patriotic  11 in stock @ $8\nBC-002 Pink         5 in stock @ $8\nBC-003 Fuchsia      6 in stock @ $8\nBC-004 Red Swirl    5 in stock @ $5');
}

// ─── One-time: record the first real sales ────────────────────────────────────
// Sale 1 — Etsy (Dawn): 2× BC-001 Patriotic @ $8. Etsy collected $1.34 tax (not our money).
//           Shipping was refunded (hand delivered) so OurRevenue = SubTotal only.
// Sale 2 — In Person (grandma): 1× BC-004 Red Swirl @ $5 + 1× BC-001 Patriotic @ $5.

function recordRealSales() {
  const ss         = SpreadsheetApp.openById(SPREADSHEET_ID);
  const salesSheet = ss.getSheetByName(SHEETS.SALES);

  // Date, ProductID, ItemID, Qty, PriceEach, SubTotal, TaxCollected, OurRevenue, Channel, SalesOrder, Notes
  const sales = [
    ['2026-07-06','BC-001','',2,8.00,16.00,1.34,16.00,'Etsy',           '','First Etsy sale — Dawn. Tax collected by Etsy. Shipping refunded (hand delivery).'],
    ['2026-07-06','BC-004','',1,5.00, 5.00,0.00, 5.00,'In Person - Cash','','Grandma — Red Swirl'],
    ['2026-07-06','BC-001','',1,5.00, 5.00,0.00, 5.00,'In Person - Cash','','Grandma — Patriotic'],
  ];

  salesSheet.getRange(salesSheet.getLastRow() + 1, 1, sales.length, 11).setValues(sales);

  decrementStock('BC-001', 3); // 2 Etsy + 1 grandma
  decrementStock('BC-004', 1); // 1 grandma

  Logger.log('✅ Recorded 3 sale rows.');
  Logger.log('BC-001 Patriotic: −3 (was 11, now 8)');
  Logger.log('BC-004 Red Swirl: −1 (was 5, now 4)');
  Logger.log('Total revenue: $26.00  |  Tax collected by Etsy (not ours): $1.34');
}
