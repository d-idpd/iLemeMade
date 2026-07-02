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
  DASHBOARD:         'Dashboard',
};

// ─── Router ───────────────────────────────────────────────────────────────────

function doGet(e) {
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
      case 'getLabor':            result = getLabor(e.parameter.itemId); break;
      case 'getDashboard':        result = getDashboard(); break;
      case 'setup':               result = setupSheets(); break;
      case 'seedData':            result = seedSampleData(); break;
      default:                    result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message, stack: err.stack };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (_) {
    return jsonOut({ error: 'Invalid JSON body' });
  }

  let result;
  try {
    switch (body.action) {
      case 'addPurchase':         result = addPurchase(body.data); break;
      case 'addProductMaterial':  result = addProductMaterial(body.data); break;
      case 'addInventoryItem':    result = addInventoryItem(body.data); break;
      case 'addSale':             result = addSale(body.data); break;
      case 'addLabor':            result = addLabor(body.data); break;
      case 'updateItemStatus':    result = updateItemStatus(body.itemId, body.status); break;
      default:                    result = { error: 'Unknown POST action: ' + body.action };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return jsonOut(result);
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
    productId:          String(r.ProductID   || ''),
    productName:        String(r.ProductName || ''),
    category:           String(r.Category    || ''),
    description:        String(r.Description || ''),
    defaultLaborMinutes: Number(r.DefaultLaborMinutes) || 0,
    defaultSalePrice:   Number(r.DefaultSalePrice)    || 0,
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
    purchaseId:    String(r.PurchaseID   || ''),
    supplyId:      String(r.SupplyID     || ''),
    purchaseDate:  formatDate(r.PurchaseDate),
    supplier:      String(r.Supplier     || ''),
    productName:   String(r.ProductName  || ''),
    quantityBought: Number(r.QuantityBought) || 0,
    totalPaid:     Number(r.TotalPaid)   || 0,
    receiptUrl:    String(r.ReceiptURL   || ''),
    notes:         String(r.Notes        || ''),
    costPerUnit:   round2((Number(r.TotalPaid) || 0) / (Number(r.QuantityBought) || 1)),
  }));

  if (params && params.supplyId) rows = rows.filter(r => r.supplyId === params.supplyId);
  return rows;
}

function getProductMaterials(params) {
  let rows = sheetToObjects(getSheet(SHEETS.PRODUCT_MATERIALS)).map(r => ({
    productId:   String(r.ProductID   || ''),
    supplyId:    String(r.SupplyID    || ''),
    amountUsed:  Number(r.AmountUsed) || 0,
    notes:       String(r.Notes       || ''),
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

  if (params && params.status)    items = items.filter(i => i.status.toLowerCase()    === params.status.toLowerCase());
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
    date:      formatDate(r.Date),
    itemId:    String(r.ItemID    || ''),
    salePrice: Number(r.SalePrice) || 0,
    channel:   String(r.Channel   || ''),
    notes:     String(r.Notes     || ''),
  }));

  if (params && params.itemId)  sales = sales.filter(s => s.itemId  === params.itemId);
  if (params && params.channel) sales = sales.filter(s => s.channel.toLowerCase() === params.channel.toLowerCase());
  return sales;
}

// ─── READ: Calculated cost for a product (from recipe + purchase prices) ──────

function getItemMaterialCost(productId) {
  const materials = getProductMaterials({ productId });
  const purchases = getPurchases({});

  // Group purchases by supplyId → weighted average cost per unit
  const avgCost = {};
  const bySupply = {};
  purchases.forEach(p => {
    if (!bySupply[p.supplyId]) bySupply[p.supplyId] = { totalPaid: 0, totalQty: 0 };
    bySupply[p.supplyId].totalPaid += p.totalPaid;
    bySupply[p.supplyId].totalQty  += p.quantityBought;
  });
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

  const total = round2(lineItems.reduce((sum, l) => sum + l.lineCost, 0));
  return { productId, total, lineItems };
}

// ─── READ: Supply levels (purchased − used across all inventory) ──────────────

function getSupplyLevels() {
  const supplies  = getSupplies();
  const purchases = getPurchases({});
  const materials = getProductMaterials({});
  const inventory = getInventory({});

  // Count how many of each product have been made (all statuses — made = consumed supplies)
  const itemsMadeByProduct = {};
  inventory.forEach(item => {
    itemsMadeByProduct[item.productId] = (itemsMadeByProduct[item.productId] || 0) + 1;
  });

  // Total purchased per supply
  const totalPurchased = {};
  const totalPaid      = {};
  purchases.forEach(p => {
    totalPurchased[p.supplyId] = (totalPurchased[p.supplyId] || 0) + p.quantityBought;
    totalPaid[p.supplyId]      = (totalPaid[p.supplyId]      || 0) + p.totalPaid;
  });

  // Total used per supply (recipe amount × items made per product)
  const totalUsed = {};
  materials.forEach(m => {
    const made = itemsMadeByProduct[m.productId] || 0;
    totalUsed[m.supplyId] = (totalUsed[m.supplyId] || 0) + (m.amountUsed * made);
  });

  return supplies.map(s => {
    const purchased  = round2(totalPurchased[s.supplyId] || 0);
    const used       = round2(totalUsed[s.supplyId]      || 0);
    const remaining  = round2(purchased - used);
    const paid       = round2(totalPaid[s.supplyId]      || 0);
    const avgCostPer = purchased > 0 ? round2(paid / purchased) : 0;

    return {
      supplyId:    s.supplyId,
      supplyName:  s.supplyName,
      category:    s.category,
      unit:        s.unit,
      purchased,
      used,
      remaining,
      totalPaid:   paid,
      avgCostPerUnit: avgCostPer,
      low: remaining < (purchased * 0.20), // flag when < 20% left
    };
  });
}

// ─── READ: Full item detail ───────────────────────────────────────────────────

function getItem(itemId) {
  if (!itemId) return { error: 'itemId required' };

  const inventory = sheetToObjects(getSheet(SHEETS.INVENTORY));
  const row = inventory.find(r => String(r.ItemID) === itemId);
  if (!row) return { error: 'Item not found: ' + itemId };

  const labor      = getLabor(itemId);
  const matCost    = getItemMaterialCost(String(row.ProductID || ''));
  const saleList   = getSales({ itemId });
  const saleRecord = saleList[0] || null;
  const salePrice  = saleRecord ? saleRecord.salePrice : 0;
  const totalCost  = round2(matCost.total + labor.laborCost);
  const profit     = round2(salePrice - totalCost);

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
    financials:  { materialCost: matCost.total, laborCost: labor.laborCost, totalCost, salePrice, profit },
  };
}

// ─── READ: Dashboard ──────────────────────────────────────────────────────────

function getDashboard() {
  const sales     = getSales({});
  const inventory = getInventory({});
  const products  = getProducts();

  const revenue = round2(sales.reduce((sum, s) => sum + s.salePrice, 0));
  const soldIds = new Set(sales.map(s => s.itemId));

  // Total cost for sold items
  let totalExpenses = 0;
  soldIds.forEach(itemId => {
    const item = inventory.find(i => i.itemId === itemId);
    if (!item) return;
    const mat   = getItemMaterialCost(item.productId);
    const labor = getLabor(itemId);
    totalExpenses += mat.total + labor.laborCost;
  });
  totalExpenses = round2(totalExpenses);

  const profit       = round2(revenue - totalExpenses);
  const savings      = round2(profit * 0.50);
  const reinvestment = round2(profit * 0.25);
  const spending     = round2(profit * 0.25);

  // Best seller by revenue
  const revenueByProduct = {};
  sales.forEach(s => {
    const item = inventory.find(i => i.itemId === s.itemId);
    const pid  = item ? item.productId : 'Unknown';
    revenueByProduct[pid] = (revenueByProduct[pid] || 0) + s.salePrice;
  });
  const bestPid     = Object.keys(revenueByProduct).sort((a, b) => revenueByProduct[b] - revenueByProduct[a])[0] || '—';
  const bestProduct = products.find(p => p.productId === bestPid);
  const bestSeller  = bestProduct ? bestProduct.productName : bestPid;

  const supplyLevels = getSupplyLevels();
  const lowSupplies  = supplyLevels.filter(s => s.low).map(s => s.supplyName);

  const metrics = {
    revenue,
    expenses:            totalExpenses,
    profit,
    inventoryCount:      inventory.filter(i => i.status === 'Available').length,
    itemsSold:           soldIds.size,
    bestSeller,
    savingsBalance:      savings,
    reinvestmentBalance: reinvestment,
    spendingBalance:     spending,
    lowSupplies,
    generatedAt:         new Date().toISOString(),
  };

  try { writeDashboardSheet(metrics, supplyLevels); } catch (_) {}
  return metrics;
}

function writeDashboardSheet(metrics, supplyLevels) {
  const sheet = getSheet(SHEETS.DASHBOARD);
  sheet.clearContents();

  // Summary metrics
  const summaryRows = [
    ['── Summary ──', ''],
    ['Revenue',             '$' + metrics.revenue],
    ['Expenses',            '$' + metrics.expenses],
    ['Profit',              '$' + metrics.profit],
    ['Items in Stock',      metrics.inventoryCount],
    ['Items Sold',          metrics.itemsSold],
    ['Best Seller',         metrics.bestSeller],
    [''],
    ['── Savings Buckets ──', ''],
    ['Savings (50%)',        '$' + metrics.savingsBalance],
    ['Reinvestment (25%)',   '$' + metrics.reinvestmentBalance],
    ['Spending (25%)',       '$' + metrics.spendingBalance],
    [''],
    ['── Supply Levels ──', ''],
  ];

  supplyLevels.forEach(s => {
    summaryRows.push([s.supplyName + ' (' + s.unit + ')', s.remaining + ' remaining' + (s.low ? ' ⚠️ LOW' : '')]);
  });

  summaryRows.push(['']);
  summaryRows.push(['Last Updated', metrics.generatedAt]);

  sheet.getRange(1, 1, summaryRows.length, 2).setValues(summaryRows);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 200);
}

// ─── WRITE endpoints ──────────────────────────────────────────────────────────

function addPurchase(data) {
  const sheet = getSheet(SHEETS.PURCHASES);
  const costPerUnit = round2((Number(data.totalPaid) || 0) / (Number(data.quantityBought) || 1));
  sheet.appendRow([
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
  getSheet(SHEETS.PRODUCT_MATERIALS).appendRow([
    data.productId,
    data.supplyId,
    Number(data.amountUsed) || 0,
    data.notes || '',
  ]);
  return { success: true };
}

function addInventoryItem(data) {
  getSheet(SHEETS.INVENTORY).appendRow([
    data.itemId,
    data.productId,
    data.status      || 'Available',
    data.color       || '',
    data.dateCreated || new Date().toISOString().split('T')[0],
    data.photoUrl    || '',
    data.story       || '',
  ]);
  return { success: true, itemId: data.itemId };
}

function addSale(data) {
  getSheet(SHEETS.SALES).appendRow([
    data.date      || new Date().toISOString().split('T')[0],
    data.itemId,
    Number(data.salePrice) || 0,
    data.channel   || '',
    data.notes     || '',
  ]);
  try { updateItemStatus(data.itemId, 'Sold'); } catch (_) {}
  return { success: true, itemId: data.itemId };
}

function addLabor(data) {
  getSheet(SHEETS.LABOR).appendRow([
    data.itemId,
    Number(data.minutesWorked) || 0,
    Number(data.hourlyRate)    || 0,
  ]);
  return { success: true };
}

function updateItemStatus(itemId, status) {
  const sheet  = getSheet(SHEETS.INVENTORY);
  const data   = sheet.getDataRange().getValues();
  const idCol  = data[0].indexOf('ItemID');
  const stCol  = data[0].indexOf('Status');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(itemId)) {
      sheet.getRange(i + 1, stCol + 1).setValue(status);
      return { success: true };
    }
  }
  return { error: 'Item not found: ' + itemId };
}

// ─── One-time SETUP ───────────────────────────────────────────────────────────

function setupSheets() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const log = [];

  const schema = {
    [SHEETS.PRODUCTS]:          ['ProductID','ProductName','Category','Description','DefaultLaborMinutes','DefaultSalePrice'],
    [SHEETS.SUPPLIES]:          ['SupplyID','SupplyName','Category','Unit','Notes'],
    [SHEETS.PURCHASES]:         ['PurchaseID','SupplyID','PurchaseDate','Supplier','ProductName','QuantityBought','TotalPaid','CostPerUnit','ReceiptURL','Notes'],
    [SHEETS.PRODUCT_MATERIALS]: ['ProductID','SupplyID','AmountUsed','Notes'],
    [SHEETS.INVENTORY]:         ['ItemID','ProductID','Status','Color','DateCreated','PhotoURL','Story'],
    [SHEETS.LABOR]:             ['ItemID','MinutesWorked','HourlyRate'],
    [SHEETS.SALES]:             ['Date','ItemID','SalePrice','Channel','Notes'],
    [SHEETS.DASHBOARD]:         ['Metric','Value'],
  };

  Object.entries(schema).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      log.push('Created: ' + name);
    } else {
      log.push('Exists:  ' + name);
    }
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

// ─── Sample Data ──────────────────────────────────────────────────────────────
// All IDs are prefixed with DEMO- and names with [DEMO] so placeholder data
// is immediately obvious and won't be confused with real entries.

function seedSampleData() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const log = [];

  // ── Products ────────────────────────────────────────────────────────────────
  const prodSheet = ss.getSheetByName(SHEETS.PRODUCTS);
  if (prodSheet.getLastRow() < 2) {
    prodSheet.getRange(2, 1, 9, 6).setValues([
      ['DEMO-BC-001','[DEMO] BearCat Keychain',        'Keychain',  '[DEMO] Gummy bear cat resin keychain with chain',           45,  12.00],
      ['DEMO-HK-001','[DEMO] Hello Kitty Resin',       'Resin Art', '[DEMO] Hello Kitty character cast in clear resin',          60,  18.00],
      ['DEMO-CB-001','[DEMO] Chicken Banana',          'Figurine',  '[DEMO] Chicken riding a banana figurine',                   30,  10.00],
      ['DEMO-DR-001','[DEMO] 3D Printed Dragon',       '3D Print',  '[DEMO] Articulated dragon in recycled filament',            20,   8.00],
      ['DEMO-RL-001','[DEMO] Resin Lemon',             'Resin Art', '[DEMO] Translucent lemon slice resin art piece',            75,  22.00],
      ['DEMO-SH-001','[DEMO] Trail Shoe Art',          'Resin Art', '[DEMO] Trail running shoe themed resin casting',            90,  28.00],
      ['DEMO-RF-001','[DEMO] Recycled Filament Cast',  '3D Print',  '[DEMO] Art piece made from recycled 3D printer filament',   15,   6.00],
      ['DEMO-BR-001','[DEMO] Handmade Bracelet',       'Jewelry',   '[DEMO] Beaded or charm bracelet, one-of-a-kind',            30,   8.00],
      ['DEMO-RK-001','[DEMO] Hand Painted Rock',       'Painted',   '[DEMO] Natural river rock with hand painted design',        20,   5.00],
    ]);
    log.push('Seeded Products: 9 rows');
  } else { log.push('Products skipped'); }

  // ── Supplies ────────────────────────────────────────────────────────────────
  const supSheet = ss.getSheetByName(SHEETS.SUPPLIES);
  if (supSheet.getLastRow() < 2) {
    supSheet.getRange(2, 1, 10, 5).setValues([
      ['DEMO-RES-001','[DEMO] Clear Casting Resin',      'Resin',     'oz',   '[DEMO] 2-part A+B clear casting resin'],
      ['DEMO-RES-002','[DEMO] UV Resin',                 'Resin',     'oz',   '[DEMO] Single-part UV-cure resin'],
      ['DEMO-FIL-001','[DEMO] PLA Filament — Black',     'Filament',  'g',    '[DEMO] Standard black PLA for dragons & casts'],
      ['DEMO-FIL-002','[DEMO] PLA Filament — Blue',      'Filament',  'g',    '[DEMO] Ocean blue PLA'],
      ['DEMO-PAI-001','[DEMO] Acrylic Paint Set',        'Paint',     'oz',   '[DEMO] Multi-color acrylic set'],
      ['DEMO-HDW-001','[DEMO] Keychain Hook + Ring Set', 'Hardware',  'each', '[DEMO] Lobster clasp + jump ring set, silver'],
      ['DEMO-HDW-002','[DEMO] LED Light Base',           'Hardware',  'each', '[DEMO] USB LED display stand, 7-color'],
      ['DEMO-PKG-001','[DEMO] Clear Poly Bag 3×4"',      'Packaging', 'each', '[DEMO] Resealable display bags'],
      ['DEMO-PKG-002','[DEMO] Kraft Hang Tag',           'Packaging', 'each', '[DEMO] Pre-punched tags for price/story label'],
      ['DEMO-GIT-001','[DEMO] Gold Holographic Glitter', 'Other',     'oz',   '[DEMO] Fine holographic gold glitter'],
    ]);
    log.push('Seeded Supplies: 10 rows');
  } else { log.push('Supplies skipped'); }

  // ── Purchases ───────────────────────────────────────────────────────────────
  const purSheet = ss.getSheetByName(SHEETS.PURCHASES);
  if (purSheet.getLastRow() < 2) {
    purSheet.getRange(2, 1, 11, 10).setValues([
      ['DEMO-PUR-001','DEMO-RES-001','2025-02-15','Amazon',       'Alumilite Clear 32oz Kit',     32, 24.99, 0.781,'','[DEMO] First resin purchase'],
      ['DEMO-PUR-002','DEMO-HDW-001','2025-02-15','Hobby Lobby',  'Lobster Clasp + Ring 50pk',    50,  6.99, 0.140,'','[DEMO] 50-pack silver hardware'],
      ['DEMO-PUR-003','DEMO-FIL-001','2025-02-20','Amazon',       'Hatchbox PLA Black 1kg',     1000, 19.99, 0.020,'','[DEMO] 1kg spool'],
      ['DEMO-PUR-004','DEMO-PAI-001','2025-02-20',"Michael's",    'Apple Barrel Acrylic 12-pack',  12,  8.99, 0.749,'','[DEMO] 12 × 1oz bottles'],
      ['DEMO-PUR-005','DEMO-PKG-001','2025-02-20','Amazon',       'Clear Poly Bag 3x4 100pk',    100,  5.99, 0.060,'','[DEMO] 100-pack bags'],
      ['DEMO-PUR-006','DEMO-PKG-002','2025-02-20','Amazon',       'Kraft Hang Tags 50pk',         50,  3.99, 0.080,'','[DEMO] 50-pack tags'],
      ['DEMO-PUR-007','DEMO-GIT-001','2025-03-01','Hobby Lobby',  'Gold Holo Glitter 2oz',         2,  3.49, 1.745,'','[DEMO] Fine gold holographic'],
      ['DEMO-PUR-008','DEMO-HDW-002','2025-03-05','Amazon',       'LED Display Base 5pk',          5,  9.99, 1.998,'','[DEMO] 7-color USB bases'],
      ['DEMO-PUR-009','DEMO-FIL-002','2025-03-10','Amazon',       'Hatchbox PLA Blue 1kg',      1000, 21.99, 0.022,'','[DEMO] Ocean blue spool'],
      ['DEMO-PUR-010','DEMO-RES-001','2025-05-01','Amazon',       'Alumilite Clear 32oz Kit',     32, 26.99, 0.844,'','[DEMO] Reorder — price went up'],
      ['DEMO-PUR-011','DEMO-RES-002','2025-04-10','Amazon',       "LET'S RESIN UV Resin 17oz",    17, 18.99, 1.117,'','[DEMO] UV resin for small details'],
    ]);
    log.push('Seeded Purchases: 11 rows');
  } else { log.push('Purchases skipped'); }

  // ── ProductMaterials (recipes) ───────────────────────────────────────────────
  const matSheet = ss.getSheetByName(SHEETS.PRODUCT_MATERIALS);
  if (matSheet.getLastRow() < 2) {
    matSheet.getRange(2, 1, 26, 4).setValues([
      // DEMO-BC-001 BearCat Keychain
      ['DEMO-BC-001','DEMO-RES-001', 0.15, '[DEMO] 0.15 oz measured by weight before/after pour'],
      ['DEMO-BC-001','DEMO-HDW-001', 1,    '[DEMO] 1 hook + ring per keychain'],
      ['DEMO-BC-001','DEMO-PKG-001', 1,    '[DEMO] 1 bag per item'],
      ['DEMO-BC-001','DEMO-PKG-002', 1,    '[DEMO] 1 hang tag per item'],
      // DEMO-HK-001 Hello Kitty
      ['DEMO-HK-001','DEMO-RES-001', 0.20, '[DEMO] 0.20 oz — slightly larger mold than BC'],
      ['DEMO-HK-001','DEMO-GIT-001', 0.01, '[DEMO] Pinch of glitter per pour'],
      ['DEMO-HK-001','DEMO-PKG-001', 1,    '[DEMO] 1 bag per item'],
      ['DEMO-HK-001','DEMO-PKG-002', 1,    '[DEMO] 1 hang tag per item'],
      // DEMO-CB-001 Chicken Banana
      ['DEMO-CB-001','DEMO-RES-001', 0.18, '[DEMO] 0.18 oz for figurine mold'],
      ['DEMO-CB-001','DEMO-PAI-001', 0.02, '[DEMO] Small amount for detail painting'],
      ['DEMO-CB-001','DEMO-PKG-001', 1,    '[DEMO] 1 bag per item'],
      // DEMO-DR-001 Dragon
      ['DEMO-DR-001','DEMO-FIL-001', 8,    '[DEMO] ~8g per slicer estimate at 15% infill'],
      ['DEMO-DR-001','DEMO-PKG-001', 1,    '[DEMO] 1 bag per item'],
      ['DEMO-DR-001','DEMO-PKG-002', 1,    '[DEMO] 1 hang tag per item'],
      // DEMO-RL-001 Resin Lemon
      ['DEMO-RL-001','DEMO-RES-001', 0.25, '[DEMO] 0.25 oz for lemon slice mold'],
      ['DEMO-RL-001','DEMO-GIT-001', 0.01, '[DEMO] Gold glitter inside each slice'],
      ['DEMO-RL-001','DEMO-PKG-001', 1,    '[DEMO] 1 bag per item'],
      ['DEMO-RL-001','DEMO-PKG-002', 1,    '[DEMO] 1 hang tag per item'],
      // DEMO-SH-001 Trail Shoe
      ['DEMO-SH-001','DEMO-RES-001', 0.30, '[DEMO] 0.30 oz for shoe mold — largest pour'],
      ['DEMO-SH-001','DEMO-PAI-001', 0.03, '[DEMO] Detail painting on sole and laces'],
      ['DEMO-SH-001','DEMO-GIT-001', 0.01, '[DEMO] Glitter accent'],
      ['DEMO-SH-001','DEMO-PKG-001', 1,    '[DEMO] 1 bag per item'],
      ['DEMO-SH-001','DEMO-PKG-002', 1,    '[DEMO] 1 hang tag per item'],
      // DEMO-RK-001 Painted Rock
      ['DEMO-RK-001','DEMO-PAI-001', 0.05, '[DEMO] ~0.05 oz paint per rock'],
      ['DEMO-RK-001','DEMO-PKG-001', 1,    '[DEMO] 1 bag per item'],
      ['DEMO-RK-001','DEMO-PKG-002', 1,    '[DEMO] 1 hang tag per item'],
    ]);
    log.push('Seeded ProductMaterials: 26 rows');
  } else { log.push('ProductMaterials skipped'); }

  // ── Inventory ────────────────────────────────────────────────────────────────
  const invSheet = ss.getSheetByName(SHEETS.INVENTORY);
  if (invSheet.getLastRow() < 2) {
    invSheet.getRange(2, 1, 12, 7).setValues([
      ['DEMO-HK-001-001','DEMO-HK-001','Sold',      'Pink/White', '2025-03-01','','[DEMO] My first Hello Kitty! I made it with pink glitter.'],
      ['DEMO-HK-001-002','DEMO-HK-001','Available', 'Blue/Silver','2025-03-15','','[DEMO] Sparkly blue with silver stars inside.'],
      ['DEMO-HK-001-003','DEMO-HK-001','Available', 'Rainbow',    '2025-04-01','','[DEMO] Rainbow swirl — my favorite one yet!'],
      ['DEMO-BC-001-001','DEMO-BC-001','Sold',      'Clear/Green','2025-02-20','','[DEMO] Green gummy cat, first keychain I ever sold!'],
      ['DEMO-BC-001-002','DEMO-BC-001','Available', 'Purple',     '2025-04-10','','[DEMO] Purple sparkle BearCat.'],
      ['DEMO-BC-001-003','DEMO-BC-001','Available', 'Orange',     '2025-04-20','','[DEMO] Halloween orange BearCat.'],
      ['DEMO-DR-001-001','DEMO-DR-001','Available', 'Blue',       '2025-05-01','','[DEMO] Articulated dragon in ocean blue filament.'],
      ['DEMO-DR-001-002','DEMO-DR-001','Sold',      'Red',        '2025-04-05','','[DEMO] Red dragon sold at the school fair!'],
      ['DEMO-RL-001-001','DEMO-RL-001','Available', 'Yellow',     '2025-05-10','','[DEMO] Classic lemon slice with gold flake.'],
      ['DEMO-RK-001-001','DEMO-RK-001','Sold',      'Multi',      '2025-03-10','','[DEMO] Ladybug rock, sold to grandma.'],
      ['DEMO-RK-001-002','DEMO-RK-001','Available', 'Multi',      '2025-05-20','','[DEMO] Rainbow paw print rock.'],
      ['DEMO-BR-001-001','DEMO-BR-001','Sold',      'Pink/Gold',  '2025-04-15','','[DEMO] Pink beaded bracelet with gold butterfly charm.'],
    ]);
    log.push('Seeded Inventory: 12 rows');
  } else { log.push('Inventory skipped'); }

  // ── Labor ────────────────────────────────────────────────────────────────────
  const laborSheet = ss.getSheetByName(SHEETS.LABOR);
  if (laborSheet.getLastRow() < 2) {
    laborSheet.getRange(2, 1, 12, 3).setValues([
      ['DEMO-HK-001-001', 60, 5.00],
      ['DEMO-HK-001-002', 45, 5.00],
      ['DEMO-HK-001-003', 75, 5.00],
      ['DEMO-BC-001-001', 40, 5.00],
      ['DEMO-BC-001-002', 40, 5.00],
      ['DEMO-BC-001-003', 40, 5.00],
      ['DEMO-DR-001-001', 20, 5.00],
      ['DEMO-DR-001-002', 20, 5.00],
      ['DEMO-RL-001-001', 90, 5.00],
      ['DEMO-RK-001-001', 20, 5.00],
      ['DEMO-RK-001-002', 20, 5.00],
      ['DEMO-BR-001-001', 30, 5.00],
    ]);
    log.push('Seeded Labor: 12 rows');
  } else { log.push('Labor skipped'); }

  // ── Sales ────────────────────────────────────────────────────────────────────
  const salesSheet = ss.getSheetByName(SHEETS.SALES);
  if (salesSheet.getLastRow() < 2) {
    salesSheet.getRange(2, 1, 5, 5).setValues([
      ['2025-03-05','DEMO-HK-001-001',18.00,'In Person','[DEMO] Sold at spring craft fair'],
      ['2025-03-12','DEMO-BC-001-001',12.00,'In Person','[DEMO] School fundraiser'],
      ['2025-04-07','DEMO-DR-001-002', 8.00,'Etsy',     '[DEMO] First Etsy sale!'],
      ['2025-04-18','DEMO-RK-001-001', 5.00,'In Person','[DEMO] Sold to grandma'],
      ['2025-04-22','DEMO-BR-001-001', 8.00,'In Person','[DEMO] Mom\'s coworker'],
    ]);
    log.push('Seeded Sales: 5 rows');
  } else { log.push('Sales skipped'); }

  return { success: true, log };
}

// ─── Run this from the script editor for full first-time setup ────────────────

function firstTimeSetup() {
  setupSheets();
  seedSampleData();
  Logger.log('✅ Setup complete. Check the spreadsheet.');
}
