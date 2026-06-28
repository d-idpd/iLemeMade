// iLemeMade — Google Sheets API client v2

import { Config } from './config.js';

async function apiFetch(params = {}) {
  const url = Config.SHEETS_API_URL;
  if (!url) throw new Error('SHEETS_API_URL not set in js/config.js. Deploy the Apps Script web app first.');
  const res  = await fetch(`${url}?${new URLSearchParams(params)}`);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const json = await res.json();
  if (json.error) throw new Error(`API: ${json.error}`);
  return json;
}

async function apiPost(body = {}) {
  const url = Config.SHEETS_API_URL;
  if (!url) throw new Error('SHEETS_API_URL not set in js/config.js.');
  const res  = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const json = await res.json();
  if (json.error) throw new Error(`API: ${json.error}`);
  return json;
}

// ─── Read API ─────────────────────────────────────────────────────────────────

export const SheetsAPI = {

  ping: ()                     => apiFetch({ action: 'ping' }),

  // Catalog
  getProducts: ()              => apiFetch({ action: 'getProducts' }),
  getSupplies: ()              => apiFetch({ action: 'getSupplies' }),

  /**
   * Supply levels — purchased, used, remaining, low-stock flag.
   * Deduction is automatic: remaining = purchased − (recipe × items made).
   */
  getSupplyLevels: ()          => apiFetch({ action: 'getSupplyLevels' }),

  /**
   * Purchase receipts.
   * @param {{ supplyId?: string }} filters
   */
  getPurchases: (filters = {}) => apiFetch({ action: 'getPurchases', ...filters }),

  /**
   * Product recipes — how much of each supply per product.
   * @param {{ productId?: string }} filters
   */
  getProductMaterials: (f = {})=> apiFetch({ action: 'getProductMaterials', ...f }),

  // Inventory
  getInventory: (f = {})       => apiFetch({ action: 'getInventory', ...f }),
  getItem: (itemId)            => apiFetch({ action: 'getItem', itemId }),
  getLabor: (itemId)           => apiFetch({ action: 'getLabor', itemId }),

  // Sales
  getSales: (f = {})           => apiFetch({ action: 'getSales', ...f }),

  // Dashboard (also refreshes the Dashboard sheet)
  getDashboard: ()             => apiFetch({ action: 'getDashboard' }),

  // ─── Write ─────────────────────────────────────────────────────────────────

  /**
   * Log a supply purchase (links to receipt URL for proof).
   * Supply levels update automatically — no separate deduction needed.
   * @param {{ purchaseId, supplyId, purchaseDate, supplier, productName,
   *           quantityBought, totalPaid, receiptUrl, notes }} data
   */
  addPurchase: (data)                => apiPost({ action: 'addPurchase', data }),

  /** Add a recipe row (how much of a supply one product uses). */
  addProductMaterial: (data)         => apiPost({ action: 'addProductMaterial', data }),

  /**
   * Add a new physical item to inventory.
   * Supplies are automatically "consumed" in getSupplyLevels — no extra step.
   * @param {{ itemId, productId, status, color, dateCreated, photoUrl, story }} data
   */
  addInventoryItem: (data)           => apiPost({ action: 'addInventoryItem', data }),

  /**
   * Record a sale. Automatically marks the item as Sold.
   * @param {{ date, itemId, salePrice, channel, notes }} data
   */
  addSale: (data)                    => apiPost({ action: 'addSale', data }),

  addLabor: (data)                   => apiPost({ action: 'addLabor', data }),
  updateItemStatus: (itemId, status) => apiPost({ action: 'updateItemStatus', itemId, status }),
};

// ─── Client-side calculation helpers ─────────────────────────────────────────

export const Calc = {

  /** Weighted average cost per unit across multiple purchases of the same supply. */
  weightedAvgCost(purchases) {
    const totalPaid = purchases.reduce((s, p) => s + p.totalPaid,      0);
    const totalQty  = purchases.reduce((s, p) => s + p.quantityBought, 0);
    return totalQty > 0 ? round2(totalPaid / totalQty) : 0;
  },

  /** Material cost for one item, given its product's recipe and supply purchase prices. */
  materialCost(productMaterials, allPurchases) {
    const bySupply = groupBy(allPurchases, p => p.supplyId);
    return round2(productMaterials.reduce((sum, m) => {
      const purchases = bySupply[m.supplyId] || [];
      const unitCost  = this.weightedAvgCost(purchases);
      return sum + (m.amountUsed * unitCost);
    }, 0));
  },

  laborCost(minutesWorked, hourlyRate) {
    return round2((minutesWorked / 60) * hourlyRate);
  },

  profit(salePrice, materialCost, laborCost) {
    return round2(salePrice - materialCost - laborCost);
  },

  savingsAllocation(profit) {
    return {
      savings:      round2(profit * Config.SAVINGS_RATE),
      reinvestment: round2(profit * Config.REINVESTMENT_RATE),
      spending:     round2(profit * Config.SPENDING_RATE),
    };
  },

  formatCurrency(n) { return '$' + Number(n).toFixed(2); },
};

function round2(n) { return Math.round(n * 100) / 100; }
function groupBy(arr, keyFn) {
  return arr.reduce((acc, item) => {
    const k = keyFn(item);
    (acc[k] = acc[k] || []).push(item);
    return acc;
  }, {});
}
