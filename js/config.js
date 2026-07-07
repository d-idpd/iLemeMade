// iLemeMade — central configuration
// After deploying the Apps Script web app, paste the URL below.

export const Config = {
  // Paste your Apps Script deployment URL here after deploying.
  // Example: 'https://script.google.com/macros/s/AKfy.../exec'
  SHEETS_API_URL: '',

  // Public CSV for the Products sheet — anyone can read, no auth needed
  PRODUCTS_CSV_URL: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQivJ5o_YEdKPRlPhpmod_UuJbe6UTTi0eRazVGNxefBFRp0ZawAdCYy1qPHVFPTvD2G6GfwPnEKOct/pub?gid=268917907&single=true&output=csv',

  // Profit allocation percentages (must sum to 1.0)
  SAVINGS_RATE:      0.50,
  REINVESTMENT_RATE: 0.25,
  SPENDING_RATE:     0.25,

  // Default hourly rate for labor calculations
  DEFAULT_HOURLY_RATE: 5.00,

  // Inventory statuses
  STATUS: {
    AVAILABLE: 'Available',
    SOLD:      'Sold',
    RESERVED:  'Reserved',
    DONATED:   'Donated',
    DAMAGED:   'Damaged',
  },

  // Sales channels
  CHANNELS: ['In Person', 'Etsy', 'Instagram', 'Gift'],

  // Product category colors (matches brand palette)
  CATEGORY_COLORS: {
    'Resin Art': '#ff9a3c',
    'Keychain':  '#b388ff',
    'Figurine':  '#ff7eb3',
    '3D Print':  '#4da6ff',
    'Jewelry':   '#ffd93d',
    'Painted':   '#6bcb77',
  },
};
