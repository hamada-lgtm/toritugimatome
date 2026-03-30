// dataStore.js - データ保持・正規化・フィルタリング

const DataStore = {
  _campaignSheets: [],  // [{sheetName: "2601", data: [...]}]
  _orders: [],          // [{受注日, 取引先名, アポイント種別, 計上金額}]
  _activeFilters: {
    sheets: []           // 選択中のシート名（空=全期間）
  },
  _listeners: [],

  // --- キャンペーンデータ ---
  addCampaignSheet(sheetName, data) {
    // 同じシート名が既にあれば上書き
    const idx = this._campaignSheets.findIndex(s => s.sheetName === sheetName);
    const normalizedData = data.map(row => ({
      ...row,
      _sheetName: sheetName,
      _campaignCost: this.normalizeAmount(row['キャンペーン単価'])
    }));
    if (idx >= 0) {
      this._campaignSheets[idx] = { sheetName, data: normalizedData };
    } else {
      this._campaignSheets.push({ sheetName, data: normalizedData });
    }
    // デフォルトで全シートを選択状態にする
    this._activeFilters.sheets = this._campaignSheets.map(s => s.sheetName);
    this._notify();
  },

  getCampaignSheetNames() {
    return this._campaignSheets.map(s => s.sheetName);
  },

  getAllCampaignData() {
    return this._campaignSheets.flatMap(s => s.data);
  },

  getFilteredCampaignData() {
    const activeSheets = this._activeFilters.sheets;
    if (activeSheets.length === 0) {
      return this.getAllCampaignData();
    }
    return this._campaignSheets
      .filter(s => activeSheets.includes(s.sheetName))
      .flatMap(s => s.data);
  },

  // --- 受注データ ---
  setOrders(data) {
    this._orders = data.map(row => ({
      ...row,
      _orderAmount: this.normalizeAmount(row['計上金額'])
    }));
    this._notify();
  },

  getOrders() {
    return this._orders;
  },

  hasOrders() {
    return this._orders.length > 0;
  },

  hasCampaignData() {
    return this._campaignSheets.length > 0;
  },

  // --- フィルター ---
  setSheetFilter(sheetNames) {
    this._activeFilters.sheets = [...sheetNames];
    this._notify();
  },

  // --- 金額正規化 ---
  normalizeAmount(value) {
    if (typeof value === 'number') return value;
    if (typeof value === 'string') {
      const cleaned = value.replace(/[,，￥¥\\$\s　円]/g, '');
      const num = parseFloat(cleaned);
      return isNaN(num) ? 0 : num;
    }
    return 0;
  },

  // --- イベント通知 ---
  onChange(callback) {
    this._listeners.push(callback);
  },

  _notify() {
    this._listeners.forEach(cb => cb());
  }
};
