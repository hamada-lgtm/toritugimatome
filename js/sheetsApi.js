// sheetsApi.js - Google Sheets APIからデータを取得

const SheetsAPI = {
  BASE_URL: 'https://sheets.googleapis.com/v4/spreadsheets',

  // スプレッドシートID（URLから抽出済み）
  CAMPAIGN_SHEET_ID: '1tt1q2zhiyi5bY2ZRBi9B_cmO-AKlGN7rZaboOmjxhVc',
  ORDERS_SHEET_ID: '18p8EptwPJN-kwpAmV07yzamaoCsDcUxJ-xD3xvxt4bc',

  _apiKey: 'AIzaSyDxRHR5qb_Z7kfR4kJ1nc3B6VDEqfY1i1A',

  /** APIキーをセット（localStorageにも保存） */
  setApiKey(key) {
    this._apiKey = key.trim();
    localStorage.setItem('sheets_api_key', this._apiKey);
  },

  /** 保存済みAPIキーを読み込み（デフォルト値にフォールバック） */
  loadApiKey() {
    const stored = localStorage.getItem('sheets_api_key');
    if (stored) this._apiKey = stored;
    return this._apiKey;
  },

  /** スプレッドシートIDをセット */
  setCampaignSheetId(id) {
    this.CAMPAIGN_SHEET_ID = id.trim();
    localStorage.setItem('campaign_sheet_id', this.CAMPAIGN_SHEET_ID);
  },

  setOrdersSheetId(id) {
    this.ORDERS_SHEET_ID = id.trim();
    localStorage.setItem('orders_sheet_id', this.ORDERS_SHEET_ID);
  },

  /** 保存済みシートIDを読み込み（古いIDは破棄） */
  loadSheetIds() {
    const OLD_ORDER_IDS = ['118oCiVwQZHuFU4kU1YeZw7RnusX50XPw20Vc7dP-fqs', '1o2yK3timPvKqkpytVP6_cgGjHdbZlYR0OkPFt1RtPTk'];
    const storedOrderId = localStorage.getItem('orders_sheet_id');
    if (storedOrderId && OLD_ORDER_IDS.includes(storedOrderId)) {
      localStorage.removeItem('orders_sheet_id');
    }
    this.CAMPAIGN_SHEET_ID = localStorage.getItem('campaign_sheet_id') || this.CAMPAIGN_SHEET_ID;
    this.ORDERS_SHEET_ID = localStorage.getItem('orders_sheet_id') || this.ORDERS_SHEET_ID;
  },

  /** Google SheetsのURLからスプレッドシートIDを抽出 */
  extractSheetId(url) {
    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    return match ? match[1] : url;
  },

  /**
   * スプレッドシートのメタデータを取得（シート名一覧）
   */
  async getSheetNames(spreadsheetId) {
    const url = `${this.BASE_URL}/${spreadsheetId}?key=${this._apiKey}&fields=sheets.properties.title`;
    const res = await fetch(url);
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error?.message || `API Error: ${res.status}`);
    }
    const data = await res.json();
    return data.sheets.map(s => s.properties.title);
  },

  /**
   * キャンペーンスプレッドシートから月次シート名を取得
   * "2601" 等の4桁数字、"2025.12月" 等の年月パターンにマッチ
   */
  async getCampaignMonthlySheets() {
    const names = await this.getSheetNames(this.CAMPAIGN_SHEET_ID);
    return names.filter(name => {
      const n = name.trim();
      return /^\d{4}$/.test(n) || /^\d{4}\.\d{1,2}月$/.test(n);
    });
  },

  /**
   * 指定シートの全データを取得
   */
  async getSheetData(spreadsheetId, sheetName, headerRow) {
    const range = encodeURIComponent(`'${sheetName}'`);
    const url = `${this.BASE_URL}/${spreadsheetId}/values/${range}?key=${this._apiKey}&valueRenderOption=UNFORMATTED_VALUE`;
    const res = await fetch(url);
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error?.message || `API Error: ${res.status}`);
    }
    const data = await res.json();
    const rows = data.values || [];
    if (rows.length <= headerRow) return [];

    // ヘッダー行を取得し、それ以降をオブジェクト配列に変換
    const headers = rows[headerRow].map(h => String(h).trim());
    const records = [];
    for (let i = headerRow + 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;
      const obj = {};
      headers.forEach((h, idx) => {
        obj[h] = row[idx] !== undefined ? row[idx] : '';
      });
      records.push(obj);
    }
    return records;
  },

  /**
   * キャンペーンデータを全月分取得
   */
  async fetchAllCampaignData(progressCallback) {
    const monthlySheets = await this.getCampaignMonthlySheets();
    if (monthlySheets.length === 0) {
      throw new Error('月次シート（4桁数字名）が見つかりません');
    }

    const results = [];
    for (const sheetName of monthlySheets) {
      if (progressCallback) progressCallback(`${sheetName} を読み込み中...`);
      // ヘッダーは6行目 → index 5
      const data = await this.getSheetData(this.CAMPAIGN_SHEET_ID, sheetName, 5);
      if (data.length === 0) continue;

      // カラム名が改行入りの場合があるので、部分一致で標準キーにマッピング
      const headers = Object.keys(data[0]);
      const colPartner = this._findColumn(headers, '取次パートナー');
      const colCompany = this._findColumn(headers, '紹介企業');
      const colCost = this._findColumn(headers, 'キャンペーン単価');
      const colMeetingDate = this._findColumn(headers, '初回商談実施日');
      console.log(`[CP] シート「${sheetName}」カラムマッピング: パートナー=${colPartner}, 企業=${colCompany}, 単価=${colCost}, 商談日=${colMeetingDate}`);

      // 標準キーに正規化して格納
      const normalized = data.map(row => {
        const obj = { ...row };
        if (colPartner && colPartner !== '取次パートナー') obj['取次パートナー'] = row[colPartner] || '';
        if (colCompany && colCompany !== '紹介企業') obj['紹介企業'] = row[colCompany] || '';
        if (colCost && colCost !== 'キャンペーン単価') obj['キャンペーン単価'] = row[colCost] || '';
        if (colMeetingDate) obj['初回商談実施日'] = row[colMeetingDate] || '';
        return obj;
      });

      // 有効な行のみフィルター
      const validData = normalized.filter(row => {
        const partner = String(row['取次パートナー'] || '').trim();
        const company = String(row['紹介企業'] || '').trim();
        return partner !== '' || company !== '';
      });
      if (validData.length > 0) {
        results.push({ sheetName: sheetName.trim(), data: validData });
      }
    }
    return results;
  },

  /**
   * ヘッダー配列からカラム名を柔軟に検索する
   * 部分一致・全角半角括弧の違いを吸収
   */
  _findColumn(headers, ...keywords) {
    for (const kw of keywords) {
      // 完全一致
      const exact = headers.find(h => h === kw);
      if (exact) return exact;
    }
    for (const kw of keywords) {
      // 部分一致（キーワードを含むヘッダー）
      const partial = headers.find(h => h.includes(kw));
      if (partial) return partial;
    }
    // 括弧の全角半角を入れ替えて再検索
    for (const kw of keywords) {
      const alt = kw.replace(/（/g, '(').replace(/）/g, ')');
      const partial = headers.find(h => h.includes(alt));
      if (partial) return partial;
      const alt2 = kw.replace(/\(/g, '（').replace(/\)/g, '）');
      const partial2 = headers.find(h => h.includes(alt2));
      if (partial2) return partial2;
    }
    return null;
  },

  /**
   * 1つのシートからデータ取得を試みる
   */
  async _tryFetchSheet(sheetName) {
    const range = encodeURIComponent(`'${sheetName}'`);
    const url = `${this.BASE_URL}/${this.ORDERS_SHEET_ID}/values/${range}?key=${this._apiKey}&valueRenderOption=UNFORMATTED_VALUE`;
    const res = await fetch(url);
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      console.warn(`[受注] シート「${sheetName}」取得失敗:`, res.status, err.error?.message || '');
      return null;
    }
    const data = await res.json();
    console.log(`[受注] シート「${sheetName}」: ${(data.values || []).length}行取得`);
    return data.values || [];
  },

  /**
   * 受注データを取得（全シートを探索して正しいデータシートを自動検出）
   */
  async fetchOrdersDataAuto(progressCallback) {
    if (progressCallback) progressCallback('受注シート情報を取得中...');
    const sheetNames = await this.getSheetNames(this.ORDERS_SHEET_ID);
    console.log('[受注] 全シート名:', sheetNames);
    if (progressCallback) progressCallback(`シート一覧: ${sheetNames.join(', ')}`);

    // 「※ERP出力」シートのみを対象とする
    const targetSheets = sheetNames.filter(name => name.includes('ERP出力'));
    if (targetSheets.length === 0) {
      console.warn('[受注] ERP出力シートが見つかりません。全シートを探索します。');
    }
    const sheetsToSearch = targetSheets.length > 0 ? targetSheets : sheetNames;

    const allRecords = [];
    const usedSheets = [];

    for (const sheetName of sheetsToSearch) {
      if (progressCallback) progressCallback(`「${sheetName}」を確認中...`);
      const rows = await this._tryFetchSheet(sheetName);
      if (!rows || rows.length <= 1) continue;

      console.log(`[受注] シート「${sheetName}」: ${rows.length}行`);

      // ヘッダー行を自動検出: 先頭20行の中で文字列セルが3つ以上ある行
      let headerRowIdx = -1;
      for (let i = 0; i < Math.min(20, rows.length); i++) {
        const row = rows[i];
        if (!row) continue;
        const textCells = row.filter(cell => typeof cell === 'string' && cell.trim().length > 0);
        if (textCells.length >= 3) {
          headerRowIdx = i;
          break;
        }
      }
      if (headerRowIdx < 0) continue;

      const headers = rows[headerRowIdx].map(h => String(h).trim());
      console.log(`[受注] シート「${sheetName}」ヘッダー(行${headerRowIdx}):`, headers);

      // カラム名マッピング
      const colMap = {
        companyName: this._findColumn(headers, '取引先名（正式名称）', '取引先名(正式名称)', '取引先名', '企業名', '会社名'),
        orderDate: this._findColumn(headers, '受注日', '申込日', '契約日'),
        appointType: this._findColumn(headers, 'アポイント種別', 'アポイント種類', '種別', '区分'),
        amount: this._findColumn(headers, '計上金額', '受注金額', '金額', '売上')
      };

      // 企業名カラムが見つかったらこのシートのデータを追加
      if (colMap.companyName) {
        if (progressCallback) progressCallback(`「${sheetName}」からデータ取得中...`);
        console.log(`[受注] シート「${sheetName}」を採用。カラムマッピング:`, colMap);

        for (let i = headerRowIdx + 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;

          const obj = {};
          headers.forEach((h, idx) => {
            obj[h] = row[idx] !== undefined ? row[idx] : '';
          });

          // 標準キーに正規化
          obj['取引先名（正式名称）'] = obj[colMap.companyName] || '';
          if (colMap.orderDate) obj['受注日'] = obj[colMap.orderDate] || '';
          if (colMap.appointType) obj['アポイント種別'] = obj[colMap.appointType] || '';
          if (colMap.amount) obj['計上金額'] = obj[colMap.amount] || '';
          obj['_sourceSheet'] = sheetName;

          const name = String(obj[colMap.companyName] || '').trim();
          if (name !== '') allRecords.push(obj);
        }

        usedSheets.push(sheetName);
        console.log(`[受注] シート「${sheetName}」: ${allRecords.length}件（累計）`);
      }
    }

    if (allRecords.length > 0) {
      console.log(`[受注] 全シート統合: ${allRecords.length}件 (${usedSheets.join(', ')})`);
      return { records: allRecords, headers: [], colMap: {companyName: true}, rawPreview: [], headerRowIdx: 0, sheetName: usedSheets.join(' + ') };
    }

    // どのシートでも企業名カラムが見つからなかった場合
    // 全シートのヘッダー情報をまとめて返す
    const allHeaders = [];
    for (const sheetName of sheetNames) {
      const rows = await this._tryFetchSheet(sheetName);
      if (!rows || rows.length === 0) continue;
      const preview = rows.slice(0, 3).map(r => (r || []).join(' | '));
      allHeaders.push({ sheetName, preview });
    }
    return {
      records: [], headers: [], colMap: {},
      rawPreview: [], headerRowIdx: -1, sheetName: null,
      allSheetPreviews: allHeaders
    };
  }
};
