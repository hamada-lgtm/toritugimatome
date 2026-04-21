// app.js - アプリ初期化・イベント制御・処理フロー統合

const App = {
  _partnerKPIs: [],
  _tableFilters: [],
  _partnerMonthRange: null,  // null=全期間、{fromIdx, toIdx}で範囲指定
  _partnerMatrix: null,       // パートナー×月のマトリックス

  init() {
    // 保存済み設定を読み込み
    SheetsAPI.loadApiKey();
    SheetsAPI.loadSheetIds();

    // 認証は最優先でバインド（他のバインド処理でエラーが出ても動作するため）
    Auth.init(() => {
      if (SheetsAPI._apiKey) {
        this._updateSettingsInputs();
        this.fetchData();
      }
    });

    // 各種イベントバインド（個別にtry-catchして、1つ失敗しても他は動作する）
    try { this.bindFetchEvents(); } catch (e) { console.error('bindFetchEvents:', e); }
    try { this.bindFilterEvents(); } catch (e) { console.error('bindFilterEvents:', e); }
    try { this.bindTableEvents(); } catch (e) { console.error('bindTableEvents:', e); }
    try { this.bindTableFilterEvents(); } catch (e) { console.error('bindTableFilterEvents:', e); }
    try { this.bindPartnerMonthEvent(); } catch (e) { console.error('bindPartnerMonthEvent:', e); }
    try { this.bindMatrixKpiEvent(); } catch (e) { console.error('bindMatrixKpiEvent:', e); }
    try { this.bindModalEvents(); } catch (e) { console.error('bindModalEvents:', e); }
    try { this.bindExportEvents(); } catch (e) { console.error('bindExportEvents:', e); }
    try { this.bindSettingsEvents(); } catch (e) { console.error('bindSettingsEvents:', e); }
  },

  // === 設定モーダル ===
  bindSettingsEvents() {
    const settingsBtn = document.getElementById('settings-btn');
    const closeBtn = document.getElementById('settings-close-btn');
    const saveBtn = document.getElementById('settings-save-btn');
    const modal = document.getElementById('settings-modal');

    settingsBtn.addEventListener('click', () => {
      this._updateSettingsInputs();
      modal.classList.add('visible');
      document.body.style.overflow = 'hidden';
    });

    closeBtn.addEventListener('click', () => {
      modal.classList.remove('visible');
      document.body.style.overflow = '';
    });

    modal.addEventListener('click', (e) => {
      if (e.target === modal) {
        modal.classList.remove('visible');
        document.body.style.overflow = '';
      }
    });

    saveBtn.addEventListener('click', () => {
      const apiKey = document.getElementById('api-key-input').value.trim();
      const campaignUrl = document.getElementById('campaign-sheet-url').value.trim();
      const ordersUrl = document.getElementById('orders-sheet-url').value.trim();

      if (!apiKey) {
        alert('APIキーを入力してください');
        return;
      }

      SheetsAPI.setApiKey(apiKey);

      if (campaignUrl) {
        const id = SheetsAPI.extractSheetId(campaignUrl);
        SheetsAPI.setCampaignSheetId(id);
      }
      if (ordersUrl) {
        const id = SheetsAPI.extractSheetId(ordersUrl);
        SheetsAPI.setOrdersSheetId(id);
      }

      modal.classList.remove('visible');
      document.body.style.overflow = '';

      // 設定保存後に自動取得
      this.fetchData();
    });
  },

  _updateSettingsInputs() {
    document.getElementById('api-key-input').value = SheetsAPI._apiKey || '';
    document.getElementById('campaign-sheet-url').value = SheetsAPI.CAMPAIGN_SHEET_ID
      ? 'https://docs.google.com/spreadsheets/d/' + SheetsAPI.CAMPAIGN_SHEET_ID
      : '';
    document.getElementById('orders-sheet-url').value = SheetsAPI.ORDERS_SHEET_ID
      ? 'https://docs.google.com/spreadsheets/d/' + SheetsAPI.ORDERS_SHEET_ID
      : '';
  },

  // === データ取得 ===
  bindFetchEvents() {
    document.getElementById('fetch-btn').addEventListener('click', () => this.fetchData());
    document.getElementById('refresh-btn').addEventListener('click', () => this.fetchData());
  },

  async fetchData() {
    if (!SheetsAPI._apiKey) {
      this._showProgress('APIキーが未設定です。「設定」ボタンからAPIキーを登録してください。', true);
      return;
    }

    const fetchBtn = document.getElementById('fetch-btn');
    const refreshBtn = document.getElementById('refresh-btn');
    fetchBtn.disabled = true;
    refreshBtn.disabled = true;

    try {
      // キャンペーンデータ取得
      this._showProgress('キャンペーンシート情報を取得中...');
      const campaignResults = await SheetsAPI.fetchAllCampaignData((msg) => {
        this._showProgress(msg);
      });

      // DataStoreに格納
      campaignResults.forEach(result => {
        DataStore.addCampaignSheet(result.sheetName, result.data);
      });

      // 受注データ取得
      this._showProgress('受注データを取得中...');
      const ordersResult = await SheetsAPI.fetchOrdersDataAuto((msg) => {
        this._showProgress(msg);
      });
      const ordersRecords = ordersResult.records;
      DataStore.setOrders(ordersRecords);

      let statusMsg = 'データ取得完了 - キャンペーン: ' + campaignResults.length + 'シート、受注: ' + ordersRecords.length + '件';
      if (ordersResult.sheetName) {
        statusMsg += ' (シート: ' + ordersResult.sheetName + ')';
      }
      if (ordersRecords.length === 0) {
        statusMsg += '\n\n⚠ 企業名カラムが自動検出できませんでした。';
        statusMsg += '\nコンソール(F12)に詳細ログが出力されています。';
        if (ordersResult.allSheetPreviews && ordersResult.allSheetPreviews.length > 0) {
          statusMsg += '\n検索済みシート:';
          ordersResult.allSheetPreviews.forEach(function(sp) {
            statusMsg += '\n\n【' + sp.sheetName + '】';
            sp.preview.forEach(function(line, i) {
              statusMsg += '\n  行' + i + ': ' + line;
            });
          });
        } else {
          statusMsg += '\n※ どのシートからもデータを取得できませんでした。シートのアクセス権限を確認してください。';
        }
      }
      this._showProgress(statusMsg);

      // ダッシュボード更新
      this.refresh();
      refreshBtn.disabled = false;

    } catch (err) {
      this._showProgress('エラー: ' + err.message, true);
    } finally {
      fetchBtn.disabled = false;
    }
  },

  _showProgress(message, isError) {
    const el = document.getElementById('fetch-progress');
    el.classList.remove('hidden', 'error');
    if (isError) el.classList.add('error');
    el.style.whiteSpace = 'pre-wrap';
    el.textContent = message;

    // ステータスバッジ更新
    const statusEl = document.getElementById('fetch-status');
    if (isError) {
      statusEl.innerHTML = '<span class="status-badge" style="background:#fce8e6;color:#d93025">エラー</span>';
    } else if (DataStore.hasCampaignData()) {
      statusEl.innerHTML = '<span class="status-badge loaded">接続中</span>';
    }
  },

  // === フィルター ===
  _applyPeriodFilter() {
    const periods = UIRenderer._sortedPeriods;
    if (periods.length === 0) return;

    const fromIdx = parseInt(document.getElementById('filter-from').value) || 0;
    const toIdx = parseInt(document.getElementById('filter-to').value) || (periods.length - 1);

    const lo = Math.min(fromIdx, toIdx);
    const hi = Math.max(fromIdx, toIdx);

    const selected = periods.slice(lo, hi + 1).map(p => p.original);
    DataStore.setSheetFilter(selected);
    this.refresh();
  },

  bindFilterEvents() {
    document.getElementById('filter-from').addEventListener('change', () => this._applyPeriodFilter());
    document.getElementById('filter-to').addEventListener('change', () => this._applyPeriodFilter());
  },

  // === テーブルソート ===
  bindTableEvents() {
    document.querySelector('#partner-table thead').addEventListener('click', (e) => {
      const th = e.target.closest('th');
      if (!th || !th.dataset.col) return;
      UIRenderer.handleSort(th.dataset.col);
      UIRenderer.renderPartnerTable(this._partnerKPIs, this._tableFilters);
    });

    document.getElementById('partner-table-body').addEventListener('click', (e) => {
      const tr = e.target.closest('tr');
      if (!tr) return;
      const partnerName = tr.dataset.partner;
      const kpi = this._partnerKPIs.find(p => p.partnerName === partnerName);
      if (kpi) {
        UIRenderer.renderDetailModal(kpi);
        UIRenderer.showModal();
      }
    });
  },

  // === テーブルフィルター ===
  _readFilters() {
    const filters = [];
    for (let i = 1; i <= 2; i++) {
      const col = document.getElementById('filter-column-' + i).value;
      const min = parseFloat(document.getElementById('filter-min-' + i).value);
      const max = parseFloat(document.getElementById('filter-max-' + i).value);
      if (col && (!isNaN(min) || !isNaN(max))) {
        filters.push({ column: col, min: isNaN(min) ? null : min, max: isNaN(max) ? null : max });
      }
    }
    return filters;
  },

  bindTableFilterEvents() {
    document.getElementById('filter-apply-btn').addEventListener('click', () => {
      this._tableFilters = this._readFilters();
      if (this._tableFilters.length === 0) return;
      UIRenderer.renderPartnerTable(this._partnerKPIs, this._tableFilters);
    });

    document.getElementById('filter-reset-btn').addEventListener('click', () => {
      this._tableFilters = [];
      for (let i = 1; i <= 2; i++) {
        document.getElementById('filter-column-' + i).value = '';
        document.getElementById('filter-min-' + i).value = '';
        document.getElementById('filter-max-' + i).value = '';
      }
      document.getElementById('filter-count').textContent = '';
      UIRenderer.renderPartnerTable(this._partnerKPIs);
    });
  },

  // === パートナー月範囲セレクタ ===
  bindPartnerMonthEvent() {
    const fromSel = document.getElementById('partner-month-from');
    const toSel = document.getElementById('partner-month-to');
    const resetBtn = document.getElementById('partner-month-reset');
    if (!fromSel || !toSel || !resetBtn) {
      console.warn('[bindPartnerMonthEvent] 要素が見つかりません');
      return;
    }

    const update = () => {
      const fromIdx = parseInt(fromSel.value);
      const toIdx = parseInt(toSel.value);
      if (isNaN(fromIdx) || isNaN(toIdx)) {
        this._partnerMonthRange = null;
      } else {
        this._partnerMonthRange = { fromIdx, toIdx };
      }
      this.refresh();
    };

    fromSel.addEventListener('change', update);
    toSel.addEventListener('change', update);
    resetBtn.addEventListener('click', () => {
      this._partnerMonthRange = null;
      fromSel.value = '0';
      toSel.value = String(fromSel.options.length - 1);
      this.refresh();
    });
  },

  // === マトリックス KPI セレクタ ===
  bindMatrixKpiEvent() {
    const select = document.getElementById('matrix-kpi');
    if (!select) return;
    select.addEventListener('change', (e) => {
      if (this._partnerMatrix) {
        UIRenderer.renderPartnerMonthlyMatrix(this._partnerMatrix, e.target.value);
      }
    });
  },

  // === モーダル ===
  bindModalEvents() {
    document.getElementById('modal-close-btn').addEventListener('click', () => UIRenderer.hideModal());
    document.getElementById('detail-modal').addEventListener('click', (e) => {
      if (e.target === e.currentTarget) UIRenderer.hideModal();
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        UIRenderer.hideModal();
        document.getElementById('settings-modal').classList.remove('visible');
        document.body.style.overflow = '';
      }
    });
  },

  // === CSVエクスポート ===
  bindExportEvents() {
    document.getElementById('export-btn').addEventListener('click', () => {
      if (this._partnerKPIs.length === 0) return;
      this._exportCSV();
    });
  },

  _exportCSV() {
    const headers = ['パートナー', '紹介数', 'キャンペーン費用', '受注数', '受注金額', '成約フィー', '受注率', 'ROI'];
    const rows = this._partnerKPIs.map(p => [
      p.partnerName,
      p.referralCount,
      Math.round(p.campaignCost),
      p.orderCount,
      Math.round(p.orderAmount),
      Math.round(p.closingFee),
      (p.conversionRate * 100).toFixed(1) + '%',
      Math.round(p.roi) + '%'
    ]);

    const csvContent = '\uFEFF' + [headers, ...rows].map(row =>
      row.map(cell => '"' + String(cell).replace(/"/g, '""') + '"').join(',')
    ).join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'partner_kpi_' + new Date().toISOString().slice(0, 10) + '.csv';
    a.click();
    URL.revokeObjectURL(url);
  },

  // === メインリフレッシュ ===
  refresh() {
    const campaignData = DataStore.getFilteredCampaignData();
    const orders = DataStore.getOrders();
    const sheetNames = DataStore.getCampaignSheetNames();

    console.log('[refresh] campaignData:', campaignData.length, '件, orders:', orders.length, '件, sheets:', sheetNames);
    if (campaignData.length > 0) {
      console.log('[refresh] campaignDataカラム名:', Object.keys(campaignData[0]).join(' | '));
    }
    if (orders.length > 0) {
      console.log('[refresh] ordersサンプル:', Object.keys(orders[0]), orders[0]);
    }

    UIRenderer.renderImportStatus(sheetNames, DataStore.hasOrders());
    UIRenderer.renderFilterBar(sheetNames, DataStore._activeFilters);
    UIRenderer.renderPartnerMonthSelector(DataStore._activeFilters.sheets);

    if (campaignData.length === 0) {
      console.warn('[refresh] campaignDataが空のためダッシュボード非表示');
      UIRenderer.hideDashboard();
      document.getElementById('export-btn').disabled = true;
      return;
    }

    // 全体サマリー・月別推移は期間フィルター全体で算出
    const matched = MatchEngine.matchAll(campaignData, orders);
    const partnerNames = [...new Set(
      campaignData.map(r => String(r['取次パートナー'] || '').trim()).filter(n => n)
    )];
    const summaryPartnerKPIs = partnerNames.map(name =>
      KPICalculator.calcPartnerKPI(name, matched)
    );
    const summaryKPI = KPICalculator.calcSummaryKPI(summaryPartnerKPIs);

    // パートナーテーブル・チャートは範囲指定があれば月フィルター適用
    let partnerData = campaignData;
    if (this._partnerMonthRange && UIRenderer._partnerSortedPeriods) {
      const { fromIdx, toIdx } = this._partnerMonthRange;
      const lo = Math.min(fromIdx, toIdx);
      const hi = Math.max(fromIdx, toIdx);
      const rangeSheets = UIRenderer._partnerSortedPeriods.slice(lo, hi + 1).map(p => p.original);
      partnerData = DataStore._campaignSheets
        .filter(s => rangeSheets.includes(s.sheetName))
        .flatMap(s => s.data);
    }
    const partnerMatched = MatchEngine.matchAll(partnerData, orders);
    const partnerNamesForTable = [...new Set(
      partnerData.map(r => String(r['取次パートナー'] || '').trim()).filter(n => n)
    )];
    this._partnerKPIs = partnerNamesForTable.map(name =>
      KPICalculator.calcPartnerKPI(name, partnerMatched)
    );

    // 月別推移KPI算出
    const filteredSheets = DataStore._campaignSheets.filter(
      s => DataStore._activeFilters.sheets.includes(s.sheetName)
    );
    const monthlyData = KPICalculator.calcMonthlyKPIs(filteredSheets, orders);

    // パートナー×月マトリックス（パートナーテーブルと同じ範囲を使用）
    let matrixSheets = filteredSheets;
    if (this._partnerMonthRange && UIRenderer._partnerSortedPeriods) {
      const { fromIdx, toIdx } = this._partnerMonthRange;
      const lo = Math.min(fromIdx, toIdx);
      const hi = Math.max(fromIdx, toIdx);
      const rangeSheets = UIRenderer._partnerSortedPeriods.slice(lo, hi + 1).map(p => p.original);
      matrixSheets = filteredSheets.filter(s => rangeSheets.includes(s.sheetName));
    }
    this._partnerMatrix = KPICalculator.calcPartnerMonthlyMatrix(matrixSheets, orders);
    const matrixKpi = document.getElementById('matrix-kpi')
      ? document.getElementById('matrix-kpi').value : 'orderAmount';

    UIRenderer.showDashboard();
    UIRenderer.renderKPICards(summaryKPI);
    UIRenderer.renderMonthlyKPI(monthlyData);
    UIRenderer.renderPartnerTable(this._partnerKPIs, this._tableFilters);
    UIRenderer.renderPartnerMonthlyMatrix(this._partnerMatrix, matrixKpi);
    ChartManager.updateAll(this._partnerKPIs, summaryKPI);

    document.getElementById('export-btn').disabled = false;
  }
};

document.addEventListener('DOMContentLoaded', () => App.init());
