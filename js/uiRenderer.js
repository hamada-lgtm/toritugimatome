// uiRenderer.js - DOM描画（KPIカード・テーブル・フィルター・モーダル）

const UIRenderer = {
  _sortState: { column: 'orderAmount', ascending: false },

  /** 金額フォーマット */
  formatCurrency(num) {
    if (num === 0) return '¥0';
    if (Math.abs(num) >= 100000000) return '¥' + (num / 100000000).toFixed(1) + '億';
    if (Math.abs(num) >= 10000) return '¥' + (num / 10000).toFixed(1) + '万';
    return '¥' + Math.round(num).toLocaleString();
  },

  /** パーセントフォーマット */
  formatPercent(num) {
    return (num * 100).toFixed(1) + '%';
  },

  /** ROIフォーマット */
  formatROI(num) {
    return Math.round(num) + '%';
  },

  /** 読み込み状態バッジ表示 */
  renderImportStatus(sheetNames, ordersLoaded) {
    const el = document.getElementById('import-status');
    let html = '';
    sheetNames.forEach(name => {
      html += '<span class="status-badge loaded">&#10003; ' + name + '</span>';
    });
    if (ordersLoaded) {
      html += '<span class="status-badge loaded">&#10003; 受注データ</span>';
    } else {
      html += '<span class="status-badge pending-status">&#9679; 受注データ未読込</span>';
    }
    el.innerHTML = html;
  },

  /** シート名を YYMM 形式にパースしてソート用情報を返す */
  _parseSheetPeriod(name) {
    const n = name.trim();
    // "2601" → 2026年01月
    if (/^\d{4}$/.test(n)) {
      const yy = parseInt(n.substring(0, 2));
      const mm = parseInt(n.substring(2, 4));
      return { original: name, display: n, sortKey: (2000 + yy) * 100 + mm };
    }
    // "2025.12月" → 2512
    const m = n.match(/^(\d{4})\.(\d{1,2})月$/);
    if (m) {
      const year = parseInt(m[1]);
      const month = parseInt(m[2]);
      const display = String(year % 100).padStart(2, '0') + String(month).padStart(2, '0');
      return { original: name, display: display, sortKey: year * 100 + month };
    }
    return { original: name, display: name, sortKey: 0 };
  },

  /** ソート済み期間リスト（bindFilterEvents で参照） */
  _sortedPeriods: [],

  /** フィルターバー描画 */
  renderFilterBar(sheetNames, activeFilters) {
    const filterBar = document.getElementById('filter-bar');

    if (sheetNames.length === 0) {
      filterBar.classList.add('hidden');
      return;
    }
    filterBar.classList.remove('hidden');

    // パース＆ソート（古い順）
    const periods = sheetNames.map(n => this._parseSheetPeriod(n)).sort((a, b) => a.sortKey - b.sortKey);
    this._sortedPeriods = periods;

    const fromSelect = document.getElementById('filter-from');
    const toSelect = document.getElementById('filter-to');

    // 現在の選択を保持
    const prevFrom = fromSelect.value;
    const prevTo = toSelect.value;

    let options = '';
    periods.forEach((p, i) => {
      options += '<option value="' + i + '">' + p.display + '</option>';
    });
    fromSelect.innerHTML = options;
    toSelect.innerHTML = options;

    // 選択復元 or デフォルト（全期間）
    if (prevFrom && prevFrom < periods.length) {
      fromSelect.value = prevFrom;
    } else {
      fromSelect.value = '0';
    }
    if (prevTo && prevTo < periods.length) {
      toSelect.value = prevTo;
    } else {
      toSelect.value = String(periods.length - 1);
    }
  },

  /** パートナーテーブル用の月セレクタを更新 */
  renderPartnerMonthSelector(activeSheets) {
    const select = document.getElementById('partner-month');
    if (!select) return;

    // activeSheets を古い順にソート
    const periods = activeSheets.map(n => this._parseSheetPeriod(n))
      .sort((a, b) => a.sortKey - b.sortKey);

    const prev = select.value;

    let options = '<option value="">全期間</option>';
    periods.forEach(p => {
      options += '<option value="' + this._escapeHtml(p.original) + '">' + p.display + '</option>';
    });
    select.innerHTML = options;

    // 選択復元（もし現在の選択期間に含まれていれば維持、なければ全期間）
    if (prev && activeSheets.includes(prev)) {
      select.value = prev;
    } else {
      select.value = '';
    }
  },

  /** KPIカード描画 */
  renderKPICards(summary) {
    document.getElementById('kpi-cards').classList.remove('hidden');

    document.getElementById('kpi-referrals').textContent = summary.totalReferrals.toLocaleString() + '件';
    document.getElementById('kpi-campaign-cost').textContent = this.formatCurrency(summary.totalCampaignCost);
    document.getElementById('kpi-orders').textContent = summary.totalOrderCount.toLocaleString() + '件';
    document.getElementById('kpi-order-amount').textContent = this.formatCurrency(summary.totalOrderAmount);
    document.getElementById('kpi-closing-fee').textContent = this.formatCurrency(summary.totalClosingFee);

    const roiEl = document.getElementById('kpi-roi');
    roiEl.textContent = this.formatROI(summary.overallROI);
    roiEl.classList.remove('positive', 'negative');
    roiEl.classList.add(summary.overallROI >= 0 ? 'positive' : 'negative');
  },

  /** 月別推移テーブル描画 */
  renderMonthlyKPI(monthlyData) {
    const section = document.getElementById('monthly-kpi-section');
    const thead = document.getElementById('monthly-kpi-head');
    const tbody = document.getElementById('monthly-kpi-body');

    if (!monthlyData || monthlyData.months.length === 0) {
      section.classList.add('hidden');
      return;
    }
    section.classList.remove('hidden');

    // タイトルに月数を反映
    const titleEl = document.getElementById('monthly-kpi-title');
    if (titleEl) titleEl.textContent = '月別推移（' + monthlyData.months.length + 'か月）';

    // thead: 指標 | 月1 | 月2 | ... | 合計
    let headHtml = '<tr><th>指標</th>';
    monthlyData.months.forEach(m => {
      headHtml += '<th class="text-right">' + m.display + '</th>';
    });
    headHtml += '<th class="text-right monthly-total">合計</th></tr>';
    thead.innerHTML = headHtml;

    // 行定義
    const rows = [
      { label: '紹介数', key: 'totalReferrals', fmt: v => v.toLocaleString() + '件' },
      { label: 'CP費用', key: 'totalCampaignCost', fmt: v => this.formatCurrency(v) },
      { label: '受注数', key: 'totalOrderCount', fmt: v => v.toLocaleString() + '件' },
      { label: '受注金額', key: 'totalOrderAmount', fmt: v => this.formatCurrency(v) },
      { label: '成約フィー', key: 'totalClosingFee', fmt: v => this.formatCurrency(v) },
      { label: 'ROI', key: 'overallROI', fmt: v => this.formatROI(v), isRoi: true }
    ];

    let bodyHtml = '';
    rows.forEach(row => {
      bodyHtml += '<tr><td class="row-label">' + row.label + '</td>';
      monthlyData.months.forEach(m => {
        const val = m.summary[row.key];
        const cls = row.isRoi ? (val >= 0 ? ' positive' : ' negative') : '';
        bodyHtml += '<td class="text-right' + cls + '">' + row.fmt(val) + '</td>';
      });
      const totalVal = monthlyData.total[row.key];
      const totalCls = row.isRoi ? (totalVal >= 0 ? ' positive' : ' negative') : '';
      bodyHtml += '<td class="text-right monthly-total' + totalCls + '">' + row.fmt(totalVal) + '</td>';
      bodyHtml += '</tr>';
    });
    tbody.innerHTML = bodyHtml;
  },

  /** パートナー別テーブル描画 */
  renderPartnerTable(partnerKPIs, filter) {
    const section = document.getElementById('partner-table-section');
    const tbody = document.getElementById('partner-table-body');
    section.classList.remove('hidden');

    // フィルター適用（複数条件AND）
    let filtered = partnerKPIs;
    const filters = Array.isArray(filter) ? filter : (filter && filter.column ? [filter] : []);
    if (filters.length > 0) {
      filtered = partnerKPIs.filter(p => {
        return filters.every(f => {
          let val = p[f.column];
          if (f.column === 'conversionRate') val = val * 100;
          if (f.min !== null && val < f.min) return false;
          if (f.max !== null && val > f.max) return false;
          return true;
        });
      });
      const countEl = document.getElementById('filter-count');
      if (countEl) countEl.textContent = filtered.length + ' / ' + partnerKPIs.length + '件';
    }

    // ソート
    const sorted = this._sortPartnerKPIs(filtered);

    let html = '';
    sorted.forEach(p => {
      html += '<tr data-partner="' + this._escapeHtml(p.partnerName) + '">'
        + '<td>' + this._escapeHtml(p.partnerName) + '</td>'
        + '<td class="text-right">' + p.referralCount + '</td>'
        + '<td class="text-right">' + this.formatCurrency(p.campaignCost) + '</td>'
        + '<td class="text-right">' + p.orderCount + '</td>'
        + '<td class="text-right">' + this.formatCurrency(p.orderAmount) + '</td>'
        + '<td class="text-right">' + this.formatCurrency(p.closingFee) + '</td>'
        + '<td class="text-right">' + this.formatPercent(p.conversionRate) + '</td>'
        + '<td class="text-right ' + (p.roi >= 0 ? 'positive' : 'negative') + '">'
        + this.formatROI(p.roi) + '</td>'
        + '</tr>';
    });

    tbody.innerHTML = html;

    // ソートインジケータ更新
    document.querySelectorAll('#partner-table thead th').forEach(th => {
      th.classList.remove('sorted');
      const icon = th.querySelector('.sort-icon');
      if (icon) icon.innerHTML = '&#9650;';
    });
    const activeTh = document.querySelector('#partner-table thead th[data-col="' + this._sortState.column + '"]');
    if (activeTh) {
      activeTh.classList.add('sorted');
      const icon = activeTh.querySelector('.sort-icon');
      if (icon) icon.innerHTML = this._sortState.ascending ? '&#9650;' : '&#9660;';
    }
  },

  /** 詳細モーダル描画 */
  renderDetailModal(partnerKPI) {
    const titleEl = document.getElementById('modal-title');
    const bodyEl = document.getElementById('modal-body');

    titleEl.textContent = partnerKPI.partnerName + ' 詳細';

    let html = '';

    // KPIサマリー
    html += '<div class="modal-kpi-row">';
    html += this._modalKpiBox('紹介数', partnerKPI.referralCount + '件');
    html += this._modalKpiBox('CP費用', this.formatCurrency(partnerKPI.campaignCost));
    html += this._modalKpiBox('受注数', partnerKPI.orderCount + '件');
    html += this._modalKpiBox('受注金額', this.formatCurrency(partnerKPI.orderAmount));
    html += this._modalKpiBox('成約フィー', this.formatCurrency(partnerKPI.closingFee));
    html += this._modalKpiBox('受注率', this.formatPercent(partnerKPI.conversionRate));
    html += this._modalKpiBox('ROI', this.formatROI(partnerKPI.roi));
    html += '</div>';

    // 紹介企業一覧
    html += '<h3>紹介企業一覧</h3>';
    html += '<div style="overflow-x:auto">';
    html += '<table style="width:100%;font-size:0.875rem;min-width:600px">';
    html += '<thead><tr><th>紹介企業</th><th>初回商談日</th><th>受注</th><th>マッチ</th><th>種別</th><th class="text-right">計上金額</th><th class="text-right">CP単価</th></tr></thead>';
    html += '<tbody>';

    partnerKPI.referrals.forEach(r => {
      const companyName = r['紹介企業'] || '';
      const matchBadge = r._matchType === 'exact'
        ? '<span class="match-badge matched"></span>完全一致'
        : r._matchType === 'partial'
          ? '<span class="match-badge partial"></span>部分一致'
          : '<span class="match-badge unmatched"></span>未マッチ';

      const orderAmount = r._orders.reduce((s, o) =>
        s + (o._orderAmount || DataStore.normalizeAmount(o['計上金額'])), 0);
      const types = [...new Set(r._orders.map(o => (o['アポイント種別'] || '').trim()).filter(t => t))].join(', ');

      const meetingDate = this._formatDate(r['初回商談実施日']);

      html += '<tr>'
        + '<td>' + this._escapeHtml(companyName) + '</td>'
        + '<td>' + this._escapeHtml(meetingDate || '-') + '</td>'
        + '<td>' + (r._matched ? '&#10003;' : '-') + '</td>'
        + '<td>' + matchBadge + '</td>'
        + '<td>' + this._escapeHtml(types || '-') + '</td>'
        + '<td class="text-right">' + (orderAmount > 0 ? this.formatCurrency(orderAmount) : '-') + '</td>'
        + '<td class="text-right">' + this.formatCurrency(r._campaignCost || DataStore.normalizeAmount(r['キャンペーン単価'])) + '</td>'
        + '</tr>';
    });

    html += '</tbody></table>';
    html += '</div>';  // overflow-x wrapper

    // アポイント種別内訳
    const types = partnerKPI.byAppointType;
    if (Object.keys(types).length > 0) {
      html += '<h3>アポイント種別内訳</h3>';
      html += '<table style="width:100%;font-size:0.875rem">';
      html += '<thead><tr><th>種別</th><th class="text-right">件数</th><th class="text-right">計上金額</th></tr></thead>';
      html += '<tbody>';
      Object.entries(types).forEach(([type, data]) => {
        html += '<tr>'
          + '<td>' + this._escapeHtml(type) + '</td>'
          + '<td class="text-right">' + data.count + '件</td>'
          + '<td class="text-right">' + this.formatCurrency(data.amount) + '</td>'
          + '</tr>';
      });
      html += '</tbody></table>';
    }

    bodyEl.innerHTML = html;
  },

  /** モーダル表示/非表示 */
  showModal() {
    document.getElementById('detail-modal').classList.add('visible');
    document.body.style.overflow = 'hidden';
  },

  hideModal() {
    document.getElementById('detail-modal').classList.remove('visible');
    document.body.style.overflow = '';
  },

  /** ダッシュボードセクションの表示/非表示 */
  showDashboard() {
    document.getElementById('empty-state').classList.add('hidden');
    document.getElementById('charts-area').classList.remove('hidden');
  },

  hideDashboard() {
    document.getElementById('empty-state').classList.remove('hidden');
    document.getElementById('kpi-cards').classList.add('hidden');
    document.getElementById('monthly-kpi-section').classList.add('hidden');
    document.getElementById('charts-area').classList.add('hidden');
    document.getElementById('partner-table-section').classList.add('hidden');
    document.getElementById('filter-bar').classList.add('hidden');
  },

  // --- 内部ヘルパー ---

  _modalKpiBox(label, value) {
    return '<div class="modal-kpi"><div class="label">' + label + '</div><div class="value">' + value + '</div></div>';
  },

  _sortPartnerKPIs(partnerKPIs) {
    const col = this._sortState.column;
    const asc = this._sortState.ascending;
    return [...partnerKPIs].sort((a, b) => {
      let va = a[col], vb = b[col];
      if (typeof va === 'string') {
        return asc ? va.localeCompare(vb) : vb.localeCompare(va);
      }
      return asc ? va - vb : vb - va;
    });
  },

  handleSort(column) {
    if (this._sortState.column === column) {
      this._sortState.ascending = !this._sortState.ascending;
    } else {
      this._sortState.column = column;
      this._sortState.ascending = column === 'partnerName';
    }
  },

  _escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  },

  /** 日付フォーマット（Sheetsのシリアル値 or 文字列に対応） */
  _formatDate(value) {
    if (!value && value !== 0) return '';
    // Sheetsのシリアル値（数値）の場合: 1900/1/1起算
    if (typeof value === 'number' && value > 1000) {
      const d = new Date((value - 25569) * 86400000);
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      return y + '/' + m + '/' + day;
    }
    return String(value).trim();
  }
};
