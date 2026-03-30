// chartManager.js - Chart.js グラフ管理

const ChartManager = {
  _charts: {},

  /** パートナー別受注金額の横棒グラフ */
  renderPartnerAmountBar(canvasId, partnerKPIs) {
    this._destroyIfExists(canvasId);
    const sorted = [...partnerKPIs].sort((a, b) => b.orderAmount - a.orderAmount);
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;

    this._charts[canvasId] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: sorted.map(p => p.partnerName),
        datasets: [
          {
            label: '受注金額',
            data: sorted.map(p => p.orderAmount),
            backgroundColor: 'rgba(26, 115, 232, 0.7)',
            borderColor: 'rgba(26, 115, 232, 1)',
            borderWidth: 1
          },
          {
            label: 'CP費用+成約フィー',
            data: sorted.map(p => p.totalCost),
            backgroundColor: 'rgba(217, 48, 37, 0.5)',
            borderColor: 'rgba(217, 48, 37, 1)',
            borderWidth: 1
          }
        ]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'bottom', labels: { font: { size: 11 } } },
          tooltip: {
            callbacks: {
              label: function(ctx) {
                return ctx.dataset.label + ': ' + ctx.raw.toLocaleString() + '円';
              }
            }
          }
        },
        scales: {
          x: {
            ticks: {
              callback: function(v) {
                if (v >= 1000000) return (v / 1000000).toFixed(1) + 'M';
                if (v >= 1000) return (v / 1000).toFixed(0) + 'K';
                return v;
              }
            }
          }
        }
      }
    });
  },

  /** パートナー別ROI棒グラフ（正負で色分け） */
  renderROIBar(canvasId, partnerKPIs) {
    this._destroyIfExists(canvasId);
    const sorted = [...partnerKPIs].sort((a, b) => b.roi - a.roi);
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;

    this._charts[canvasId] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: sorted.map(p => p.partnerName),
        datasets: [{
          label: 'ROI (%)',
          data: sorted.map(p => Math.round(p.roi)),
          backgroundColor: sorted.map(p =>
            p.roi >= 0 ? 'rgba(13, 144, 79, 0.7)' : 'rgba(217, 48, 37, 0.7)'
          ),
          borderColor: sorted.map(p =>
            p.roi >= 0 ? 'rgba(13, 144, 79, 1)' : 'rgba(217, 48, 37, 1)'
          ),
          borderWidth: 1
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: function(ctx) { return 'ROI: ' + ctx.raw + '%'; }
            }
          }
        },
        scales: {
          y: {
            ticks: { callback: function(v) { return v + '%'; } }
          }
        }
      }
    });
  },

  /** 受注率ドーナツチャート */
  renderConversionDoughnut(canvasId, summaryKPI) {
    this._destroyIfExists(canvasId);
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;

    const won = summaryKPI.totalOrderCount;
    const lost = summaryKPI.totalReferrals - won;

    this._charts[canvasId] = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels: ['受注', '未受注'],
        datasets: [{
          data: [won, Math.max(0, lost)],
          backgroundColor: ['#0d904f', '#dadce0'],
          hoverOffset: 10,
          borderWidth: 2,
          borderColor: '#fff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '60%',
        plugins: {
          legend: { position: 'bottom', labels: { font: { size: 11 } } },
          tooltip: {
            callbacks: {
              label: function(context) {
                const total = context.dataset.data.reduce((a, b) => a + b, 0);
                const pct = total > 0 ? ((context.raw / total) * 100).toFixed(1) : 0;
                return context.label + ': ' + context.raw + '件 (' + pct + '%)';
              }
            }
          }
        }
      }
    });
  },

  /** チャート破棄 */
  _destroyIfExists(canvasId) {
    if (this._charts[canvasId]) {
      this._charts[canvasId].destroy();
      delete this._charts[canvasId];
    }
  },

  /** 全チャート更新 */
  updateAll(partnerKPIs, summaryKPI) {
    this.renderPartnerAmountBar('chart-partner-bar-canvas', partnerKPIs);
    this.renderROIBar('chart-roi-bar-canvas', partnerKPIs);
    this.renderConversionDoughnut('chart-conversion-canvas', summaryKPI);
  }
};
