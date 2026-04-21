// kpiCalculator.js - KPI算出ロジック

const KPICalculator = {
  /** 成約フィー = 受注金額 × 7% */
  calcClosingFee(orderAmount) {
    return orderAmount * 0.07;
  },

  /**
   * パートナー別KPIを算出
   */
  calcPartnerKPI(partnerName, matchedReferrals) {
    const referrals = matchedReferrals.filter(r =>
      (r['取次パートナー'] || '').trim() === partnerName
    );

    const referralCount = referrals.length;

    // キャンペーン費用合計
    const campaignCost = referrals.reduce((sum, r) =>
      sum + (r._campaignCost || DataStore.normalizeAmount(r['キャンペーン単価'])), 0);

    // 受注ありの紹介企業（ユニーク企業名ベース）
    const wonReferrals = referrals.filter(r => r._matched);
    const wonCompanies = new Set(wonReferrals.map(r => (r['紹介企業'] || '').trim()));
    const orderCount = wonCompanies.size;

    // 受注金額合計（同じ受注の重複カウントを防止）
    const countedOrders = new Set();
    let orderAmount = 0;
    const byAppointType = {};

    referrals.forEach(r => {
      r._orders.forEach(o => {
        if (countedOrders.has(o)) return; // 同一受注オブジェクトはスキップ
        countedOrders.add(o);
        const amt = o._orderAmount || DataStore.normalizeAmount(o['計上金額']);
        orderAmount += amt;

        const type = (o['アポイント種別'] || '不明').trim();
        if (!byAppointType[type]) byAppointType[type] = { count: 0, amount: 0 };
        byAppointType[type].count++;
        byAppointType[type].amount += amt;
      });
    });

    const closingFee = this.calcClosingFee(orderAmount);
    const totalCost = campaignCost + closingFee;
    const conversionRate = referralCount > 0 ? (orderCount / referralCount) : 0;
    const roi = totalCost > 0
      ? ((orderAmount - totalCost) / totalCost) * 100
      : 0;

    return {
      partnerName,
      referralCount,
      campaignCost,
      orderCount,
      orderAmount,
      closingFee,
      totalCost,
      conversionRate,
      roi,
      byAppointType,
      referrals
    };
  },

  /**
   * 月別KPIを算出（期間フィルターで選択された全月）
   * @param {Array} campaignSheets - DataStore._campaignSheets（フィルター済み）
   * @param {Array} orders - 受注データ
   * @returns {{ months: Array<{period, display, summary}>, total: Object }}
   */
  calcMonthlyKPIs(campaignSheets, orders) {
    // 各シートをパース＆ソート（古い順）
    const parsed = campaignSheets.map(s => ({
      ...s,
      period: UIRenderer._parseSheetPeriod(s.sheetName)
    })).sort((a, b) => a.period.sortKey - b.period.sortKey);

    // 各月ごとにKPI算出
    const months = parsed.map(sheet => {
      const matched = MatchEngine.matchAll(sheet.data, orders);
      const partnerNames = [...new Set(
        sheet.data.map(r => String(r['取次パートナー'] || '').trim()).filter(n => n)
      )];
      const partnerKPIs = partnerNames.map(name =>
        this.calcPartnerKPI(name, matched)
      );
      const summary = this.calcSummaryKPI(partnerKPIs);
      return {
        period: sheet.sheetName,
        display: sheet.period.display,
        summary: summary
      };
    });

    // 合計（全選択期間）
    const allData = campaignSheets.flatMap(s => s.data);
    const allMatched = MatchEngine.matchAll(allData, orders);
    const allPartnerNames = [...new Set(
      allData.map(r => String(r['取次パートナー'] || '').trim()).filter(n => n)
    )];
    const allPartnerKPIs = allPartnerNames.map(name =>
      this.calcPartnerKPI(name, allMatched)
    );
    const total = this.calcSummaryKPI(allPartnerKPIs);

    return { months, total };
  },

  /**
   * パートナー別 月次推移マトリックスを算出
   * @param {Array} campaignSheets - フィルター済みのシート配列
   * @param {Array} orders - 受注データ
   * @returns {{ months: Array<{period, display}>, partners: Array<{name, monthly, total}> }}
   */
  calcPartnerMonthlyMatrix(campaignSheets, orders) {
    // 各シートをパース＆ソート（古い順）
    const parsed = campaignSheets.map(s => ({
      ...s,
      period: UIRenderer._parseSheetPeriod(s.sheetName)
    })).sort((a, b) => a.period.sortKey - b.period.sortKey);

    const months = parsed.map(s => ({
      period: s.sheetName,
      display: s.period.display
    }));

    // 全パートナー名を収集
    const allPartners = new Set();
    parsed.forEach(s => {
      s.data.forEach(r => {
        const name = String(r['取次パートナー'] || '').trim();
        if (name) allPartners.add(name);
      });
    });

    // 各月のマッチング結果を事前計算
    const monthlyMatched = parsed.map(s => ({
      sheetName: s.sheetName,
      matched: MatchEngine.matchAll(s.data, orders)
    }));

    // 全期間のマッチング結果
    const allData = parsed.flatMap(s => s.data);
    const allMatched = MatchEngine.matchAll(allData, orders);

    // 各パートナーごとに月別KPIと合計を計算
    const partners = [...allPartners].map(name => {
      const monthly = {};
      monthlyMatched.forEach(m => {
        monthly[m.sheetName] = this.calcPartnerKPI(name, m.matched);
      });
      const total = this.calcPartnerKPI(name, allMatched);
      return { name, monthly, total };
    });

    // 合計の受注金額で降順ソート
    partners.sort((a, b) => b.total.orderAmount - a.total.orderAmount);

    return { months, partners };
  },

  /**
   * 全体サマリーKPIを算出
   */
  calcSummaryKPI(partnerKPIs) {
    const totalReferrals = partnerKPIs.reduce((s, p) => s + p.referralCount, 0);
    const totalCampaignCost = partnerKPIs.reduce((s, p) => s + p.campaignCost, 0);
    const totalOrderCount = partnerKPIs.reduce((s, p) => s + p.orderCount, 0);
    const totalOrderAmount = partnerKPIs.reduce((s, p) => s + p.orderAmount, 0);
    const totalClosingFee = partnerKPIs.reduce((s, p) => s + p.closingFee, 0);
    const totalCost = totalCampaignCost + totalClosingFee;
    const overallROI = totalCost > 0
      ? ((totalOrderAmount - totalCost) / totalCost) * 100
      : 0;

    return {
      totalReferrals,
      totalCampaignCost,
      totalOrderCount,
      totalOrderAmount,
      totalClosingFee,
      overallROI
    };
  }
};
