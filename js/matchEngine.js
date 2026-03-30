// matchEngine.js - 紹介企業と受注情報の企業名突合

const MatchEngine = {
  /**
   * 企業名を正規化する
   * 法人格の除去、全角半角統一、スペース除去
   */
  normalize(name) {
    if (!name) return '';
    let n = String(name).trim();

    // 全角英数→半角
    n = n.replace(/[Ａ-Ｚａ-ｚ０-９]/g, s =>
      String.fromCharCode(s.charCodeAt(0) - 0xFEE0)
    );

    // 法人格を除去
    n = n.replace(/(株式会社|有限会社|合同会社|合資会社|合名会社|一般社団法人|一般財団法人|公益社団法人|公益財団法人|特定非営利活動法人|NPO法人|医療法人|社会福祉法人|学校法人)/g, '');
    n = n.replace(/[\(（](株|有|合|医|社|学|財)[\)）]/g, '');

    // スペース・記号を除去
    n = n.replace(/[\s　・\-\.\,、。\u3000]/g, '');

    return n.toLowerCase();
  },

  /**
   * 紹介企業リストと受注リストを突合する
   * @param {Array} referrals - 紹介企業レコード配列
   * @param {Array} orders - 受注レコード配列
   * @returns {Array} マッチ結果（各紹介企業に受注情報を付与）
   */
  matchAll(referrals, orders) {
    // 受注側の正規化名→レコード群のMapを構築
    const orderMap = new Map();
    orders.forEach(order => {
      const rawName = order['取引先名（正式名称）'] || order['取引先名'] || '';
      const key = this.normalize(rawName);
      if (key === '') return;
      if (!orderMap.has(key)) orderMap.set(key, []);
      orderMap.get(key).push(order);
    });

    return referrals.map(ref => {
      const refName = ref['紹介企業'] || '';
      const refKey = this.normalize(refName);

      if (refKey === '') {
        return { ...ref, _matched: false, _matchType: 'none', _orders: [] };
      }

      // 完全一致のみ（法人格除去・正規化後の一致）
      const exactMatches = orderMap.get(refKey);
      if (exactMatches && exactMatches.length > 0) {
        return { ...ref, _matched: true, _matchType: 'exact', _orders: exactMatches };
      }

      return { ...ref, _matched: false, _matchType: 'none', _orders: [] };
    });
  }
};
