// csvParser.js - CSV読み込みとパース

const CSVParser = {
  /**
   * SLキャンペーン対象リストCSVをパースする
   * ヘッダーが6行目にあるため、先頭5行をスキップ
   */
  parseCampaignCSV(file, sheetName) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const text = e.target.result;
        this._parseCampaignText(text, sheetName, resolve, reject);
      };
      // まずUTF-8で読み込み試行
      reader.readAsText(file, 'UTF-8');
    });
  },

  _parseCampaignText(text, sheetName, resolve, reject) {
    Papa.parse(text, {
      header: true,
      dynamicTyping: false,
      skipEmptyLines: 'greedy',
      beforeFirstChunk: function(chunk) {
        const lines = chunk.split(/\r\n|\r|\n/);
        // 先頭5行を除去して6行目をヘッダーにする
        if (lines.length > 5) {
          lines.splice(0, 5);
        }
        return lines.join('\n');
      },
      transformHeader: function(header) {
        return header.trim();
      },
      complete: function(results) {
        // 有効なデータ行のみフィルター（取次パートナーまたは紹介企業が存在する行）
        const data = results.data.filter(row => {
          const partner = row['取次パートナー'] || '';
          const company = row['紹介企業'] || '';
          return partner.trim() !== '' || company.trim() !== '';
        });
        resolve({ sheetName: sheetName, data: data });
      },
      error: function(error) {
        reject(error);
      }
    });
  },

  /**
   * 全社受注情報CSVをパースする（ヘッダーは1行目）
   */
  parseOrdersCSV(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const text = e.target.result;
        Papa.parse(text, {
          header: true,
          dynamicTyping: false,
          skipEmptyLines: 'greedy',
          transformHeader: function(header) {
            return header.trim();
          },
          complete: function(results) {
            // 有効な行のみ（取引先名が存在する行）
            const data = results.data.filter(row => {
              const name = row['取引先名（正式名称）'] || row['取引先名'] || '';
              return name.trim() !== '';
            });
            resolve(data);
          },
          error: function(error) {
            reject(error);
          }
        });
      };
      reader.readAsText(file, 'UTF-8');
    });
  },

  /**
   * Shift_JISで再読み込みしてパースし直す
   */
  parseCampaignCSVShiftJIS(file, sheetName) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const text = e.target.result;
        this._parseCampaignText(text, sheetName, resolve, reject);
      };
      reader.readAsText(file, 'Shift_JIS');
    });
  },

  parseOrdersCSVShiftJIS(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const text = e.target.result;
        Papa.parse(text, {
          header: true,
          dynamicTyping: false,
          skipEmptyLines: 'greedy',
          transformHeader: function(header) { return header.trim(); },
          complete: function(results) {
            const data = results.data.filter(row => {
              const name = row['取引先名（正式名称）'] || row['取引先名'] || '';
              return name.trim() !== '';
            });
            resolve(data);
          },
          error: function(error) { reject(error); }
        });
      };
      reader.readAsText(file, 'Shift_JIS');
    });
  },

  /**
   * 自動エンコーディング検出付きパース
   * UTF-8で試行→文字化け検出→Shift_JISでリトライ
   */
  async autoParseCampaignCSV(file, sheetName) {
    const result = await this.parseCampaignCSV(file, sheetName);
    if (this._hasGarbledText(result.data)) {
      return this.parseCampaignCSVShiftJIS(file, sheetName);
    }
    return result;
  },

  async autoParseOrdersCSV(file) {
    const result = await this.parseOrdersCSV(file);
    if (this._hasGarbledText(result)) {
      return this.parseOrdersCSVShiftJIS(file);
    }
    return result;
  },

  /** 文字化け判定（置換文字の存在チェック） */
  _hasGarbledText(data) {
    const sample = JSON.stringify(data).slice(0, 500);
    return /\ufffd/.test(sample);
  }
};
