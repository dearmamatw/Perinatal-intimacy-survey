/**
 * 媽媽，你睡了嗎？｜親密關係時光走廊問卷 — 後端（收資料 + 即時統計）
 * doPost：接住問卷答案，寫成試算表一行一人
 * doGet ：?stats=1 回傳統計 JSON（參與人數 + 指定題目分布），給完成頁與結果頁用
 */

var HEADERS = [
  '填答時間',
  '年齡','性別','生活地區','目前狀態','回答哪一次','第幾胎','計畫中懷孕','關係狀態','在一起多久',
  '孕前-有無性生活','孕前-頻率','孕前-性滿意','孕前-關係滿意',
  '孕期-有無性生活','孕期-哪階段有','孕期-頻率','孕期-性滿意','孕期-擔心胎兒','孕期-影響原因','孕期-對關係影響','孕期-非性親密方式','孕期-舒服的方式','孕期-前戲需求','孕期-經驗分享',
  '產後-有無恢復性生活','產後-恢復時間','產後-影響原因','產後-性滿意','產後-擔心','產後-非性親密方式','產後-最大因素',
  '現在-性生活狀態','現在-親密關係狀態','現在-感受',
  '回顧-變化最大','回顧-如何溝通','回顧-給人的建議','回顧-補充'
];

// 即時統計要公開的題目（只挑安全、有共鳴的；滿意度、體位等敏感題不公開）
var PUBLIC_STATS = [
  { col:'孕期-有無性生活', label:'孕期是否有性生活', type:'single' },
  { col:'孕期-擔心胎兒',   label:'進行時會不會擔心影響胎兒', type:'single' },
  { col:'孕期-影響原因',   label:'孕期影響性生活的原因（可複選）', type:'multi' },
  { col:'孕期-舒服的方式', label:'孕期覺得比較舒服安心的體位（可複選）', type:'multi', disclaimer:true },
  { col:'產後-恢復時間',   label:'產後多久恢復性生活', type:'single' },
  { col:'產後-影響原因',   label:'產後影響親密的原因（可複選）', type:'multi' },
  { col:'現在-感受',       label:'對此刻狀態的感受', type:'single' }
];

var MIN_N = 30; // 滿此門檻才公開比例分布（參與人數則永遠公開）

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getSheet_();
    if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
    var row = HEADERS.map(function(h){ return data[h] != null ? data[h] : ''; });
    sheet.appendRow(row);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  if (!e || !e.parameter || e.parameter.stats !== '1') {
    return ContentService.createTextOutput('OK - 問卷後端運作中')
      .setMimeType(ContentService.MimeType.TEXT);
  }
  try {
    var sheet = getSheet_();
    var lastRow = sheet.getLastRow();
    var n = Math.max(0, lastRow - 1);

    var result = { ok:true, n:n, minN:MIN_N, stats:[] };

    if (n >= MIN_N) {
      var values = sheet.getDataRange().getValues();
      var headerRow = values[0];

      PUBLIC_STATS.forEach(function(cfg){
        var colIdx = headerRow.indexOf(cfg.col);
        if (colIdx < 0) return;
        var counts = {};
        var answered = 0;
        for (var r = 1; r < values.length; r++) {
          var cell = String(values[r][colIdx] || '').trim();
          if (!cell || cell === '（跳過）') continue;
          answered++;
          var parts = (cfg.type === 'multi') ? cell.split('、') : [cell];
          parts.forEach(function(p){
            p = p.trim();
            if (p.indexOf('其他') === 0) p = '其他';
            if (!p) return;
            counts[p] = (counts[p] || 0) + 1;
          });
        }
        var items = Object.keys(counts).map(function(k){
          return { option:k, count:counts[k],
                   pct: answered ? Math.round(counts[k] / answered * 100) : 0 };
        }).sort(function(a,b){ return b.count - a.count; }).slice(0, 6);

        result.stats.push({ label:cfg.label, answered:answered, type:cfg.type, disclaimer:!!cfg.disclaimer, items:items });
      });
    }
    return json_(result);
  } catch (err) {
    return json_({ ok:false, error:String(err) });
  }
}

function getSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('回覆')
      || SpreadsheetApp.getActiveSpreadsheet().insertSheet('回覆');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
