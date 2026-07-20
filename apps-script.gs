/**
 * 媽媽，妳睡了嗎？｜孕期親密關係問卷 — 後端接收程式
 * 用途：接住問卷網頁送來的答案，寫成試算表「一行一人」。
 * 使用方式見同資料夾的「設定說明」文件。
 */

// 固定欄位順序（跟問卷網頁的 labelMap 對齊；改題目時兩邊要一起改）
var HEADERS = [
  '填答時間',
  '年齡','性別','生活地區','目前狀態','回答哪一次','第幾胎','計畫中懷孕','關係狀態','在一起多久',
  '孕前-有無性生活','孕前-頻率','孕前-性滿意','孕前-關係滿意',
  '孕期-有無性生活','孕期-哪階段有','孕期-頻率','孕期-性滿意','孕期-擔心胎兒','孕期-影響原因','孕期-對關係影響','孕期-非性親密方式','孕期-舒服的方式','孕期-前戲需求','孕期-經驗分享',
  '產後-有無恢復性生活','產後-恢復時間','產後-影響原因','產後-性滿意','產後-擔心','產後-非性親密方式','產後-最大因素',
  '現在-親密狀態','現在-感受',
  '回顧-變化最大','回顧-如何溝通','回顧-給人的建議','回顧-補充'
];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('回覆') 
             || SpreadsheetApp.getActiveSpreadsheet().insertSheet('回覆');

    // 第一次寫入：先放標題列
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
    }

    // 依固定欄位順序取值，缺的留空
    var row = HEADERS.map(function(h){ return data[h] != null ? data[h] : ''; });
    sheet.appendRow(row);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// 讓瀏覽器測試用（直接開網址會看到這行字，代表部署成功）
function doGet() {
  return ContentService.createTextOutput('OK - 問卷後端運作中')
    .setMimeType(ContentService.MimeType.TEXT);
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
