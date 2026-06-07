// CLICLIC 체험단 평가 — Google Apps Script v4
// POST: 새 후기 저장 / GET: 후기 목록 반환 (JSONP 지원)

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        '제출시간','닉네임','수령제품','체험기간',
        '스테인리스경험','눌러붙음','세척편의성','손잡이편의성',
        '식기세척기오븐','요리유형',
        '별점_탈부착','별점_그립','별점_균형','별점_표면',
        '별점_열전달','별점_세척','별점_디자인','별점_수납',
        '종합평점','좋았던점','아쉬운점','자유한마디',
        'NPS','재구매','SNS링크','SNS계정','사진URLs'
      ]);
    }

    const d = JSON.parse(e.postData.contents);
    sheet.appendRow([
      d.timestamp, d.nick, d.product, d.duration,
      d.ss_exp, d.stick, d.clean, d.handle,
      d.appliance, d.cook_types,
      d.star_handle_attach, d.star_handle_grip, d.star_balance, d.star_surface,
      d.star_heat, d.star_clean, d.star_design, d.star_storage,
      d.avg_score, d.pros, d.cons, d.free,
      d.nps, d.repurchase, d.sns_link, d.sns_id,
      (d.photos || []).join('|')
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', msg: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// GET: 후기 목록 반환 (JSONP 콜백 지원)
function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rows = sheet.getDataRange().getValues();
    const callback = e.parameter.callback || '';

    if (rows.length <= 1) {
      const empty = JSON.stringify([]);
      return ContentService
        .createTextOutput(callback ? callback + '(' + empty + ')' : empty)
        .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
    }

    const data = rows.slice(1).reverse().map(r => {
      let dateStr = '';
      try {
        const d = new Date(r[0]);
        if (!isNaN(d)) dateStr = d.toISOString().slice(0, 10);
      } catch(e) {}

      return {
        nick:    r[1] ? String(r[1]) : '익명',
        product: r[2] ? String(r[2]) : '',
        avg:     isNaN(parseFloat(r[18])) ? 0 : Math.round(parseFloat(r[18]) * 10) / 10,
        pros:    r[19] ? String(r[19]) : '—',
        cons:    r[20] ? String(r[20]) : '—',
        date:    dateStr,
        hasPhoto: !!(r[26] && String(r[26]).trim()),
        photos:   r[26] ? String(r[26]).split('|').filter(u => u.startsWith('http')) : []
      };
    });

    const json = JSON.stringify(data);
    return ContentService
      .createTextOutput(callback ? callback + '(' + json + ')' : json)
      .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);

  } catch(err) {
    const empty = JSON.stringify([]);
    const callback = (e.parameter && e.parameter.callback) || '';
    return ContentService
      .createTextOutput(callback ? callback + '(' + empty + ')' : empty)
      .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
  }
}
