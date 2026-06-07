// CLICLIC 체험단 평가 — Google Apps Script
// 사용법: Google Apps Script에 붙여넣기 후 웹앱으로 배포

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // 헤더가 없으면 첫 행에 추가
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        '제출시간','닉네임','수령제품','체험기간',
        '스테인리스경험','눌러붙음비교','세척편의성','손잡이편의성',
        '식기세척기/오븐','요리유형',
        '별점_탈부착','별점_그립감','별점_무게균형','별점_표면조리감',
        '별점_열전달','별점_세척','별점_디자인','별점_수납',
        '종합평점','좋았던점','아쉬운점','자유한마디',
        'NPS','재구매의향','SNS링크','SNS계정'
      ]);
    }

    const data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      data.timestamp, data.nick, data.product, data.duration,
      data.ss_exp, data.stick, data.clean, data.handle,
      data.appliance, data.cook_types,
      data.star_handle_attach, data.star_handle_grip, data.star_balance, data.star_surface,
      data.star_heat, data.star_clean, data.star_design, data.star_storage,
      data.avg_score, data.pros, data.cons, data.free,
      data.nps, data.repurchase, data.sns_link, data.sns_id
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 연결 테스트용 (브라우저에서 URL 직접 접속 시 확인)
function doGet(e) {
  return ContentService
    .createTextOutput('CLICLIC Sheets 연동 정상 작동 중')
    .setMimeType(ContentService.MimeType.TEXT);
}
