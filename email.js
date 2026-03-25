
/**
 * 논문 요약 결과를 이메일로 전송하는 함수
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - 논문 데이터가 있는 스프레드시트
 * @return {string} 처리 결과
 */
function sendSummariesToEmail(spreadsheet) {
  try {
    // 스프레드시트 객체 확인 및 대체 로직
    if (!spreadsheet) {
      console.error("스프레드시트 객체가 전달되지 않았습니다.");
      try {
        // 현재 활성화된 스프레드시트를 대신 사용
        spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        console.log("현재 활성화된 스프레드시트를 사용합니다: " + spreadsheet.getName());
      } catch (e) {
        console.error("활성화된 스프레드시트를 가져오는 데 실패했습니다:", e);
        return "스프레드시트를 찾을 수 없습니다.";
      }
    }
    
    // 시트 가져오기
    const sheet = spreadsheet.getSheetByName('journal_crawl_db');
    if (!sheet) {
      console.error("'journal_crawl_db' 시트를 찾을 수 없습니다.");
      return "시트 없음";
    }
    
    // 데이터 범위와 헤더 가져오기
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1) {
      console.error("전송할 데이터가 없습니다.");
      return "데이터 없음";
    }
    
    // 전체 데이터와 헤더 가져오기
    const headerRange = sheet.getRange(1, 1, 1, lastCol);
    const headers = headerRange.getValues()[0];
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const data = dataRange.getValues();
    
    // 필요한 열 인덱스 찾기
    const titleColIndex = headers.indexOf("Title");
    const journalColIndex = headers.indexOf("Journal");
    const dateColIndex = headers.indexOf("Date");
    const pmidColIndex = headers.indexOf("PMID");
    const pubTypeColIndex = headers.indexOf("Publication Type");
    const summaryColIndex = headers.indexOf(MESSAGES.SUMMARY_HEADER);
    const includedColIndex = headers.indexOf("Included");  // ✅ 필터링용 컬럼

    if (titleColIndex === -1 || pmidColIndex === -1 || summaryColIndex === -1) {
      console.error("필요한 열을 찾을 수 없습니다.");
      return "필요한 열 없음";
    }

    // ✅ Included="O"인 논문만 필터링
    const filteredData = data.filter((row, index) => {
      if (includedColIndex === -1) {
        // Included 컬럼이 없으면 모든 논문 포함 (하위 호환성)
        console.log(`Row ${index + 2}: No "Included" column, including all papers`);
        return true;
      }

      const isIncluded = row[includedColIndex];
      const hasSummary = row[summaryColIndex] && String(row[summaryColIndex]).trim() !== "";

      if (isIncluded === "O" && hasSummary) {
        console.log(`Row ${index + 2}: Included="O" with summary, adding to email`);
        return true;
      } else {
        console.log(`Row ${index + 2}: Included="${isIncluded}", hasSummary=${hasSummary}, skipping`);
        return false;
      }
    });

    console.log(`Total papers: ${data.length}, Filtered papers for email: ${filteredData.length}`);
    
    // ✅ 필터링된 논문이 없으면 이메일 전송 안 함
    if (filteredData.length === 0) {
      console.log("Included='O'인 논문이 없어서 이메일을 전송하지 않습니다.");
      return "필터링된 논문 없음";
    }

    // 현재 날짜 형식화
    const today = new Date();
    const formattedDate = Utilities.formatDate(today, "GMT+9", "yyyy-MM-dd");

    const startDate = new Date(today);
    startDate.setDate(today.getDate() - CONFIG.DAYS_RANGE);

    const searchPeriod = `${formatKoreanDate(startDate)}부터 ${formatKoreanDate(today)}까지`;
    Logger.log(searchPeriod);

    // ✅ 이메일 제목: 필터링된 논문 수 사용
    const subject = `[알레르기연구팀] ${searchPeriod} 두드러기/혈관부종/아나필락시스/비만세포증/식품알레르기 관련 논문 - 총 ${filteredData.length}개 논문`;

    // 이메일 본문 시작
    let emailBody = `<div style="font-family: Arial, sans-serif;">`;
    emailBody += `
    <h4 style="
      font-size: 22px;
      font-weight: 700;
      margin-bottom: 16px;
    ">
      최근 ${CONFIG.DAYS_RANGE}일 간 (${searchPeriod}) 두드러기/혈관부종/아나필락시스/비만세포증/식품알레르기 논문 요약 </h4>`;
    emailBody += `
    <p style="
      font-size: 16px;
      font-weight: 600;
      margin: 8px 0 18px 0;
    ">
      총 ${data.length}개 검색 중 상위 ${filteredData.length}개의 논문 요약을 공유합니다.</p>`;
    //emailBody += `<p>스프레드시트 링크: <a href="${spreadsheet.getUrl()}">${spreadsheet.getName()}</a></p>`;
    emailBody += `<hr style="margin: 20px 0;">`;

    // ✅ 각 논문 정보 추가 (필터링된 데이터만)
    for (let i = 0; i < filteredData.length; i++) {
      const row = filteredData[i];
      
      const title = row[titleColIndex] || "제목 정보 없음";
      const journal = journalColIndex !== -1 ? row[journalColIndex] : "저널 정보 없음";
      const date = dateColIndex !== -1 ? row[dateColIndex] : "";
      const pmid = row[pmidColIndex] || "PMID 정보 없음";
      const pubType = pubTypeColIndex !== -1 ? row[pubTypeColIndex] : "출판 유형 정보 없음";
      const summary = row[summaryColIndex] || "요약 정보 없음";
      
      // PubMed 링크 생성
      const pubmedLink = `https://pubmed.ncbi.nlm.nih.gov/${pmid}/`;
      
      // 요약 본문에서 정보 추출 (GPT 요약 형식에 따라 조정 필요)
      let authors = "정보 없음";
      let researchType = pubType || "정보 없음";
      
      // GPT 요약에서 저자 정보와 연구 유형 추출 시도
      //if (summary && summary.includes("* 저자:")) {
      //  const authorMatch = summary.match(/\* 저자:\s*([^\n]*)/);
      //  if (authorMatch && authorMatch[1]) {
      //    authors = authorMatch[1].trim();
      //  }
      //}
      
      //if (summary && summary.includes("* 연구 방법:")) {
      //  const methodMatch = summary.match(/\* 연구 방법:\s*([^\n]*)/);
      //  if (methodMatch && methodMatch[1]) {
      //    researchType = methodMatch[1].trim();
      //  }
      //}
      
      // 논문 요약 포맷팅
      emailBody += `<div style="margin-bottom: 30px; border: 1px solid #ddd; padding: 15px; border-radius: 5px;">`;
      emailBody += `<div style="border-bottom: 1px dashed #ccc; padding-bottom: 10px; margin-bottom: 10px;">`;
      emailBody += `<div style="font-size: 18px; font-weight: bold; color: #2c3e50; margin-bottom: 10px;">📔: ${title}</div>`;
      
      // 구분선과 논문 정보
      emailBody += `<div style="color: #555; font-size: 14px;">`;
      emailBody += `────────────────────────────── <br>`;
      //emailBody += `<strong>제목:</strong> ${title}<br>`;
      //emailBody += `<strong>게재일:</strong> ${year || "정보 없음"}<br>`;
      //emailBody += `<strong>저자:</strong> ${authors}<br>`;
      //emailBody += `<strong>연구 유형:</strong> ${researchType}<br>`;
      //emailBody += `<strong>저널:</strong> ${journal}<br>`;
      //emailBody += `<strong>요약:</strong><br>`;
      emailBody += `</div>`;
      emailBody += `</div>`;
      
      // 요약 내용
      emailBody += `<div style="font-size: 16px; line-height: 1.7; color: #333; background-color: #f9f9f9; padding: 10px; border-left: 4px solid #4285f4;">`;
      emailBody += summary.replace(/\n/g, '<br>');
      emailBody += `</div>`;
      
      // 링크 및 구분선
      emailBody += `<div style="margin-top: 10px; font-size: 14px;">`;
      emailBody += `<strong>링크:</strong> <a href="${pubmedLink}" target="_blank">${pubmedLink}</a><br>`;
      emailBody += `────────────────────────────── `;
      emailBody += `</div>`;
      emailBody += `</div>`;
    }
    
    // 이메일 본문 마무리
    emailBody += `<hr style="margin: 20px 0;">`;
    emailBody += '<p stype="font-size: 16px;"> <br> 최근 7일(전자출판기준) 발표된 두드러기/혈관부종/아나필락시스/비만세포증/식품알레르기 관련 논문들 중 선별한 논문들에 대한 요약입니다. </p>';
    emailBody += `<p style="color: #777; font-size: 12px;">이 이메일은 GPT에 의해 자동으로 생성되었습니다.</p>`;
    emailBody += `</div>`;
    
    // 플레인 텍스트 버전 생성 (HTML 태그 제거)
    const plainText = emailBody
      .replace(/<[^>]*>/g, '')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"');
    
    // 이메일 수신자 설정
    const recipients = CONFIG.EMAIL_RECIPIENTS;
    const primaryRecipient = CONFIG.EMAIL_TO_PRIMARY;

    // 이메일 전송
    MailApp.sendEmail({
      to: primaryRecipient,
      subject: subject,
      htmlBody: emailBody,
      body: plainText,
      bcc: recipients,
      name: "논문 요약 자동화"
    });
    console.log("DEBUG mail subject:", subject);
    console.log(`${recipients}에게 이메일 전송 완료`);
    return { ok: true, subject, emailBody };
    //return "이메일 전송 완료";
    
  } catch (error) {
    console.error("이메일 전송 오류:", error);
    return `이메일 전송 오류: ${error.message}`;
  }
}