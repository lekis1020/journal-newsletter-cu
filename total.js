/**
* 전체 워크플로우 실행 함수 (이메일 전송 포함)
 */

function runWeeklyDigestWorkflow() {
  try {
    console.log("워크플로우 시작...");

    // 1️⃣ PubMed → Spreadsheet
    const spreadsheet = fetchPubMedWeeklyAndSave();
    if (!spreadsheet) {
      sendNoResultsEmail("금주 검색 대상 논문이 없습니다.");
      return "논문 없음";
    }

    // 2️⃣ Notion (초록 저장)
    syncSheetToNotionPapersDB(spreadsheet, { includeSummary: false });

    // 3️⃣ GPT 점수 평가 및 Top 15 선정
    console.log("점수 평가 및 필터링 시작...");
    scoreAndFilterPapers(spreadsheet);
    console.log("점수 평가 완료");

    // 4️⃣ GPT 요약 생성 (Included="O"인 논문만)
    console.log("GPT 요약 생성 시작...");
    summarizePubMedArticlesWithGPT(spreadsheet);
    console.log("GPT 요약 완료");

    // 5️⃣ Notion (요약 업데이트)
    syncSheetToNotionPapersDB(spreadsheet, { includeSummary: true });

    // 6️⃣ 이메일 발송
    // 이메일 전송 (✅ subject/body를 받음)
    const mailResult = sendSummariesToEmail(spreadsheet);
    // ✅ 메일 성공 후 Notion 뉴스레터 DB에 저장/업데이트
    if (mailResult && mailResult.ok) {
      const dateISO = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
      overwriteNewsletterBodyByEmailHtml(
        dateISO,
        mailResult.subject,
        mailResult.emailBody
      );
    }
     console.log("워크플로우 완료");
     return "이메일 전송 및 노션 저장 완료";
  } catch (error) {
    console.error("워크플로우 오류:", error);
    return "오류: " + error.message;
  }
}
/**
* 검색된 논문이 없을때 안내 메일을 전송하는 함수
*/

function sendNoResultsEmail(message) {
      // 이메일 제목

  const today = new Date();
  const formattedDate = Utilities.formatDate(today, "GMT+9", "yyyy-MM-dd");

  const startDate = new Date(today);
  startDate.setDate(today.getDate() - CONFIG.DAYS_RANGE);

  const searchPeriod = `${formatKoreanDate(startDate)}부터 ${formatKoreanDate(today)}까지`;
  Logger.log(searchPeriod); 

  const emailSubject = `[호산구면역질환WG Journal Letter] ${searchPeriod} Eosinophil/Immunodysregulation ${searchPeriod} `;

  // 이메일 본문 시작
  let emailBody = `<div style="font-family: Arial, sans-serif;">`;
  emailBody += `<h4>최근 ${CONFIG.DAYS_RANGE}일 간 (${searchPeriod}) 호산구면역질환 관련 논문 요약 </h4>`;
  emailBody += `<p>지난 주 새로 출간된 논문은 검색되지 않았습니다.<br>`;
  emailBody += `평안한 한 주 보내시기 바랍니다.</p>`;
  emailBody += `<hr style="margin: 20px 0;">`;
// 이메일 본문 마무리
  emailBody += `<p style="color: #777; font-size: 12px;">이 이메일은 GPT에 의해 자동으로 생성되었습니다.</p>`;
  emailBody += `</div>`;

  const plainText = htmlToPlainText(emailBody);

  const recipients = CONFIG.EMAIL_RECIPIENTS;
  const primaryRecipient = CONFIG.EMAIL_TO_PRIMARY;

  MailApp.sendEmail({
    to: primaryRecipient,
    subject: emailSubject,
    bcc: recipients,
    name: "호산구 면역질환 논문 요약 자동화",
    htmlBody: emailBody,
    body: plainText,
  });
  //console.log(emailBody);
  console.log("결과 없음 이메일 전송 완료");
}

 
/**
 * 활성화된 스프레드시트의 요약 결과만 이메일로 전송하는 함수
 * (독립적으로 이메일 전송만 실행하고 싶은 경우 사용)
 */
function sendEmailFromActiveSpreadsheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (spreadsheet) {
      console.log("활성화된 스프레드시트: " + spreadsheet.getName());
      return sendSummariesToEmail(spreadsheet);
    } else {
      const errorMsg = "활성화된 스프레드시트가 없습니다.";
      console.error(errorMsg);
      return errorMsg;
    }
  } catch (error) {
    console.error("이메일 전송 실행 오류:", error);
    return "오류: " + error.message;
  }
}