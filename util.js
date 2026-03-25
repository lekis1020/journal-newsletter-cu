/**
 * 공통 유틸리티 함수 모음
 */

// ===== 날짜 관련 함수 =====

/**
 * 한국어 날짜 형식으로 포맷팅
 * @param {Date} date - 날짜 객체
 * @return {string} "M월 D일" 형식의 문자열
 */
function formatKoreanDate(date) {
  return `${date.getMonth() + 1}월 ${date.getDate()}일`;
}

/**
 * PubMed 검색을 위한 날짜 형식 변환
 * @param {Date} date - 날짜 객체
 * @return {string} "YYYY/MM/DD" 형식의 문자열
 */
function formatPubMedDate(date) {
  const y = date.getFullYear();
  const m = ("0" + (date.getMonth() + 1)).slice(-2);
  const day = ("0" + date.getDate()).slice(-2);
  return `${y}/${m}/${day}`;
}

/**
 * ISO 형식 날짜 문자열 반환
 * @return {string} "YYYY-MM-DD" 형식의 오늘 날짜
 */
function getTodayISO() {
  return Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
}

// ===== 텍스트 처리 함수 =====

/**
 * Notion 텍스트를 지정된 길이로 자르기
 * @param {string} s - 원본 문자열
 * @param {number} maxLen - 최대 길이
 * @return {string} 잘린 문자열
 */
function clampNotionText(s, maxLen) {
  const str = String(s || "");
  if (str.length <= maxLen) return str;
  return str.slice(0, maxLen - 1) + "…";
}

/**
 * 텍스트를 Notion rich_text 청크로 변환
 * @param {string} text - 변환할 텍스트
 * @param {number} chunkSize - 청크 크기 (기본값: 1900)
 * @return {Array} Notion rich_text 배열
 */
function toNotionRichTextChunks(text, chunkSize) {
  chunkSize = chunkSize || 1900;
  const s = String(text || "");
  const chunks = [];

  for (let i = 0; i < s.length; i += chunkSize) {
    chunks.push({ text: { content: s.slice(i, i + chunkSize) } });
  }

  // Notion API Limit: rich_text array max 100 items
  if (chunks.length > 100) {
    console.warn(`Text too long (${chunks.length} chunks), truncating to 100.`);
    const truncated = chunks.slice(0, 100);
    const last = truncated[99];
    if (last.text.content.length > (chunkSize - 20)) {
      last.text.content = last.text.content.slice(0, chunkSize - 20) + "...(truncated)";
    } else {
      last.text.content += "...(truncated)";
    }
    return truncated;
  }

  return chunks.length ? chunks : [{ text: { content: "" } }];
}

/**
 * 마크다운 볼드 표시를 HTML로 변환
 * @param {string} text - 변환할 텍스트
 * @return {string} 변환된 텍스트
 */
function normalizeBoldMarkup(text) {
  if (!text) return text;
  // **bold** -> <b>bold</b>
  return text.replace(/\*\*(.+?)\*\*/g, "<b>$1</b>");
}

/**
 * JSON 응답에서 마크다운 코드 블록 제거
 * @param {string} jsonStr - 원본 JSON 문자열
 * @return {string} 정리된 JSON 문자열
 */
function cleanMarkdownFromJson(jsonStr) {
  return jsonStr.replace(/```json/g, "").replace(/```/g, "").trim();
}

// ===== Notion API 관련 함수 =====

/**
 * Notion API 요청 헤더 생성
 * @return {Object} Notion API 헤더 객체
 */
function getNotionHeaders() {
  return {
    Authorization: `Bearer ${CONFIG.NOTION_API_KEY}`,
    "Notion-Version": "2022-06-28"
  };
}

// ===== 배열/데이터 처리 함수 =====

/**
 * 열 인덱스 찾기 (여러 후보 이름 지원)
 * @param {Array} headers - 헤더 배열
 * @param {Array} candidates - 후보 이름 배열
 * @return {number} 찾은 인덱스 (없으면 -1)
 */
function findColumnIndex(headers, candidates) {
  const normalize = (s) => String(s || "").trim().toLowerCase();
  const candidatesNorm = candidates.map(normalize).filter(Boolean);

  for (let i = 0; i < headers.length; i++) {
    if (candidatesNorm.includes(normalize(headers[i]))) {
      return i;
    }
  }
  return -1;
}

// ===== 에러 처리 함수 =====

/**
 * 재시도 로직을 포함한 API 호출
 * @param {Function} apiCallFn - 실행할 API 함수
 * @param {number} maxRetries - 최대 재시도 횟수
 * @param {number} baseDelay - 기본 대기 시간 (ms)
 * @param {number} multiplier - 지수 백오프 승수
 * @return {*} API 응답
 */
function retryApiCall(apiCallFn, maxRetries, baseDelay, multiplier) {
  maxRetries = maxRetries || CONFIG.MAX_RETRIES;
  baseDelay = baseDelay || CONFIG.RETRY_DELAY;
  multiplier = multiplier || CONFIG.RETRY_MULTIPLIER;

  let attempt = 0;
  let lastError;

  while (attempt < maxRetries) {
    try {
      return apiCallFn();
    } catch (error) {
      lastError = error;
      attempt++;

      if (attempt >= maxRetries) {
        throw error;
      }

      const delay = baseDelay * Math.pow(multiplier, attempt - 1);
      console.log(`API 재시도 ${attempt}/${maxRetries}, 대기: ${delay}ms`);
      Utilities.sleep(delay);
    }
  }

  throw lastError;
}

// ===== HTML/이메일 관련 함수 =====

/**
 * HTML을 플레인 텍스트로 변환
 * @param {string} html - HTML 문자열
 * @return {string} 플레인 텍스트
 */
function htmlToPlainText(html) {
  return html
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"');
}