/**
 * 최근 N일(EDAT) AND 지정 저널 AND (호산구 또는 면역결핍/면역조절 키워드 포함) AND (암/항암/종양면역 키워드 포함 시 제외 + (선택) MeSH Major 암 주제 제외)
 *
 * 호산구면역질환 WG newsletter
 * immune related disease
 * PID
 * - Eosinophilia
 * 제목에 호산구가 들어가는 모든 논문 검색
 * 특정 저널에서 원하는 mesh Terms 을 가지고 있는 논문 검색
 * 근거 수준 높은 논문들로만 선별
 * PubMed 논문 검색 및 GPT 요약 시스템
 * 1. PubMed에서 최근 논문 검색
 * 2. 논문 초록 및 메타데이터 수집
 * 3. GPT를 사용한 요약 생성
 * 4. 결과를 스프레드시트에 저장
 *
 * 설정은 config.gs에서 관리됩니다.
 */

// ===== 2. 메인 실행 함수 =====

/**
 * 전체 워크플로우를 실행하는 함수
 * 1. PubMed에서 최근 논문 검색 및 저장
 * 2. 저장된 논문에 대해 GPT 요약 생성
 */
/**
 * 전체 워크플로우 실행 함수 (변경됨)
 */
// 2. 메인 실행 함수
function fetchAndSummarizeAll() {
  // 1) PubMed 검색 및 스프레드시트 저장 (Raw Data)
  const spreadsheet = fetchPubMedWeeklyAndSave();
  if (!spreadsheet) {
    console.error("스프레드시트 생성 실패");
    return;
  }

  // 2) GPT 점수 산정 및 필터링 (점수 계산 후 'Included' 여부 표시)
  //    반환값: 필터링된 논문 개수 (혹은 처리 결과)
  scoreAndFilterPapers(spreadsheet);

  // 3) Notion에 원본 메타/초록 저장 (요약 제외)
  //    (옵션: 점수 필터 통과한 것만 저장하려면 여기서 필터링 로직 추가 필요. 
  //     일단 모든 검색 결과를 저장하고 싶다면 그대로 두거나, 
  //     Filtered 된 것만 원한다면 4번 단계와 합칠 수도 있음.
  //     사용자 요구사항: "뉴스레터에는 2점 이상의 검색 결과들을 포함" -> Notion 저장도 필터링 된 것만 하는 게 자연스러움)
  
  // 4) GPT 요약 생성 (Sheet의 Summary 컬럼 채움 - Included 인 것만)
  summarizePubMedArticlesWithGPT(spreadsheet);

  // 5) Notion에 Summary 업데이트 (Included 인 것만)
  syncSheetToNotionPapersDB(spreadsheet, { includeSummary: true, filterHighScoresOnly: true });

  console.log("전체 워크플로우 완료");
}

function scoreAndFilterPapers(spreadsheet) {
  const sheet = spreadsheet.getSheetByName("journal_crawl_db");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 헤더 추가
  const newHeaders = ["Scores", "Final Score", "Included", "Exclusion Reason"];
  const startCol = headers.length + 1;
  sheet.getRange(1, startCol, 1, newHeaders.length).setValues([newHeaders]);
  
  // 인덱스 찾기
  const titleIdx = headers.indexOf("Title");
  const abstractIdx = headers.indexOf("Abstract");
  const journalIdx = headers.indexOf("Journal");
  
  if (titleIdx === -1 || abstractIdx === -1) {
    console.error("필수 컬럼 누락");
    return;
  }

  const rows = data.slice(1);
  const results = []; // 처리 결과 임시 저장

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const title = row[titleIdx];
    const abstract = String(row[abstractIdx] || "").trim();
    const journal = row[journalIdx];

    // 1. [NEW] 초록 없음 -> 즉시 제외
    if (!abstract) {
        results.push({
            details: "No Abstract",
            score: 0,
            included: false,
            reason: "No Abstract"
        });
        continue;
    }

    try {
      // 2. GPT 평가
      const relevance = evaluatePaperRelevance(title, abstract);
      
      // 3. 점수 계산
      const calc = calculateScore(relevance, journal);
      
      // 결과 저장
      results.push({
        details: JSON.stringify(relevance),
        score: calc.finalScore,
        included: calc.isIncluded, // 점수 기준(3점) 통과 여부
        reason: calc.reason,
        originalIndex: i
      });
      
      // Rate Limit 방지
      Utilities.sleep(500); 

    } catch (e) {
      console.error(`Row ${i+2} scoring error:`, e);
      results.push({
          details: "Error",
          score: 0,
          included: false,
          reason: e.message
      });
    }
  }

  // 4. [NEW] Top 15 선정 로직
  // 점수 통과(isIncluded=true)한 것들만 필터링
  const passedCandidates = results
    .map((r, idx) => ({ ...r, arrayIndex: idx }))
    .filter(r => r.included);

  // 점수 내림차순 정렬
  passedCandidates.sort((a, b) => b.score - a.score);

  // 상위 N개만 유지, 나머지는 탈락 처리
  const cutoffIndex = Math.min(passedCandidates.length, CONFIG.MAX_INCLUSION);
  
  // 선정된 인덱스 Set
  const finalIncludedIndices = new Set(passedCandidates.slice(0, cutoffIndex).map(c => c.arrayIndex));

  // 결과 배열 생성 (Spreadsheet 업데이트용)
  const updates = results.map((r, idx) => {
      let isFinalIncluded = r.included;
      let finalReason = r.reason;

      // 점수는 넘었으나 Top 15 등수에 못 든 경우
      if (r.included && !finalIncludedIndices.has(idx)) {
          isFinalIncluded = false;
          finalReason += " (Excluded by Top 15 Limit)";
      }

      return [
          r.details,
          r.score,
          isFinalIncluded ? "O" : "X", // 최종 O/X
          finalReason
      ];
  });

  // 배치 업데이트
  sheet.getRange(2, startCol, updates.length, newHeaders.length).setValues(updates);
}

function evaluatePaperRelevance(title, abstract) {
  // ✅ OPTIMIZED: 프롬프트 압축 (토큰 30% 절약)
  const prompt = `Detect keywords in title/abstract:

Title: "${title}"
Abstract: "${abstract}"

Categories:
1. CANCER: cancer, carcinoma, tumor, malignancy, neoplasm, metastasis, oncology, lymphoma, leukemia, sarcoma, melanoma
2. EOSINOPHIL: eosinophil*, hypereosinophil*, HES, EGPA, churg-strauss, eosinophilic esophagitis
3. IMMUNE: immunodeficiency, primary immunodeficiency, IEI, CVID, SCID, agammaglobulinemia, hyper-IgM, selective IgA (exclude: immune checkpoint, CAR-T)

Return JSON:
{
  "hasCancerInTitle": bool,
  "hasCancerInAbstract": bool,
  "hasEosInTitle": bool,
  "hasEosInAbstract": bool,
  "hasImmuneInTitle": bool,
  "hasImmuneInAbstract": bool
}`;

  const jsonStr = callGPT(prompt, true, CONFIG.OPENAI_API_KEY_SCORING); // 점수 산정용 키 사용
  
  if (!jsonStr) {
      console.error("GPT returned null/empty response");
      return {
        hasCancerInTitle: false, hasCancerInAbstract: false,
        hasEosInTitle: false, hasEosInAbstract: false,
        hasImmuneInTitle: false, hasImmuneInAbstract: false
      };
  }

  try {
    // [DEBUG] 원본 응답 확인
    console.log("DEBUG GPT Raw response:", jsonStr);

    // 마크다운 코드 블록 제거
    const cleanStr = jsonStr.replace(/```json/g, "").replace(/```/g, "").trim();
    const parsed = JSON.parse(cleanStr);
    console.log("DEBUG Parsed Flags:", JSON.stringify(parsed));
    return parsed;
  } catch (e) {
    console.error("JSON parse error", e, "Raw:", jsonStr);
    return {
      hasCancerInTitle: false, hasCancerInAbstract: false,
      hasEosInTitle: false, hasEosInAbstract: false,
      hasImmuneInTitle: false, hasImmuneInAbstract: false
    };
  }
}

function calculateScore(flags, journalName) {
  let cancerScore = 0;
  let eosScore = 0;
  let immuneScore = 0;
  
  // Cancer Scoring
  // "cancer 와 관련된 키워드가 제목, 결론에 포함된 경우 각각 1점씩 감점을 하고, 제목, 결론에 모두 포함된 경우에는 3점을 감점한다."
  if (flags.hasCancerInTitle && flags.hasCancerInAbstract) {
    cancerScore = -3;
  } else {
    if (flags.hasCancerInTitle) cancerScore -= 1;
    if (flags.hasCancerInAbstract) cancerScore -= 1;
  }
  
  // Eosinophil Scoring
  // "제목, 결론에 포함된 경우 각각 2점씩 점수를 매기고, 제목, 결론에 모두 포함된 경우 전체 5점을 매긴다."
  if (flags.hasEosInTitle && flags.hasEosInAbstract) {
    eosScore = 5;
  } else {
    if (flags.hasEosInTitle) eosScore += 2;
    if (flags.hasEosInAbstract) eosScore += 2;
  }

  // Immune Scoring
  // "면역 질환... 제목, 결론에 포함된 경우 각각 2점씩 점수를 매기고, 제목, 결론에 모두 포함된 경우 전체 5점을 매긴다."
  if (flags.hasImmuneInTitle && flags.hasImmuneInAbstract) {
    immuneScore = 5;
  } else {
    if (flags.hasImmuneInTitle) immuneScore += 2;
    if (flags.hasImmuneInAbstract) immuneScore += 2;
  }
  
  // Journal Bonus
  // "저널이름에 'allergy' 가 들어가 있는 경우에는 전체 점수에서 1점 가산점"
  //let journalBonus = 0;
  //if (journalName && journalName.toLowerCase().includes("allergy")) {
  //  journalBonus = 1;
  //}
  
  // Final Calculation
  // "호산구와 면역질환은 따로 점수를 매기고 둘중 점수가 높은 것을 선택하여 cancer 관련된 점수와 합산하여 최종 점수를 결정한다."
  const baseScore = Math.max(eosScore, immuneScore);
  const finalScore = baseScore + cancerScore ;
  
  // Threshold
  // "전체 점수가 CONFIG.MIN_RELEVANCE_SCORE(3) 이상인 경우 ... 포함"
  const isIncluded = finalScore >= CONFIG.MIN_RELEVANCE_SCORE;
  
  return {
    finalScore,
    isIncluded,
    reason: `Base(${baseScore}) + Cancer(${cancerScore})  = ${finalScore}`
  };
}

function syncSheetToNotionPapersDB(spreadsheet, opts) {
  opts = opts || {};
  const includeSummary = !!opts.includeSummary;

  const sheet = spreadsheet.getSheetByName("journal_crawl_db");
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();

  const norm = (s) => String(s || "").trim().toLowerCase();
  const findIdx = (candidates) => {
    const cand = candidates.map(norm).filter(Boolean);
    for (let i = 0; i < headers.length; i++) {
      if (cand.includes(norm(headers[i]))) return i;
    }
    return -1;
  };

  const titleIdx    = findIdx(["Title"]);
  const journalIdx  = findIdx(["Journal"]);
  const dateIdx     = findIdx(["Date"]);
  const authorsIdx  = findIdx(["Authors"]);
  const pmidIdx     = findIdx(["PMID"]);
  const pubTypeIdx  = findIdx(["Publication Type"]);
  const abstractIdx = findIdx(["Abstract"]);

  // Summary 헤더는 유연하게 탐색
  const summaryIdx = findIdx([
    "Summary",
    "GPT Summary",
    "GPT 요약",
    "요약",
    "요약 결과",
    "AI Summary",
    (typeof MESSAGES !== "undefined" && MESSAGES.SUMMARY_HEADER) ? MESSAGES.SUMMARY_HEADER : ""
  ]);

  if (titleIdx === -1 || pmidIdx === -1) {
    throw new Error("Sheet headers missing: Title/PMID");
  }

  for (const row of values) {
    const pmid = row[pmidIdx];
    if (!pmid) continue;

    const paper = {
      pmid: String(pmid),
      title: row[titleIdx],
      journal: journalIdx !== -1 ? row[journalIdx] : "",
      date: dateIdx !== -1 ? row[dateIdx] : "",
      authors: authorsIdx !== -1 ? row[authorsIdx] : "",
      pubType: pubTypeIdx !== -1 ? row[pubTypeIdx] : "",
      abstract: abstractIdx !== -1 ? row[abstractIdx] : "",
      summary: (includeSummary && summaryIdx !== -1) ? row[summaryIdx] : ""
    };

  // includeSummary=true인데 summary가 비어 있으면 굳이 업데이트 안 하도록
    if (includeSummary && !paper.summary) continue;

    // 점수 기반 필터링 (Optional)
    if (opts.filterHighScoresOnly) {
       // Sheet에 있는 Included 컬럼 값을 확인해야 하는데, 
       // 현재 values는 초기 읽어온 값이라 업데이트된 점수 컬럼이 없을 수 있음.
       // 따라서, summarizePubMedArticlesWithGPT 이후에 이 함수를 호출할 때는
       // 시트를 다시 읽거나 인자를 통해 제어해야 함.
       // 이번 구현에서는 summarizePubMedArticlesWithGPT가 "Included"가 아닌 행엔 요약을 안 남기는 방식 등을 권장.
       
       // 간단히: summary가 있는 것만 보낸다면, summarize 함수에서 Included 인 것만 요약하면 됨.
       if (!paper.summary || paper.summary.startsWith("오류") || paper.summary === "") continue;
    }

    upsertPaperToNotion(paper, { includeSummary });
    Utilities.sleep(350);
  }

  console.log(`Notion papers sync done (includeSummary=${includeSummary})`);
}

function upsertPaperToNotion(paper, opts) {
  opts = opts || {};
  const includeSummary = !!opts.includeSummary;

  const existingId = findPaperPageIdByPMID(paper.pmid);

  const props = {
    "Title": { title: [{ text: { content: String(paper.title || "제목 없음") } }] },
    "PMID": { rich_text: [{ text: { content: String(paper.pmid) } }] },
    "Journal": { rich_text: [{ text: { content: String(paper.journal || "") } }] },
    "PubMed Link": { url: `https://pubmed.ncbi.nlm.nih.gov/${paper.pmid}/` }
  };

  // (있을 때만) Authors
  if (paper.authors) {
    props["Authors"] = { rich_text: [{ text: { content: String(paper.authors) } }] };
  }

  // ✅ Abstract: 최대 보존
  if (paper.abstract !== undefined) {
    props["Abstract"] = { rich_text: toNotionRichTextChunks(paper.abstract, 1900) };
  }

  // PubType
  if (paper.pubType) {
    props["Publication Type"] = {
      multi_select: String(paper.pubType)
        .split(",")
        .map(t => ({ name: t.trim() }))
        .filter(x => x.name)
    };
  }

  // Date
  if (paper.date instanceof Date) {
    props["Date"] = { date: { start: Utilities.formatDate(paper.date, "GMT+9", "yyyy-MM-dd") } };
  } else if (/^\d{4}-\d{2}-\d{2}$/.test(String(paper.date))) {
    props["Date"] = { date: { start: String(paper.date) } };
  }

  // Digest Date (DB에 있는 경우에만 쓰세요)
  // props["Digest Date"] = { date: { start: getTodayISO() } };

  // ✅ Summary: 잘라서 저장
  if (includeSummary) {
    props["Summary"] = {
      rich_text: [{ text: { content: clampNotionText(paper.summary, 2000) } }]
    };
  }

  if (existingId) {
    return notionUpdatePage(existingId, props);
  }
  return notionCreatePage(CONFIG.NOTION_PAPERS_DB_ID, props);
}

function findPaperPageIdByPMID(pmid) {
  const url = `https://api.notion.com/v1/databases/${CONFIG.NOTION_PAPERS_DB_ID}/query`;

  const payload = {
    page_size: 1,
    filter: {
      property: "PMID",
      rich_text: { equals: String(pmid) }
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: getNotionHeaders(),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  const body = res.getContentText();

  if (code !== 200) {
    throw new Error("Notion query failed: " + body);
  }

  const json = JSON.parse(body);
  return (json.results && json.results.length > 0) ? json.results[0].id : null;
}

function notionCreatePage(databaseId, props) {
  const url = "https://api.notion.com/v1/pages";

  const payload = {
    parent: { database_id: databaseId },
    properties: props
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: getNotionHeaders(),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  const body = res.getContentText();

  if (code !== 200) {
    throw new Error("Notion create failed: " + body);
  }

  return JSON.parse(body).id;
}


function notionUpdatePage(pageId, props) {
  const url = `https://api.notion.com/v1/pages/${pageId}`;

  const payload = { properties: props };

  const options = {
    method: "patch",
    contentType: "application/json",
    headers: getNotionHeaders(),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  const body = res.getContentText();

  if (code !== 200) {
    throw new Error("Notion update failed: " + body);
  }

  return true;
}



function upsertNewsletterToNotion(params) {
  const dateISO = params.dateISO;          // "yyyy-MM-dd"
  const subject = params.subject || "";
  const htmlBody = params.htmlBody || "";
  const spreadsheetUrl = params.spreadsheetUrl || "";

  const existingId = findNewsletterPageIdByDate(dateISO);

  const props = {
    "Title": { title: [{ text: { content: `${dateISO} Newsletter` } }] },
    "Date": { date: { start: dateISO } },
    "Subject": { rich_text: [{ text: { content: subject } }] },

    // ✅ HTML은 길어지므로 chunk 저장 (각 chunk 2000 미만 권장)
    "Email HTML": { rich_text: toNotionRichTextChunks(htmlBody, 1900) },

    "Sent At": { date: { start: new Date().toISOString() } }
  };

  if (spreadsheetUrl) props["Spreadsheet URL"] = { url: spreadsheetUrl };

  if (existingId) {
    notionUpdatePage(existingId, props);
    return existingId;
  } else {
    return notionCreatePage(CONFIG.NOTION_NEWSLETTER_DB_ID, props);
  }
}

function testNotionWrite() {
  const today = getTodayISO();
  const pageId = getOrCreateDailyPageId(today);
  appendPaperToggle(pageId, "TEST TITLE", "00000000", "✅ TEST BLOCK: 노션 API 연결 성공");
}

/**
 * PubMed에서 최근 7일(EDAT) 논문 검색 및 저장
 */ 

function fetchPubMedWeeklyAndSave() {
  const DAYS_RANGE = Number(CONFIG.DAYS_RANGE) || 7;
  const MAX_RESULTS = Number(CONFIG.MAX_RESULTS) || 100;

  /* ---------- EDAT 날짜 조건 ---------- */
  const today = new Date();
  const start = new Date();
  start.setDate(today.getDate() - DAYS_RANGE);

  const formatDate = d => {
    const y = d.getFullYear();
    const m = ("0" + (d.getMonth() + 1)).slice(-2);
    const day = ("0" + d.getDate()).slice(-2);
    return `${y}/${m}/${day}`;
  };

  const dateRangeEDAT =
    `"${formatDate(start)}"[EDAT] : "${formatDate(today)}"[EDAT]`;

  /* ---------- (선택) 저널 필터 ---------- */
  const journalQuery = (CONFIG.JOURNALS || [])
    .map(j => `"${j}"[Journal]`)
    .join(" OR ");
  const journalFilter = journalQuery ? `(${journalQuery})` : null;

  /* ---------- MUST HAVE : 특이 질환 키워드 ---------- */
  const mustHaveBlock = `
  (
    eosinophil*[Title/Abstract]
    OR eosinophilia*[Title/Abstract]
    OR hypereosinophil*[Title/Abstract]
    OR "hypereosinophilic syndrome"[Title/Abstract]
    OR HES[Title/Abstract]
    OR EGPA[Title/Abstract]
    OR "churg-strauss"[Title/Abstract]
    OR "eosinophilic esophagitis"[Title/Abstract]

    OR immunodeficien*[Title/Abstract]
    OR "primary immunodeficien*"[Title/Abstract]
    OR "inborn errors of immunity"[Title/Abstract]
    OR IEI[Title/Abstract]
    OR CVID[Title/Abstract]
    OR SCID[Title/Abstract]
    OR agammaglobulin*[Title/Abstract]
    OR "hyper-IgM"[Title/Abstract]
    OR "selective IgA"[Title/Abstract]
    OR hypogammaglobulinemia*[Title/Abstract]

  )
  `;

  /* ---------- EXCLUDE : 암 / 항암 / 종양면역 (Title 기준 강제) ---------- */
  const cancerExclusionBlock = `
  (
    /* --- Text (Title/Abstract) --- */
    cancer*[tiab]
    OR carcinoma*[tiab]
    OR tumor*[tiab]
    OR malignan*[tiab]
    OR neoplas*[tiab]
    OR metast*[tiab]
    OR oncolog*[tiab]
    OR glioma*[tiab]
    OR gliomagenesis[tiab]
    OR lymphoma*[tiab]
    OR leukemia*[tiab]
    OR neuroblastoma*[tiab]
    OR sarcoma*[tiab]
    OR melanoma*[tiab]
    OR "squamous cell carcinoma"[tiab]
    OR "germ cell tumor"[tiab]
    OR "immune checkpoint"[tiab]
    OR checkpoint*[tiab]
    OR PD-1[tiab]
    OR PD-L1[tiab]
    OR CTLA-4[tiab]
    OR CAR-T[tiab]
    OR TCR-T[tiab]
    OR "cell therapy"[tiab]
    OR "tumor microenvironment"[tiab]
    OR "immune escape"[tiab]
    OR "tumor suppress*"[tiab]
    OR xenograft*[tiab]
    OR organoid*[tiab]
    OR Neoplasms[Majr]
    OR Immune Checkpoint Inhibitors[Majr]
    OR Carcinoma[Majr]
    OR Lymphoma[Majr]
    OR Leukemia[Majr]
  )
`;

  /* ---------- 최종 쿼리 조립 ---------- */
  // 3. 쿼리 수정: 암 제외 로직 제거, 저널 제한 제거
  // "대원칙 (호산구 질환, 면역결핍, 면역 조절 질환 들은 포함하고, 암에 관련된 것은 제외함) 은 지키고"
  // -> 여기서 "제외함"은 검색 단계에서의 제외가 아니라 "점수 평가를 통한 제외"로 해석됨 (사용자 요청: "오로지 평가는 GPT 가 평가 하여... 결정한다")
  // 따라서 검색은 Broad 하게 가져와야 함.
  
  const andParts = [
    `(${dateRangeEDAT})`,
    journalFilter, // 저널 필터 복구
    mustHaveBlock,
    // `NOT ${cancerExclusionBlock}` // 암 제외는 여전히 GPT 점수제로 유지
  ].filter(Boolean);

  let finalQuery = andParts.join(" AND ");
  finalQuery = finalQuery.replace(/\s+/g, " ").trim();

  console.log("DEBUG finalQuery =", finalQuery);

  /* ---------- PubMed ESearch ---------- */
  const esearchUrl =
    "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi";

  const params = {
    db: "pubmed",
    term: finalQuery,
    retmode: "json",
    retmax: MAX_RESULTS,
    sort: "pubdate",
    usehistory: "y"
  };
  if (CONFIG.PUBMED_API_KEY) params.api_key = CONFIG.PUBMED_API_KEY;

  const payload = Object.keys(params)
    .map(k => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`)
    .join("&");

  const options = {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    payload,
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(esearchUrl, options);
    if (response.getResponseCode() !== 200) {
      console.error("ESearch 실패:", response.getContentText());
      return null;
    }

    const json = JSON.parse(response.getContentText());
    const err = json?.esearchresult?.ERROR;
    if (err) {
      console.error("PubMed ERROR:", err);
      console.error("QUERY:", finalQuery);
      return null;
    }

    const idList = json?.esearchresult?.idlist || [];
    if (idList.length === 0) {
      console.log("조건에 맞는 논문 없음");
      return null;
    }

    const detailedResults = fetchPubMedData(idList);
    if (!detailedResults || detailedResults.length === 0) return null;

    /* ---------- 엄격 분류 (Original / Review / Other) ---------- */
    const REVIEW_TYPES = ["systematic review", "meta-analysis", "review"];
    const ORIGINAL_TYPES = [
      "clinical trial",
      "randomized controlled trial",
      "controlled clinical trial",
      "observational study",
      "comparative study",
      "multicenter study",
      "evaluation study"
    ];

    const enriched = detailedResults.map(r => {
      const pubTypes = r.publicationTypes || [];
      const typesLower = pubTypes.map(t => String(t).toLowerCase());

      const isReview = typesLower.some(t =>
        REVIEW_TYPES.some(rt => t.includes(rt))
      );
      const isOriginal = typesLower.some(t =>
        ORIGINAL_TYPES.some(ot => t.includes(ot))
      );

      return Object.assign({}, r, {
        articleCategory: isReview
          ? "Review"
          : isOriginal
          ? "Original"
          : "Other"
      });
    });

    return saveResultsToSheet(enriched);

  } catch (e) {
    console.error("PubMed 검색 예외:", e);
    return null;
  }
}

/**
 * PubMed에서 여러 논문의 상세 정보 가져오기
 * @param {string[]} idList - PMID 배열
 * @return {Array} 논문 데이터 배열
 */
/**
 * PubMed에서 논문 데이터를 가져오는 통합 함수
 * 단일 PMID 또는 PMID 배열을 처리할 수 있음
 * 
 * @param {string|string[]} pmidInput - 단일 PMID 또는 PMID 배열
 * @param {Object} options - 옵션 객체 (상세 정보 지정 등)
 * @return {Object|Array} 단일 객체 또는 논문 데이터 배열
 */
function fetchPubMedData(pmidInput, options = {}) {
  const defaults = {
    includeAbstract: true, // 초록 포함 여부
    includeAuthors: true,  // 저자 정보 포함 여부
    detailedFormat: false, // true: 객체 형식 반환, false: 배열 형식 반환
    maxRetries: CONFIG.MAX_RETRIES
  };
  
  const settings = { ...defaults, ...options };
  const isSinglePmid = typeof pmidInput === 'string';
  const pmids = isSinglePmid ? [pmidInput] : pmidInput;
  
  try {
    console.log(`PubMed 데이터 가져오기 시작: ${isSinglePmid ? '단일 PMID' : pmids.length + '개 PMID'}`);
    
    if (pmids.length === 0) {
      console.warn('PMID가 제공되지 않았습니다.');
      return isSinglePmid ? {} : [];
    }
    
    // XML 형식으로 상세 정보 가져오기
    const ids = pmids.join(",");
    const detailUrl = `https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pubmed&id=${ids}&retmode=xml`;
    
    // 재시도 로직 적용
    let attempt = 0;
    let xmlResponse;
    
    while (attempt < settings.maxRetries) {
      try {
        xmlResponse = UrlFetchApp.fetch(detailUrl).getContentText();
        break;
      } catch (fetchError) {
        attempt++;
        if (attempt >= settings.maxRetries) throw fetchError;
        
        // 지수 백오프
        const delay = CONFIG.RETRY_DELAY * Math.pow(CONFIG.RETRY_MULTIPLIER, attempt - 1);
        Utilities.sleep(delay);
        console.log(`PubMed API 재시도 ${attempt}/${settings.maxRetries}, 대기: ${delay}ms`);
      }
    }
    
    const document = XmlService.parse(xmlResponse);
    const root = document.getRootElement();
    const articles = root.getChildren('PubmedArticle');
    
    console.log(`XML 파싱 완료, ${articles.length}개 논문 처리 중...`);
    
    // 결과 저장 배열
    const results = [];
    
    articles.forEach(article => {
      try {
        const citation = article.getChild('MedlineCitation');
        if (!citation) {
          console.log('MedlineCitation 요소를 찾을 수 없습니다.');
          return;
        }
        
        const articleNode = citation.getChild('Article');
        if (!articleNode) {
          console.log('Article 요소를 찾을 수 없습니다.');
          return;
        }
        
        // 제목 추출
        const titleNode = articleNode.getChild('ArticleTitle');
        const articleTitle = titleNode ? titleNode.getText() : 'No Title';
        
        // 저널 추출
        const journalNode = articleNode.getChild('Journal');
        const journalTitleNode = journalNode ? journalNode.getChild('Title') : null;
        const journal = journalTitleNode ? journalTitleNode.getText() : 'No Journal';
        
        // 출판 날짜 정보 추출
        const journalIssueNode = journalNode ? journalNode.getChild('JournalIssue') : null;
        const pubDateNode = journalIssueNode ? journalIssueNode.getChild('PubDate') : null;
        const pubYear = pubDateNode && pubDateNode.getChild('Year') ? pubDateNode.getChild('Year').getText() : 'NA';
        const pubMonth = pubDateNode && pubDateNode.getChild('Month') ? pubDateNode.getChild('Month').getText() : 'NA';
        const pubDay = pubDateNode && pubDateNode.getChild('Day') ? pubDateNode.getChild('Day').getText() : 'NA';
        
        // 날짜 형식 구성
        let pubDate = pubYear;
        if (pubMonth !== 'NA') pubDate += ` ${pubMonth}`;
        if (pubDay !== 'NA') pubDate += ` ${pubDay}`;
        
        // PMID 추출
        const pmid = citation.getChild('PMID') ? citation.getChild('PMID').getText() : 'No PMID';
        
        // 출판 유형 추출
        const pubTypeListNode = articleNode.getChild('PublicationTypeList');
        let pubTypeList = '';
        
        if (pubTypeListNode) {
          const pubTypes = pubTypeListNode.getChildren('PublicationType');
          pubTypeList = pubTypes.map(pt => pt.getText()).join(", ");
        }
        
        // 초록 추출 (설정에 따라)
        let abstract = '';
        if (settings.includeAbstract) {
          const abstractNode = articleNode.getChild('Abstract');
          if (abstractNode) {
            const abstractTextNodes = abstractNode.getChildren('AbstractText');
            abstract = abstractTextNodes.map(node => {
              const label = node.getAttribute('Label');
              const text = node.getText();
              return label ? `${label}: ${text}` : text;
            }).join("\n");
          }
        }
        
        // 저자 정보 추출 (설정에 따라)
        let authors = '';
        if (settings.includeAuthors) {
          const authorListNode = articleNode.getChild('AuthorList');
          if (authorListNode) {
            const authorNodes = authorListNode.getChildren('Author');
            authors = authorNodes.map(author => {
              const lastName = author.getChild('LastName') ? author.getChild('LastName').getText() : '';
              const initials = author.getChild('Initials') ? author.getChild('Initials').getText() : '';
              return lastName + (initials ? ` ${initials}` : '');
            }).join(", ");
          }
        }
        
        // 결과 형식에 따라 반환 데이터 구성
        if (settings.detailedFormat) {
          // 객체 형식 (키-값 페어)
          results.push({
            title: articleTitle,
            journal: journal,
            pubDate: pubDate,
            authors: authors,
            pmid: pmid,
            publicationType: pubTypeList,
            abstract: abstract
          });
        } else {
          // 배열 형식 (순서가 중요)
          const resultRow = [articleTitle, journal, pubDate, authors, pmid, pubTypeList, abstract];
          // 옵션에 따라 추가 필드 포함         
          results.push(resultRow);
        }
        
      } catch (articleError) {
        console.error('개별 논문 파싱 중 오류:', articleError);
      }
    });
    
    console.log(`${results.length}개 논문의 데이터 처리 완료`);
    
    // 단일 PMID인 경우 첫 번째 결과만 반환, 그렇지 않으면 전체 배열 반환
    return isSinglePmid ? (results[0] || (settings.detailedFormat ? {} : [])) : results;
    
  } catch (error) {
    console.error('PubMed 데이터 가져오기 실패:', error);
    throw error;
  }
}

// ===== 4. 데이터 저장 함수 =====

/**
 * 논문 데이터를 새 스프레드시트에 저장
 * @param {Array} data - 논문 데이터 배열
 */

// 1. saveResultsToSheet 함수 수정 - 스프레드시트 객체 반환

function saveResultsToSheet(data) {
  try {
    const today = new Date();
    const formattedDate = Utilities.formatDate(today, "GMT+9", "yyyyMMdd");
    const fileName = `journal_crawl_db_${formattedDate}`;

    const spreadsheet = SpreadsheetApp.create(fileName);
    const sheet = spreadsheet.getActiveSheet();
    sheet.setName("journal_crawl_db");

    // ✅ 8열 헤더 (articleCategory 포함)
    const headers = ["Title", "Journal", "Date", "Authors", "PMID", "Publication Type", "Abstract", "ArticleCategory"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    if (!Array.isArray(data) || data.length === 0) {
      console.log("저장할 데이터 없음");
      SpreadsheetApp.flush();
      return spreadsheet;
    }

    // ✅ 핵심: item이 (1) 배열, (2) '0'~'6' 키를 가진 객체, (3) 일반 객체인 경우 모두 대응
    const rows = data.map(item => {
      // (1) 이미 배열 형태로 온 경우
      if (Array.isArray(item)) {
        const row7 = item.slice(0, 7);
        const category = item[7] || ""; // 혹시 8번째가 이미 있으면
        return row7.concat(category);
      }

      // (2) '0'~'6' 키를 가진 객체 형태 (현재 로그 케이스)
      const hasIndexedKeys = item && Object.prototype.hasOwnProperty.call(item, "0");
      if (hasIndexedKeys) {
        const row7 = [item["0"], item["1"], item["2"], item["3"], item["4"], item["5"], item["6"]]
          .map(v => (v === undefined || v === null) ? "" : String(v));
        const category = (item.articleCategory === undefined || item.articleCategory === null)
          ? ""
          : String(item.articleCategory);
        return row7.concat(category);
      }

      // (3) 일반 객체(필드명 기반)로 오는 경우(혹시 나중에 fetchPubMedData를 바꾸면 이쪽이 동작)
      const pubTypes = item.publicationTypes || item.pubTypes || "";
      const authors = Array.isArray(item.authors) ? item.authors.join(", ") : (item.authors || "");
      const row7 = [
        item.title || "",
        item.journal || item.source || "",
        item.edat || item.pubdate || item.date || "",
        authors,
        item.pmid || "",
        Array.isArray(pubTypes) ? pubTypes.join("; ") : String(pubTypes || ""),
        item.abstract || item.abstractText || ""
      ].map(v => (v === undefined || v === null) ? "" : String(v));

      const category = item.articleCategory ? String(item.articleCategory) : "";
      return row7.concat(category);
    });

    // 길이 보정(혹시라도)
    const normalizedRows = rows.map(r => {
      if (r.length === headers.length) return r;
      if (r.length > headers.length) return r.slice(0, headers.length);
      return r.concat(Array(headers.length - r.length).fill(""));
    });

    sheet.getRange(2, 1, normalizedRows.length, headers.length).setValues(normalizedRows);

    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);

    SpreadsheetApp.flush();

    console.log(`스프레드시트 "${fileName}" 생성 완료, ${normalizedRows.length}개 저장됨`);
    console.log(`URL: ${spreadsheet.getUrl()}`);

    return spreadsheet;
  } catch (error) {
    console.error("스프레드시트 저장 실패:", error);
    throw error;
  }
}

/**
 * GPT API 호출 함수
 * @param {string} prompt - 프롬프트 텍스트
 * @return {string} GPT 응답 텍스트
 */

function callGPT(prompt, isJsonMode, forcedApiKey) {
  const url = "https://api.openai.com/v1/chat/completions";

  // 사용할 API 키 결정
  const apiKey = forcedApiKey || CONFIG.OPENAI_API_KEY_SUMMARY;

  // ✅ 용도에 따라 다른 토큰 제한 사용
  const maxTokens = isJsonMode
    ? CONFIG.MAX_COMPLETION_TOKENS_SCORING  // 점수 평가: 500 토큰
    : CONFIG.MAX_COMPLETION_TOKENS_SUMMARY; // 요약 생성: 1500 토큰

  // ✅ DEBUG: 프롬프트 정보 로깅
  const taskType = isJsonMode ? "SCORING" : "SUMMARY";
  console.log(`━━━━━━━━━━ ${taskType} REQUEST START ━━━━━━━━━━`);
  console.log(`  - Task: ${taskType}`);
  console.log(`  - Prompt length: ${prompt.length} chars`);
  console.log(`  - Max completion tokens: ${maxTokens}`);
  console.log(`  - Reasoning effort: ${CONFIG.REASONING_EFFORT}`);
  console.log(`  - Model: ${CONFIG.GPT_MODEL}`);
  console.log(`  - Using API key: ${apiKey === CONFIG.OPENAI_API_KEY_SCORING ? 'SCORING' : 'SUMMARY'}`);

  // ✅ OPTIMIZATION: System message로 공통 규칙 분리
  const systemMessage = isJsonMode
    ? "You analyze medical papers for keywords. Return only valid JSON."
    : `You are a medical expert summarizing papers for Korean clinicians. Rules:
- Use medical terms in English (cytokines, drugs, diseases)
- Summarize in Korean, max 1200 chars
- Highlight key terms: <b>drug names</b>, <b>cytokines</b>, <b>trial names</b>
- Authors: first 3 + corresponding author et al.
- No hallucination - use only provided info, mark missing info as "-"`;

  const payload = {
    model: CONFIG.GPT_MODEL,
    messages: [
      { role: "system", content: systemMessage },
      { role: "user", content: prompt }
    ],
    max_completion_tokens: maxTokens,
    reasoning_effort: CONFIG.REASONING_EFFORT  // ✅ reasoning 강도 제한
  };

  // JSON 모드 지원 (gpt-4-turbo 이상에서만 작동)
  if (isJsonMode) {
    payload.response_format = { type: "json_object" };
  }

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${apiKey}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  // ✅ 재시도 로직 추가
  let lastError = null;
  for (let attempt = 1; attempt <= CONFIG.MAX_RETRIES; attempt++) {
    try {
      if (attempt > 1) {
        const delay = CONFIG.RETRY_DELAY * Math.pow(CONFIG.RETRY_MULTIPLIER, attempt - 2);
        console.log(`GPT API 재시도 ${attempt}/${CONFIG.MAX_RETRIES}, 대기: ${delay}ms`);
        Utilities.sleep(delay);
      }
      const response = UrlFetchApp.fetch(url, options);
      const status = response.getResponseCode();
      const body = response.getContentText();

      if (status !== 200) {
        console.error(`GPT Error (Status ${status}):`, body);
        const json = JSON.parse(body);
        if (json.error) {
          throw new Error(json.error.message);
        }
        throw new Error(`API returned status ${status}`);
      }

      const json = JSON.parse(body);

      // ✅ 디버깅: 응답 구조 확인
      console.log("DEBUG - API Response Structure (attempt " + attempt + "):");
      console.log("  - max_completion_tokens:", maxTokens);
      console.log("  - reasoning_effort:", CONFIG.REASONING_EFFORT);
      console.log("  - has choices:", !!json.choices);
      console.log("  - choices length:", json.choices ? json.choices.length : 0);
      if (json.choices && json.choices.length > 0) {
        const msg = json.choices[0].message;
        console.log("  - has message:", !!msg);
        console.log("  - has content:", msg && msg.content ? true : false);
        console.log("  - content length:", msg && msg.content ? msg.content.length : 0);
        console.log("  - finish_reason:", json.choices[0].finish_reason);
      }

      // Chat Completions API 표준 응답 파싱
      if (json.choices && json.choices.length > 0) {
        const choice = json.choices[0];
        const message = choice.message;

        // ✅ 성공 로깅
        if (message && message.content) {
          console.log(`━━━━━━━━━━ ${taskType} RESPONSE SUCCESS ━━━━━━━━━━`);
          console.log(`  - Finish reason: ${choice.finish_reason}`);
          console.log(`  - Response length: ${message.content.length} chars`);
          console.log(`  - Usage: ${json.usage ? JSON.stringify(json.usage) : 'N/A'}`);
          console.log(`━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
          return message.content.trim();
        }

        // ✅ 컨텐츠가 없는 경우 상세 로깅
        console.error(`━━━━━━━━━━ ${taskType} RESPONSE FAILED ━━━━━━━━━━`);
        console.error("GPT response missing content (attempt " + attempt + "):");
        console.error("  - finish_reason:", choice.finish_reason);
        console.error("  - message exists:", !!message);
        console.error("  - message.content:", message ? message.content : "no message");
        console.error("  - usage:", json.usage ? JSON.stringify(json.usage) : 'N/A');
        console.error("  - full message object:", JSON.stringify(message));
        console.error(`━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);

        // 빈 응답을 에러로 처리하여 재시도
        if (attempt < CONFIG.MAX_RETRIES) {
          lastError = new Error(`Empty content, finish_reason: ${choice.finish_reason}`);
          continue; // 재시도
        }
      }

      console.warn("GPT response has no recognizable content. Full response:", JSON.stringify(json).slice(0, 500));
      return "";

    } catch (error) {
      lastError = error;
      console.error("GPT API 호출 오류 (attempt " + attempt + "):", error);
      console.error("Error stack:", error.stack);

      // 마지막 시도가 아니면 계속
      if (attempt < CONFIG.MAX_RETRIES) {
        continue;
      }
    }
  }

  // 모든 재시도 실패
  console.error("모든 GPT API 재시도 실패:", lastError);
  return null;
}

/**
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - 요약 논문 결과 스프레드시트
 * @return {string} GPT 요약 결과
 */
// 3. summarizePubMedArticlesWithGPT 함수 수정 - 스프레드시트 매개변수 추가

function summarizePubMedArticlesWithGPT(spreadsheet) {
  console.log("논문 GPT 요약 작업 시작...");

  // 스프레드시트 / 시트 확인
  let sheet;
  try {
    // 매개변수로 받은 스프레드시트 사용
    sheet = spreadsheet.getSheetByName('journal_crawl_db');
    if (!sheet) {
      console.error("'journal_crawl_db' 시트를 찾을 수 없습니다.");
      return "시트 없음";
    }
  } catch (error) {
    console.error("시트 접근 오류:", error);
    return "시트 접근 오류" + error.message;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    console.error("데이터가 없습니다.");
    return "데이터 없음";
  }

  // ✅ FIX: 먼저 lastCol 저장 (Summary 열 추가 전)
  const lastCol = sheet.getLastColumn();

  // 요약 결과 열 추가
  const targetCol = lastCol + 1;
  sheet.getRange(1, targetCol).setValue(MESSAGES.SUMMARY_HEADER);

  // ✅ FIX: Summary 열 추가 후에 헤더 읽기 (정확한 동기화)
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // 필요한 열 인덱스 찾기
  const titleIdx = headers.indexOf("Title");
  const journalIdx = headers.indexOf("Journal");
  const dateIdx = headers.indexOf("Date");
  const authorsIdx = headers.indexOf("Authors");
  const pmidIndex = headers.indexOf("PMID");
  const pubtypeIdx = headers.indexOf("Publication Type");
  const abstractIdx = headers.indexOf("Abstract");

  // 필수 열 확인
  if (titleIdx === -1 || abstractIdx === -1 || pmidIndex === -1) {
    return "필수 열(제목, PMID, 초록)을 찾을 수 없습니다.";
  }

// ✅ FIX: 이제 headers와 dataRows의 열 개수가 일치
  const dataRows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  let successCount = 0;
  let failCount = 0;
  const startTime = new Date().getTime();
  //let lastSaveTime = startTime;
  //console.log(`총 ${pmids.length}개의 PMID에 대해 요약을 진행합니다.`);
  

for (let i = 0; i < dataRows.length; i++) {
    const rowIndex = i + 2;
    const row = dataRows[i];

    const pmid = row[pmidIndex] || "";
    const title = row[titleIdx] || "";
    const journal = row[journalIdx] || "";
    const date = row[dateIdx] || "";
    const authors = row[authorsIdx] || "";
    const pubtype = row[pubtypeIdx] || "";
    const abstract = row[abstractIdx] || "";
    
    // ✅ FIX: 점수 필터링 확인 (디버깅 로그 추가)
    const includedIdx = headers.indexOf("Included");
    if (includedIdx !== -1) {
       const isIncluded = row[includedIdx];
       console.log(`DEBUG Row ${rowIndex}: includedIdx=${includedIdx}, isIncluded="${isIncluded}", type=${typeof isIncluded}`);

       // ✅ 빈 문자열이나 undefined도 처리
       if (isIncluded !== "O" && isIncluded !== "") {
          console.log(`Row ${rowIndex}: Included != "O" (value: "${isIncluded}"), Skip summary.`);
          continue;
       }

       // ✅ Included가 비어있으면 요약 진행 (scoreAndFilterPapers 실행 안 했을 경우)
       if (isIncluded === "") {
          console.log(`Row ${rowIndex}: Included is empty, proceeding with summary (no filtering applied)`);
       }
    } else {
       console.log(`Row ${rowIndex}: No "Included" column found, proceeding with summary (all papers)`);
    }

    if (!pmid) {
      console.log(`행 ${rowIndex}: PMID가 없어서 스킵`);
      sheet.getRange(rowIndex, targetCol).setValue(MESSAGES.NO_PMID);
      failCount++;
      continue;
    }

    if (!abstract) {
      console.log(`행 ${rowIndex}: 초록이 없어서 스킵`);
      // ✅ FIX: Date 타입 체크 후 포맷팅
      let formattedDate = "";
      if (date instanceof Date) {
        formattedDate = Utilities.formatDate(date, "GMT+9", "yyyy년 MM월 dd일");
      } else if (date) {
        formattedDate = String(date);
      }
      const message = `초록이 없습니다.\n 📅: ${formattedDate} \n 📒: ${journal}\n👥: ${authors}`;
      sheet.getRange(rowIndex, targetCol).setValue(message);

      failCount++;
      continue;
    }
    

    try {
           
      // ✅ OPTIMIZED: 프롬프트 압축 (토큰 50% 절약)
        const prompt = `Summarize this paper in Korean:

Title: ${title}
Journal: ${journal}
Date: ${date}
Authors: ${authors}
PMID: ${pmid}
Abstract: ${abstract}

Output format:
• 🗓️: [date]
• 📒: [journal]
• 👤: [first 3 + corr. author et al.]
• 주요 대상 질환: [disease]
• 연구 방법: [study type + brief method, 2-3 lines]
• 🎯: [key results, 2-3 lines]
• 임상적용 가능성: [clinical implications]
• 제한점: [limitations]
• Tag: #disease #mechanism #keyword #drug/study_type

Write in Korean. Use English for medical terms.`;
      
      let summary = callGPT(prompt, false, CONFIG.OPENAI_API_KEY_SUMMARY); // 요약용 키 사용
      summary = normalizeBoldMarkup(summary);
      
      // 요약 결과 저장
      sheet.getRange(rowIndex, targetCol).setValue(summary);
      successCount++;
      console.log("DEBUG prompt length:", prompt.length, "PMID:", pmid);


      // 너무 빠른 API 호출 방지
      
    } catch (error) {
      console.error(`PMID ${pmid} 처리 오류:`, error);
      sheet.getRange(rowIndex, targetCol).setValue(`오류: ${error.message}`);
      failCount++;
      
    }

    // 처리 속도 측정 및 남은 시간 예상
    const itemsDone = i + 1;
    const currentTime = new Date().getTime();
    const elapsedSec = (currentTime - startTime) / 1000;
    const avgTimePerItem = elapsedSec / itemsDone;
    const remainingItems = dataRows.length - itemsDone;
    const estimatedRemaining = (avgTimePerItem * remainingItems).toFixed(1);
    console.log(`[${itemsDone}/${dataRows.length}] PMID: ${pmid}, 평균속도(건/초): ${(1/avgTimePerItem).toFixed(2)}, 예상 남은 시간: ${Math.floor(estimatedRemaining/60)}분 ${Math.floor(estimatedRemaining%60)}초`);
  }

  // 마지막 저장
  try {
    SpreadsheetApp.flush();
  } catch (flushErr) {
    console.error("마지막 flush 오류:", flushErr);
  }

  const resultMsg = `요약 작업 완료! 성공: ${successCount}, 실패: ${failCount}`;
  console.log(resultMsg);
  return resultMsg;
}