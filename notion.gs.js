/******************************
 * Notion Newsletter 관련 함수
 * Notion API 헤더는 utils.gs의 getNotionHeaders() 사용
 ******************************/

function findNewsletterPageIdByDate(dateISO) {
  const url = `https://api.notion.com/v1/databases/${CONFIG.NOTION_NEWSLETTER_DB_ID}/query`;

  const payload = {
    page_size: 1,
    filter: {
      property: "Date",
      date: { equals: dateISO }
    }
  };

  const res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: getNotionHeaders(),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) throw new Error("Notion newsletter query failed: " + res.getContentText());

  const json = JSON.parse(res.getContentText());
  return (json.results && json.results.length > 0) ? json.results[0].id : null;
}


function createNewsletterPage(dateISO, subject) {
  const payload = {
    parent: { database_id: CONFIG.NOTION_NEWSLETTER_DB_ID },
    properties: {
      "Title": { title: [{ text: { content: `${dateISO} Newsletter` } }] },
      "Date": { date: { start: dateISO } },
      "Subject": { rich_text: [{ text: { content: String(subject || "") } }] }
    }
  };

  const res = UrlFetchApp.fetch("https://api.notion.com/v1/pages", {
    method: "post",
    contentType: "application/json",
    headers: getNotionHeaders(),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    throw new Error("Notion newsletter create failed: " + res.getContentText());
  }

  return JSON.parse(res.getContentText()).id;
}


/******************************
 * 3) 페이지 본문 덮어쓰기: children listing / delete / clear
 ******************************/
function listBlockChildrenAll(blockId) {
  const all = [];
  let cursor = null;

  while (true) {
    const url = `https://api.notion.com/v1/blocks/${blockId}/children?page_size=100` +
      (cursor ? `&start_cursor=${encodeURIComponent(cursor)}` : "");

    const res = UrlFetchApp.fetch(url, {
      method: "get",
      headers: getNotionHeaders(),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      throw new Error("Notion get children failed: " + res.getContentText());
    }

    const json = JSON.parse(res.getContentText());
    const results = json.results || [];
    for (const b of results) all.push(b);

    if (!json.has_more) break;
    cursor = json.next_cursor;
  }

  return all;
}


function deleteBlock(blockId) {
  const url = `https://api.notion.com/v1/blocks/${blockId}`;

  const res = UrlFetchApp.fetch(url, {
    method: "delete",
    headers: getNotionHeaders(),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    throw new Error("Notion delete block failed: " + res.getContentText());
  }
}

function clearPageBody(pageId) {
  const children = listBlockChildrenAll(pageId);
  for (const b of children) {
    deleteBlock(b.id);
    Utilities.sleep(120); // rate limit 완화
  }
}

/******************************
 * 4) Email HTML -> Notion blocks 변환
 * - h4 -> heading_2
 * - p  -> paragraph
 * - hr -> divider
 * - br -> 줄바꿈
 * - b  -> bold annotation
 * - a  -> link
 ******************************/
function htmlInlineToRichText(s) {
  let text = String(s || "");

  // 링크 토큰화: [[A|url|label]]
  text = text.replace(/<a\s+[^>]*href="([^"]+)"[^>]*>(.*?)<\/a>/gi, '[[A|$1|$2]]');
  // 볼드 토큰화
  text = text.replace(/<b>/gi, '[[B]]').replace(/<\/b>/gi, '[[/B]]');

  const rich = [];

  while (text.length > 0) {
    const aIdx = text.indexOf('[[A|');
    const bIdx = text.indexOf('[[B]]');
    const next = [aIdx, bIdx].filter(i => i >= 0).sort((x, y) => x - y)[0];

    if (next === undefined) {
      if (text) rich.push({ type: "text", text: { content: text } });
      break;
    }

    if (next > 0) {
      rich.push({ type: "text", text: { content: text.slice(0, next) } });
      text = text.slice(next);
    }

    // 링크
    if (text.startsWith('[[A|')) {
      const end = text.indexOf(']]', 4);
      if (end === -1) {
        rich.push({ type: "text", text: { content: text } });
        break;
      }
      const body = text.slice(4, end); // url|label
      const parts = body.split('|');
      const url = parts[0] || "";
      const label = (parts.slice(1).join('|') || url).replace(/\[\[B\]\]|\[\[\/B\]\]/g, "");
      rich.push({ type: "text", text: { content: label, link: { url } } });
      text = text.slice(end + 2);
      continue;
    }

    // 볼드
    if (text.startsWith('[[B]]')) {
      const end = text.indexOf('[[/B]]', 5);
      if (end === -1) {
        rich.push({ type: "text", text: { content: text } });
        break;
      }
      const inner = text.slice(5, end);
      rich.push({
        type: "text",
        text: { content: inner },
        annotations: { bold: true }
      });
      text = text.slice(end + 6);
      continue;
    }
  }

  return rich.length ? rich : [{ type: "text", text: { content: "" } }];
}

function htmlEmailToNotionBlocks(emailHtml) {
  let html = String(emailHtml || "");

  // div는 경계로만 사용
  html = html.replace(/<div[^>]*>/gi, "\n").replace(/<\/div>/gi, "\n");
  html = html.replace(/<span[^>]*>/gi, "").replace(/<\/span>/gi, "");
  html = html.replace(/&nbsp;/gi, " ");
  html = html.replace(/\r/g, "");

  // hr/h4/p/br 처리
  html = html.replace(/<hr[^>]*>/gi, "\n[[DIVIDER]]\n");
  html = html.replace(/<h4[^>]*>/gi, "\n[[H4]]\n").replace(/<\/h4>/gi, "\n[[/H4]]\n");
  html = html.replace(/<p[^>]*>/gi, "\n[[P]]\n").replace(/<\/p>/gi, "\n[[/P]]\n");
  html = html.replace(/<br\s*\/?>/gi, "\n");

  // b/a 제외 태그 제거
  html = html.replace(/<(?!\/?(b|a)\b)[^>]+>/gi, "");

  const lines = html
    .split("\n")
    .map(s => s.trim())
    .filter(s => s.length > 0);

  const blocks = [];

  let mode = null;     // null | "H4" | "P"
  let buffer = [];

  const flush = () => {
    const text = buffer.join("\n").trim();
    buffer = [];
    if (!text) return;

    if (mode === "H4") {
      blocks.push({
        object: "block",
        type: "heading_2",
        heading_2: { rich_text: htmlInlineToRichText(text) }
      });
    } else if (mode === "P") {
      blocks.push({
        object: "block",
        type: "paragraph",
        paragraph: { rich_text: htmlInlineToRichText(text) }
      });
    } else {
      // 일반 텍스트
      blocks.push({
        object: "block",
        type: "paragraph",
        paragraph: { rich_text: htmlInlineToRichText(text) }
      });
    }
  };

  for (const line of lines) {
    // divider는 모드와 무관하게 즉시 처리
    if (line === "[[DIVIDER]]") {
      // 현재 버퍼 비우고 divider 추가
      flush();
      blocks.push({ object: "block", type: "divider", divider: {} });
      continue;
    }

    // 모드 시작/종료 토큰 처리
    if (line === "[[H4]]") { flush(); mode = "H4"; continue; }
    if (line === "[[/H4]]") { flush(); mode = null; continue; }
    if (line === "[[P]]") { flush(); mode = "P"; continue; }
    if (line === "[[/P]]") { flush(); mode = null; continue; }

    // 내용 누적
    buffer.push(line);
  }

  // 마지막 남은 버퍼 처리
  flush();

  return blocks;
}
/******************************
 * 5) blocks append (chunk)
 ******************************/
function appendBlocksInChunks(pageId, blocks) {
  const url = `https://api.notion.com/v1/blocks/${pageId}/children`;
  const CHUNK = 80;

  for (let i = 0; i < blocks.length; i += CHUNK) {
    const chunk = blocks.slice(i, i + CHUNK);

    const res = UrlFetchApp.fetch(url, {
      method: "patch",
      contentType: "application/json",
      headers: getNotionHeaders(),
      payload: JSON.stringify({ children: chunk }),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      throw new Error("Notion append failed: " + res.getContentText());
    }

    Utilities.sleep(250);
  }
}

function upsertNewsletterPageId(dateISO, subject) {
  const existing = findNewsletterPageIdByDate(dateISO);
  if (existing) return existing;
  return createNewsletterPage(dateISO, subject);
}

/******************************
 * 6) 최종: 날짜 1페이지 + 본문 덮어쓰기(HTML -> blocks)
 ******************************/
function overwriteNewsletterBodyByEmailHtml(dateISO, subject, emailHtml) {
  const pageId = upsertNewsletterPageId(dateISO, subject);

  // (선택) 제목/Subject도 매번 업데이트하고 싶으면 여기서 notionUpdatePage 호출 추가 가능
  // 지금은 본문 덮어쓰기에 집중.

  // 1) 기존 본문 제거
  clearPageBody(pageId);

  // 2) 새 본문 블록 생성
  const blocks = htmlEmailToNotionBlocks(emailHtml);

  // 3) 새 본문 append
  appendBlocksInChunks(pageId, blocks);

  return pageId;
}