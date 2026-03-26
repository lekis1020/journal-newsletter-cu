# CU-Ana Journal Newsletter (Apps Script)

두드러기/혈관부종/아나필락시스/비만세포증/식품알레르기 관련 PubMed 논문을 주기적으로 수집하고,
점수화/요약 후 이메일로 발송하는 Google Apps Script 프로젝트입니다.

## 주요 기능

- PubMed 최근 N일 논문 검색
- 논문 relevance 스코어링 (`gpt` 또는 `regex`)
- 상위 논문 GPT 요약 생성
- 결과 스프레드시트 저장
- 이메일 자동 발송

## 파일 구조

- `Code.js` : 검색/스코어링/요약 핵심 로직
- `config.js` : 설정값 + Script Properties 로딩
- `email.js` : 이메일 본문 생성/발송
- `total.js` : 워크플로우 엔트리 함수
- `util.js` : 공통 유틸 함수
- `notion.gs.js` : Notion 연동 유틸(필요 시)
- `appsscript.json` : Apps Script manifest

## 사전 설정 (Script Properties)

Apps Script 프로젝트 설정 > Script properties 에 아래 키를 추가하세요.

- `OPENAI_API_KEY_SCORING`
- `OPENAI_API_KEY_SUMMARY`
- `PUBMED_API_KEY` (선택)
- `NOTION_API_KEY` (Notion 사용 시)
- `NOTION_PAPERS_DB_ID` (Notion 사용 시)
- `NOTION_NEWSLETTER_DB_ID` (Notion 사용 시)
- `EMAIL_RECIPIENTS`
- `EMAIL_TO_PRIMARY`

## 실행 함수

### 1) 전체 실행 (추천)

```javascript
runWeeklyDigestWorkflow()
```

순서:
1. PubMed 검색 + 스프레드시트 생성
2. 점수 계산 및 Included 결정
3. GPT 요약 생성
4. 이메일 발송

### 2) 검색만 테스트

```javascript
testSearchQuery()
```

## 스프레드시트 규칙

- 생성 파일명: `journal_cu_ana_db_YYYYMMDD`
- 주요 시트명: `journal_cu_ana_db`

## 주요 설정값 (`config.js`)

- `DAYS_RANGE`: 검색 기간(일)
- `MAX_RESULTS`: PubMed 최대 검색 수
- `MIN_RELEVANCE_SCORE`: 포함 최소 점수
- `MAX_INCLUSION`: 최종 포함 최대 논문 수
- `SCORING_MODE`: `gpt` 또는 `regex`
- `COMPARE_WITH_GPT_SCORING`: GPT 비교 컬럼 활성화 여부

## 로컬 개발/배포

```bash
# 상태 확인
clasp status

# Apps Script 원격 반영
clasp push
```

> `.claspignore`로 로컬 아티팩트(`.git`, `.omx`, `imports`, 리포트 파일 등)는 푸시에서 제외됩니다.
