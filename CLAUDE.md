# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a **Google Apps Script** project that automates a weekly medical journal newsletter for an allergy research team (알레르기연구팀). It searches PubMed for recent papers on urticaria/angioedema, anaphylaxis, mastocytosis, and food allergy, scores them with GPT, generates Korean-language summaries, and distributes results via email.

## Deployment

This project uses **clasp** (Google Apps Script CLI) for deployment. Files are pushed directly to Google Apps Script.

```bash
clasp push        # Push local files to Apps Script
clasp pull        # Pull remote changes
clasp open        # Open the script in browser
```

The `.clasp.json` maps file extensions: `.js` and `.gs` files are script files, `.json` are config files.

There are no local build, lint, or test commands. All execution happens in the Google Apps Script runtime (V8), which provides global objects like `UrlFetchApp`, `SpreadsheetApp`, `MailApp`, `Utilities`, `PropertiesService`, `XmlService`, and `Logger`.

## Architecture

### Workflow Pipeline (`total.js:runWeeklyDigestWorkflow`)

The main entry point orchestrates a 4-step pipeline:

1. **PubMed Search** (`Code.js:fetchPubMedWeeklyAndSave`) - Searches PubMed via E-utilities API for recent papers matching urticaria/angioedema/anaphylaxis/mastocytosis/food allergy keywords in configured journals, saves to a new Google Spreadsheet
2. **GPT Scoring** (`Code.js:scoreAndFilterPapers`) - Uses GPT to detect cancer/urticaria/anaphylaxis/mastocytosis keywords, calculates relevance scores, marks top 15 as "Included"
3. **GPT Summarization** (`Code.js:summarizePubMedArticlesWithGPT`) - Generates Korean summaries for included papers only
4. **Email** (`email.js:sendSummariesToEmail`) - Sends HTML email newsletter

### File Responsibilities

- **`config.js`** - All configuration: API keys (via `PropertiesService`), business settings (scoring thresholds, model params, journal list), constants
- **`Code.js`** - Core logic: PubMed search/fetch/parse, GPT calls, scoring algorithm, spreadsheet operations
- **`email.js`** - Email composition and sending (HTML formatted newsletter)
- **`total.js`** - Workflow orchestrators: `runWeeklyDigestWorkflow()` (main), `sendNoResultsEmail()`, `sendEmailFromActiveSpreadsheet()`
- **`notion.gs.js`** - Notion Newsletter DB operations: page upsert, HTML-to-Notion-blocks conversion, block CRUD
- **`util.js`** - Shared utilities: date formatting, Notion API helpers, text processing, retry logic

### Scoring System (`Code.js:calculateScore`)

Papers are scored on a point system:
- **Urticaria/Angioedema keywords**: +2 per location (title/abstract), +5 if both
- **Anaphylaxis/Food Allergy keywords**: +2 per location, +5 if both
- **Mastocytosis keywords**: +2 per location, +5 if both
- **Cancer keywords**: -1 per location, -3 if both
- Final = max(urticaria, anaphylaxis, mastocytosis) + cancer penalty
- Threshold: `MIN_RELEVANCE_SCORE` (default 3), top `MAX_INCLUSION` (default 15) papers included

### Key Configuration (`config.js`)

- Secrets are stored in Google Apps Script Properties Service (not in code)
- `BUSINESS_CONFIG` controls search range, scoring thresholds, GPT model/tokens, and email settings
- `JOURNALS` array defines the target journal list for PubMed filtering
- Two separate OpenAI API keys: one for scoring (`OPENAI_API_KEY_SCORING`), one for summarization (`OPENAI_API_KEY_SUMMARY`)

### Data Flow

- **Spreadsheet**: Created fresh each run as `journal_crawl_db_YYYYMMDD`, with sheet name `journal_crawl_db`. Columns are progressively added: base 8 columns from PubMed, then Scores/Final Score/Included/Exclusion Reason from scoring, then GPT Summary
- The `Included` column uses "O"/"X" values (not boolean)

### GPT Integration

- `callGPT()` in `Code.js` handles all OpenAI API calls with retry logic and exponential backoff
- JSON mode is used for scoring (structured keyword detection), plain text for Korean summaries
- System messages are used to reduce token usage
- Model is configurable via `GPT_MODEL` (currently `gpt-5-mini`)
