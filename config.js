/**
 * 환경 설정 관리
 * 민감한 정보는 보안을 위해 Google Apps Script Properties Service를 사용합니다.
 *
 * 설정 방법:
 * 1. 스크립트 에디터에서 프로젝트 설정 > 스크립트 속성 추가
 * 2. 다음 키들을 추가:
 *    - OPENAI_API_KEY_SCORING
 *    - OPENAI_API_KEY_SUMMARY
 *    - PUBMED_API_KEY
 *    - NOTION_API_KEY
 *    - NOTION_PAPERS_DB_ID
 *    - NOTION_NEWSLETTER_DB_ID
 *    - EMAIL_RECIPIENTS
 *    - EMAIL_TO_PRIMARY
 */

// ===== API 키 및 민감 정보 설정 =====
function getSecretConfig() {
  const props = PropertiesService.getScriptProperties();

  return {
    OPENAI_API_KEY_SCORING: props.getProperty('OPENAI_API_KEY_SCORING') || '',
    OPENAI_API_KEY_SUMMARY: props.getProperty('OPENAI_API_KEY_SUMMARY') || '',
    PUBMED_API_KEY: props.getProperty('PUBMED_API_KEY') || '',
    NOTION_API_KEY: props.getProperty('NOTION_API_KEY') || '',
    NOTION_PAPERS_DB_ID: props.getProperty('NOTION_PAPERS_DB_ID') || '',
    NOTION_NEWSLETTER_DB_ID: props.getProperty('NOTION_NEWSLETTER_DB_ID') || '',
    EMAIL_RECIPIENTS: props.getProperty('EMAIL_RECIPIENTS') || '',
    EMAIL_TO_PRIMARY: props.getProperty('EMAIL_TO_PRIMARY') || 'allergy@aumc.ac.kr'
  };
}

// ===== 비즈니스 설정 =====
const BUSINESS_CONFIG = {
  // 검색 설정
  DAYS_RANGE: 7,                    // 검색 기간 (일)
  MAX_RESULTS: 200,                 // 최대 검색 결과 수

  // 점수 설정
  MIN_RELEVANCE_SCORE: 3,           // 최소 관련성 점수
  MAX_INCLUSION: 15,                // 최대 포함 논문 수 (Top N)

  // 이메일 설정
  EMAIL_BATCH_SIZE: 15,             // 한 이메일에 포함할 논문 수
  EMAIL_MAX_LENGTH: 100000,         // 이메일 최대 길이 제한
  EMAIL_SUBJECT_PREFIX: '[CU-Ana Newsletter]',

  // API 설정
  GPT_MODEL: 'gpt-5-mini',          // GPT 모델명 (gpt-5-mini, gpt-4o-mini, gpt-4-turbo 등)
  MAX_RETRIES: 3,                   // API 재시도 횟수
  MAX_COMPLETION_TOKENS_SCORING: 500,   // GPT 점수 평가용 최대 토큰 (간단한 JSON)
  MAX_COMPLETION_TOKENS_SUMMARY: 1500,  // GPT 요약 생성용 최대 토큰 (복잡한 한국어 요약)
  REASONING_EFFORT: 'minimal',      // GPT-5 reasoning 강도 (minimal, low, medium, high)
  RETRY_DELAY: 1000,                // 기본 재시도 대기 시간(ms)
  RETRY_MULTIPLIER: 2,              // 재시도 지수 백오프 승수

  // 토큰 최적화 설정
  USE_SYSTEM_MESSAGE: true,         // System message 사용 여부 (토큰 30-40% 절약)
  USE_COMPRESSED_PROMPTS: true      // 압축된 프롬프트 사용 (영어 기반, 토큰 50% 절약)
};

// ===== 저널 목록 =====
const JOURNALS = [
  // === Allergy & Clinical Immunology ===
  "The Journal of allergy and clinical immunology",
  "J Allergy Clin Immunol",
  "Allergy",
  "The journal of allergy and clinical immunology. In practice",
  "J Allergy Clin Immunol Pract",
  "Clinical and experimental allergy",
  "Clin Exp Allergy",
  "Annals of allergy, asthma & immunology",
  "Ann Allergy Asthma Immunol",
  "Allergology international",
  "Allergol Int",
  "The World Allergy Organization journal",
  "World Allergy Organ J",
  "Journal of Investigational Allergology and Clinical Immunology",
  "J Investig Allergol Clin Immunol",
  "Allergy, Asthma and Immunology Research",
  "Allergy Asthma Immunol Res",
  "Allergy, asthma, and clinical immunology",
  "Allergy Asthma Clin Immunol",
  "Allergy and asthma proceedings",
  "Allergy Asthma Proc",
  "Allergy, Asthma and Respiratory Disease",
  "Asia pacific Allergy",
  "Clinical and Translational Immunology",
  "Immunological Reviews",
  "Frontiers in Immunology",
  "Clinical and Translational Allergy",
  "Clin Transl Allergy",
  "Pediatric allergy and immunology",
  "Pediatr Allergy Immunol",
  "Current opinion in allergy and clinical immunology",
  "Curr Opin Allergy Clin Immunol",
  "International archives of allergy and immunology",
  "Int Arch Allergy Immunol",

  // === Dermatology ===
  "Journal of the American Academy of Dermatology",
  "J Am Acad Dermatol",
  "JAMA dermatology",
  "JAMA Dermatol",
  "The British journal of dermatology",
  "Br J Dermatol",
  "The Journal of investigative dermatology",
  "J Invest Dermatol",
  "Journal of the European Academy of Dermatology and Venereology",
  "J Eur Acad Dermatol Venereol",
  "Contact dermatitis",
  "Acta dermato-venereologica",
  "Acta Derm Venereol",
  "Journal of dermatological science",
  "J Dermatol Sci",
  "Dermatology",

  // === General Medical ===
  "The New England journal of medicine",
  "N Engl J Med",
  "NEJM evidence",
  "NEJM Evid",
  "The Lancet",
  "JAMA",
  "BMJ"
];

// ===== 출판 유형 =====
const PUB_TYPES = [
  "Meta-Analysis",
  "Randomized Controlled Trial",
  "Review",
  "Systematic Review",
  "Original article"
];

// ===== 메시지 상수 =====
const MESSAGES = {
  NO_ABSTRACT: "초록을 찾을 수 없습니다.",
  NO_PMID: "PMID를 찾을 수 없습니다.",
  SUMMARY_HEADER: "GPT 요약",
  COLUMN_HEADERS: ["Title", "Journal", "Year", "PMID", "Publication Type", "GPT 요약"]
};

// ===== 통합 설정 객체 =====
function getConfig() {
  const secretConfig = getSecretConfig();

  return {
    // 민감 정보 (Properties에서 로드)
    ...secretConfig,

    // 비즈니스 설정
    ...BUSINESS_CONFIG,

    // 정적 데이터
    JOURNALS: JOURNALS,
    PUB_TYPES: PUB_TYPES
  };
}

// 하위 호환성을 위한 전역 CONFIG 객체
const CONFIG = getConfig();