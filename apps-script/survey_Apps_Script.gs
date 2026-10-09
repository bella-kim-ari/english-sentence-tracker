
/**
 * English Sentence Tracker — Survey API
 *
 * 설문 응답을 검증하고 Google Sheets의 responses 시트에 저장한다.
 *
 * 보안 처리:
 * - Q1~Q8 필수 응답 및 선택값 검증
 * - Q7 최대 2개 선택 및 배타적 선택 검증
 * - 스프레드시트 수식 주입 방어
 * - 전체 요청 빈도 제한 (1분당 60회)
 */

const SURVEY_CONFIG = {
  sheetName: "responses",
  rateLimit: 60,
  rateWindowMs: 60 * 1000,
  maxFeedbackLength: 2000
};

const SURVEY_OPTIONS = {
  source: [
    "구글 검색",
    "네이버 검색",
    "네이버 블로그",
    "기타 블로그·콘텐츠",
    "커뮤니티",
    "지인으로부터 링크 공유받음",
    "기타"
  ],

  study_purpose: [
    "일상 회화",
    "취업·업무",
    "시험·자격증",
    "여행·해외생활",
    "영어 영상·콘텐츠 이해",
    "자기계발",
    "기타"
  ],

  study_methods: [
    "유튜브 등 무료 동영상 플랫폼",
    "영어 학습 앱(듀오링고, 말해보카 등)",
    "온라인 영어 강의·학습 사이트",
    "오프라인 영어학원·수업",
    "전화영어·화상영어",
    "교재·문제집 등 인쇄 자료를 이용한 독학",
    "AI 대화 도구(ChatGPT 등)",
    "기타"
  ],

  study_management: [
    "영상을 시청하고 별도로 기록하지 않는다.",
    "나중에 다시 보려고 영상을 저장하거나 채널을 구독한다.",
    "중요한 문장이나 표현을 종이 노트에 기록한다.",
    "중요한 문장이나 표현을 디지털 형태(메모 앱, 문서, 스프레드시트 등)로 기록한다.",
    "유튜브로 영어 공부를 하지 않는다.",
    "기타"
  ],

  study_device: [
    "PC·노트북",
    "스마트폰",
    "태블릿(아이패드, 갤럭시탭 등)",
    "상황에 따라 여러 기기를 비슷한 비중으로 사용한다.",
    "유튜브로 영어 공부를 하지 않는다."
  ],

  study_pain_points: [
    "공부한 영어 문장을 따로 정리하기 번거롭다.",
    "이전에 공부한 문장을 다시 찾아 복습하기 어렵다.",
    "영상을 시청해도 말하기 연습으로 이어지지 않는다.",
    "배운 문장을 직접 활용해 작문하기 어렵다.",
    "꾸준히 학습하거나 진도를 관리하기 어렵다.",
    "특별히 불편한 점은 없다.",
    "유튜브로 영어 공부를 해본 적이 없다.",
    "기타"
  ],

  interested_features: [
    "유튜브 자막에서 공부할 영어 문장을 자동으로 정리하는 기능",
    "공부한 영어 문장을 한곳에 모아 저장하는 기능",
    "영어 문장을 반복해서 말하고, 최종 녹음까지 학습 기록으로 관리하는 기능",
    "한국어 뜻을 보고 직접 영어 문장을 작성하는 기능",
    "오늘 공부할 문장과 학습 완료 여부를 관리하는 기능",
    "지금까지 공부한 문장 수와 학습 현황을 확인하는 기능",
    "특별히 도움이 될 것 같은 기능이 없다"
  ],

  usage_intent: ["1", "2", "3", "4", "5"]
};

/**
 * 설문 응답 수신 및 저장
 */
function doPost(e) {
  try {
    // 요청 빈도 제한
    if (!checkRateLimit_()) {
      return createResponse_("rate_limited");
    }

    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SURVEY_CONFIG.sheetName);

    if (!sheet) {
      throw new Error("responses 시트를 찾을 수 없습니다.");
    }

    const p = (e && e.parameter) || {};

    // 서버 측 입력값 검증
    const validation = validateSurvey_(p);

    if (!validation.valid) {
      console.warn("설문 검증 실패: " + validation.reason);
      return createResponse_("invalid_request");
    }

    // 사용자 입력값을 일반 텍스트로 안전하게 저장
    const row = [
      new Date(),
      safeText_(p.source),
      safeText_(p.study_purpose),
      safeText_(p.study_methods),
      safeText_(p.study_management),
      safeText_(p.study_device),
      safeText_(p.study_pain_points),
      safeText_(p.interested_features),
      safeText_(p.usage_intent),
      safeText_(p.feedback || "")
    ];

    sheet.appendRow(row);

    return createResponse_("ok");

  } catch (error) {
    console.error("Survey submission error:", error);
    return createResponse_("error");
  }
}

/**
 * 설문 입력값 검증
 */
function validateSurvey_(p) {
  const requiredFields = [
    "source",
    "study_purpose",
    "study_methods",
    "study_management",
    "study_device",
    "study_pain_points",
    "interested_features",
    "usage_intent"
  ];

  for (const field of requiredFields) {
    if (
      typeof p[field] !== "string" ||
      p[field].trim() === ""
    ) {
      return {
        valid: false,
        reason: "Missing required field: " + field
      };
    }
  }

  const multipleFields = new Set([
    "study_methods",
    "study_management",
    "study_pain_points",
    "interested_features"
  ]);

  for (const field of requiredFields) {
    const values = multipleFields.has(field)
      ? p[field].split(" | ").map(v => v.trim())
      : [p[field].trim()];

    if (
      values.some(v => !SURVEY_OPTIONS[field].includes(v)) ||
      new Set(values).size !== values.length
    ) {
      return {
        valid: false,
        reason: "Invalid value: " + field
      };
    }

    if (
      !multipleFields.has(field) &&
      values.length !== 1
    ) {
      return {
        valid: false,
        reason: "Invalid selection count: " + field
      };
    }

    if (
      field === "interested_features" &&
      values.length > 2
    ) {
      return {
        valid: false,
        reason: "Q7 selection limit exceeded"
      };
    }
  }

  // Q7 배타적 선택 검증
  const features = p.interested_features.split(" | ");

  if (
    features.includes("특별히 도움이 될 것 같은 기능이 없다") &&
    features.length > 1
  ) {
    return {
      valid: false,
      reason: "Q7 exclusive option conflict"
    };
  }

  // Q4, Q6 배타적 선택 검증
  const exclusiveRules = {
    study_management: [
      "영상을 시청하고 별도로 기록하지 않는다.",
      "유튜브로 영어 공부를 하지 않는다."
    ],
    study_pain_points: [
      "특별히 불편한 점은 없다.",
      "유튜브로 영어 공부를 해본 적이 없다."
    ]
  };

  for (const field of Object.keys(exclusiveRules)) {
    const values = p[field].split(" | ");

    if (
      values.length > 1 &&
      values.some(v => exclusiveRules[field].includes(v))
    ) {
      return {
        valid: false,
        reason: "Exclusive option conflict: " + field
      };
    }
  }

  if (
    typeof (p.feedback || "") !== "string" ||
    (p.feedback || "").length > SURVEY_CONFIG.maxFeedbackLength
  ) {
    return {
      valid: false,
      reason: "Invalid feedback"
    };
  }

  return { valid: true };
}

/**
 * 스프레드시트 수식 주입 방어
 *
 * 수식으로 해석될 수 있는 문자열은
 * 작은따옴표를 앞에 추가하여 텍스트로 저장한다.
 */
function safeText_(value) {
  const text = String(value == null ? "" : value);

  if (/^\s*[=+\-@]/.test(text)) {
    return "'" + text;
  }

  return text;
}

/**
 * 요청 빈도 제한
 *
 * 전체 요청을 기준으로 60초 동안 최대 60회 허용한다.
 * 동시 요청 시 카운터 충돌을 방지하기 위해 LockService를 사용한다.
 */
function checkRateLimit_() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    return false;
  }

  try {
    const cache = CacheService.getScriptCache();
    const key = "survey_rate_limit_v1";
    const now = Date.now();

    const stored = cache.get(key);

    let state = stored
      ? JSON.parse(stored)
      : { startedAt: now, count: 0 };

    if (
      now - state.startedAt >= SURVEY_CONFIG.rateWindowMs
    ) {
      state = {
        startedAt: now,
        count: 0
      };
    }

    if (state.count >= SURVEY_CONFIG.rateLimit) {
      return false;
    }

    state.count++;

    cache.put(
      key,
      JSON.stringify(state),
      60
    );

    return true;

  } finally {
    lock.releaseLock();
  }
}

/**
 * API 응답 생성
 */
function createResponse_(status) {
  return ContentService
    .createTextOutput(status)
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * 웹 앱 상태 확인
 */
function doGet() {
  return ContentService
    .createTextOutput(
      JSON.stringify({
        status: "ok",
        message: "English Sentence Tracker survey API is running."
      })
    )
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 수동 테스트
 *
 * 실행 시 responses 시트에 테스트 응답 1행이 추가된다.
 */
function testDoPost() {
  const mockEvent = {
    parameter: {
      source: "네이버 블로그",
      study_purpose: "일상 회화",
      study_methods: "유튜브 등 무료 동영상 플랫폼 | AI 대화 도구(ChatGPT 등)",
      study_management: "나중에 다시 보려고 영상을 저장하거나 채널을 구독한다.",
      study_device: "PC·노트북",
      study_pain_points: "이전에 공부한 문장을 다시 찾아 복습하기 어렵다.",
      interested_features: "공부한 영어 문장을 한곳에 모아 저장하는 기능",
      usage_intent: "4",
      feedback: "설문 저장 테스트"
    }
  };

  const result = doPost(mockEvent);
  console.log(result.getContent());
}
