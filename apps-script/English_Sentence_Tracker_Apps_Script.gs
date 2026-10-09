function trackerConfig\_() {



  return {



    sheets: {



      raw: "자막입력",



      bank: "문장은행",



      repeat: "말하기연습",



      writing: "작문연습",



      today: "학습기록",



      dashboard: "학습대시보드",



      reset: "데이터관리",



      settings: "설정"



    }



  };



}



function trackerStyleConfig\_() {



  return {



    blank: "#FFFFFF",



    base: "#F8FAFC",



    input: "#FFF2CC",



    blue: "#EFF6FF",



    rawText: "#C1E4F5",



    gray: "#E7E6E6",



    delete: "#F0FDF4",



    done: "#DCFCE7"



  };



}



/\* =========================================================



   메뉴



\========================================================= \*/



function onOpen() {



  // 스프레드시트를 열 때 데이터가 없는 행의 배경색을 초기화한다.



  // 데이터가 있는 행에는 시트별 서식을 다시 적용한다.



  try {



    syncDataRowVisuals\_();



    syncCheckboxesToDataRows\_();



    ensureLearningConditionalFormatting\_();



  } catch (err) {



    console.log("초기 화면 정리 중 오류: " + err);



  }



  SpreadsheetApp.getUi()



    .createMenu("영어문장 자동정리")



    .addItem("배포용 시트 초기설정", "setupDistributionTemplate")



    .addSeparator()



    .addItem("1. 자막 자동 정리", "processSubtitleText")



    .addSeparator()



    .addItem("마지막 자막 입력 되돌리기", "rollbackLastSubtitleInput")



    .addItem("선택 문장 완전 삭제", "deleteSelectedBankSentences")



    .addSeparator()



    .addItem("학습상태 다시 계산", "recalculateAllLearningStates")



    .addItem("학습대시보드 새로고침", "refreshLearningDashboard")



    .addSeparator()



    .addItem("전체 학습데이터 초기화", "resetAllLearningData")



    .addSeparator()



    .addItem("시트 구조 점검", "validateTrackerStructure")



    .addToUi();



}



/\* =========================================================



   초기 설정



\========================================================= \*/



function setupDistributionTemplate() {



  const ss = SpreadsheetApp.getActiveSpreadsheet();



  const cfg = trackerConfig\_().sheets;



  const rawSheet = ss.getSheetByName(cfg.raw);



  const bankSheet = ss.getSheetByName(cfg.bank);



  const repeatSheet = ss.getSheetByName(cfg.repeat);



  const writingSheet = ss.getSheetByName(cfg.writing);



  const todaySheet = ss.getSheetByName(cfg.today);



  const settingsSheet = ss.getSheetByName(cfg.settings);



  if (!rawSheet || !bankSheet || !repeatSheet || !writingSheet || !todaySheet || !settingsSheet) {



    SpreadsheetApp.getUi().alert(



      "필수 시트를 찾을 수 없습니다.\n\n" +



      "자막입력, 문장은행, 말하기연습, 작문연습, 학습기록, 설정을 확인하세요."



    );



    return;



  }



  applyCheckboxesPreservingValues\_(repeatSheet, 6, 6);



  applyCheckboxesPreservingValues\_(writingSheet, 7, 1);



  applyCheckboxesPreservingValues\_(writingSheet, 9, 1);



  applyCheckboxesPreservingValues\_(todaySheet, 6, 4);



  const statusRule = SpreadsheetApp.newDataValidation()



    .requireValueInList(["대기", "정리중", "완료", "문장 없음"], true)



    .setAllowInvalid(false)



    .build();



  rawSheet.getRange("F2:F").setDataValidation(statusRule);



  applyBankCategoryValidation\_(ss);



  applyBankDeleteCheckboxes\_(bankSheet);



  repeatSheet.getRange("L2:L").setNumberFormat("0%");



  bankSheet.getRange("G2:G").setNumberFormat("yyyy-mm-dd hh:mm");



  bankSheet.getRange("H2:H").setNumberFormat("yyyy-mm-dd");



  todaySheet.getRange("A2:A").setNumberFormat("yyyy-mm-dd");



  syncDataRowVisuals\_();



  ensureLearningConditionalFormatting\_();



  SpreadsheetApp.flush();



  SpreadsheetApp.getUi().alert(



    "초기설정 완료\n\n" +



    "자막 자동 정리는 문장부호 문장을 우선 추출하고, " +



    "한국어 사이 영어와 반복 문장을 보조적으로 복구합니다."



  );



}



function applyCheckboxesPreservingValues\_(sheet, startCol, numCols) {



  if (!sheet) return;



  const maxRows = Math.max(sheet.getMaxRows(), 2);



  const numRows = maxRows - 1;



  const range = sheet.getRange(2, startCol, numRows, numCols);



  const values = range.getValues();



  const keyCol = getCheckboxKeyColumn\_(sheet);



  // 기존 체크박스 데이터 유효성 검사 규칙을 제거한 후 데이터 행에만 다시 적용한다.



  range.clearDataValidations();



  const keys = sheet.getRange(2, keyCol, numRows, 1).getValues();



  for (let i = 0; i < numRows; i++) {



    const rowNum = i + 2;



    const hasData = String(keys[i][0] || "").trim() !== "";



    const rowRange = sheet.getRange(rowNum, startCol, 1, numCols);



    if (hasData) {



      const rowValues = values[i].map(value => value === true);



      rowRange.insertCheckboxes();



      rowRange.setValues([rowValues]);



    } else {



      // 키 값이 없는 행의 셀 내용을 비운다.



      rowRange.clearContent();



    }



  }



}



function getCheckboxKeyColumn\_(sheet) {



  const name = sheet.getName();



  const cfg = trackerConfig\_().sheets;



  if (name === cfg.today) return 2; // 학습기록: B열 문장ID



  return 1;                         // 말하기연습/작문연습: A열 고유ID



}



function syncCheckboxesToDataRows\_() {



  const ss = SpreadsheetApp.getActiveSpreadsheet();



  const cfg = trackerConfig\_().sheets;



  const bankSheet = ss.getSheetByName(cfg.bank);



  const repeatSheet = ss.getSheetByName(cfg.repeat);



  const writingSheet = ss.getSheetByName(cfg.writing);



  const todaySheet = ss.getSheetByName(cfg.today);



  applyCheckboxesPreservingValues\_(repeatSheet, 6, 6);



  applyCheckboxesPreservingValues\_(writingSheet, 7, 1);



  applyCheckboxesPreservingValues\_(writingSheet, 9, 1);



  applyCheckboxesPreservingValues\_(todaySheet, 6, 4);



  applyBankDeleteCheckboxes\_(bankSheet);



}



function applyBankCategoryValidation\_(ss) {



  const bankSheet = ss.getSheetByName(trackerConfig\_().sheets.bank);



  const settingsSheet = ss.getSheetByName(trackerConfig\_().sheets.settings);



  if (!bankSheet || !settingsSheet) return;



  const lastSettingsRow = Math.max(settingsSheet.getLastRow(), 2);



  const values = settingsSheet



    .getRange(2, 1, lastSettingsRow - 1, 2)



    .getValues();



  const categories = values



    .filter(row => String(row[0] || "").trim() === "카테고리")



    .map(row => String(row[1] || "").trim())



    .filter(Boolean);



  if (categories.length === 0) return;



  const rule = SpreadsheetApp.newDataValidation()



    .requireValueInList(categories, true)



    .setAllowInvalid(false)



    .build();



  const maxRows = Math.max(bankSheet.getMaxRows(), 2);



  bankSheet.getRange(2, 6, maxRows - 1, 1).setDataValidation(rule);



}



function applyBankDeleteCheckboxes\_(bankSheet) {



  if (!bankSheet) return;



  if (bankSheet.getRange("J1").getValue() !== "삭제") {



    bankSheet.getRange("J1").setValue("삭제");



  }



  const maxRows = Math.max(bankSheet.getMaxRows(), 2);



  const numRows = maxRows - 1;



  const range = bankSheet.getRange(2, 10, numRows, 1);



  const values = range.getValues();



  const keys = bankSheet.getRange(2, 1, numRows, 1).getValues();



  // 삭제 체크박스를 데이터가 있는 행에만 적용한다.



  range.clearDataValidations();



  for (let i = 0; i < numRows; i++) {



    const rowNum = i + 2;



    const hasData = String(keys[i][0] || "").trim() !== "";



    const cell = bankSheet.getRange(rowNum, 10);



    if (hasData) {



      cell.insertCheckboxes();



      cell.setValue(values[i][0] === true);



    } else {



      cell.clearContent();



    }



  }



}



/\* =========================================================



   편집 시 학습상태 자동 동기화



\========================================================= \*/



function onEdit(e) {



  if (!e || !e.range) return;



  const sheet = e.range.getSheet();



  const sheetName = sheet.getName();



  const row = e.range.getRow();



  const col = e.range.getColumn();



  if (row < 2) return;



  // 자막입력 시트에서 편집된 데이터 행의 배경색을 갱신한다.



  // 여러 행이 동시에 편집된 경우 각 행을 처리한다.



  if (sheetName === trackerConfig\_().sheets.raw) {



    const startRow = e.range.getRow();



    const endRow = startRow + e.range.getNumRows() - 1;



    for (let r = startRow; r <= endRow; r++) {



      if (r < 2) continue;



      if (rowHasAnyData\_(sheet, r, 1, 7)) {



        styleRawRows\_(sheet, r, 1);



      } else {



        clearRowBackground\_(sheet, r, 7);



      }



    }



  }



  let sentenceId = "";



  let learningStarted = false;



  if (sheetName === trackerConfig\_().sheets.repeat && col >= 6 && col <= 11) {



    sentenceId = String(sheet.getRange(row, 2).getValue() || "").trim();



    learningStarted = e.value === "TRUE";



  }



  if (



    sheetName === trackerConfig\_().sheets.writing &&



    [5, 7, 8, 9].indexOf(col) !== -1



  ) {



    sentenceId = String(sheet.getRange(row, 2).getValue() || "").trim();



    if (col === 5 || col === 8) {



      learningStarted = String(e.range.getValue() || "").trim() !== "";



    } else {



      learningStarted = e.value === "TRUE";



    }



  }



  if (!sentenceId) return;



  const ss = e.source;



  if (learningStarted) {



    setFirstStudyDateIfNeeded\_(ss, sentenceId);



  }



  updateLearningStateBySentenceId\_(ss, sentenceId);



}



/\* =========================================================



   자막 입력 처리 및 학습 시트 저장



   문장 추출 및 중복 처리 기준



   1) 문장부호(. ! ?)로 구분되는 문장을 우선 추출



   2) 한국어가 포함된 자막에서 반복 영어 예문 탐지



   3) 순수 영어 자막에서 문장부호 없는 반복 발화 탐지



   4) 정규화된 문장 키를 기준으로 중복 저장 방지



\========================================================= \*/



function processSubtitleText() {



  const ss = SpreadsheetApp.getActiveSpreadsheet();



  const cfg = trackerConfig\_().sheets;



  const rawSheet = ss.getSheetByName(cfg.raw);



  const bankSheet = ss.getSheetByName(cfg.bank);



  const repeatSheet = ss.getSheetByName(cfg.repeat);



  const writingSheet = ss.getSheetByName(cfg.writing);



  const todaySheet = ss.getSheetByName(cfg.today);



  if (!rawSheet || !bankSheet || !repeatSheet || !writingSheet || !todaySheet) {



    SpreadsheetApp.getUi().alert(



      "자막입력, 문장은행, 말하기연습, 작문연습, 학습기록 시트를 확인하세요."



    );



    return;



  }



  const rawLastRow = rawSheet.getLastRow();



  if (rawLastRow < 2) {



    SpreadsheetApp.getUi().alert("자막입력에 처리할 데이터가 없습니다.");



    return;



  }



  const rawData = rawSheet.getRange(2, 1, rawLastRow - 1, 7).getValues();



  const existingBank = getExistingBank\_(bankSheet);



  // 동일한 영상 ID를 가진 여러 입력 행을 하나의 자막 텍스트로 결합한다.



  // 영상 ID가 없는 입력 행에는 새로운 영상 ID를 할당한다.



  const groups = [];



  const groupMap = {};



  rawData.forEach((row, i) => {



    const rowNum = i + 2;



    const rawText = String(row[4] || "").trim();



    const status = String(row[5] || "").trim();



    if (!rawText || status === "완료") return;



    let videoId = String(row[0] || "").trim();



    if (!videoId) {



      videoId = "V" + Utilities.getUuid().slice(0, 8);



      rawSheet.getRange(rowNum, 1).setValue(videoId);



    }



    if (!row[3]) {



      rawSheet.getRange(rowNum, 4).setValue(new Date());



    }



    if (!groupMap[videoId]) {



      groupMap[videoId] = {



        videoId: videoId,



        rowNums: [],



        texts: []



      };



      groups.push(groupMap[videoId]);



    }



    groupMap[videoId].rowNums.push(rowNum);



    groupMap[videoId].texts.push(rawText);



  });



  let processedGroups = 0;



  let extractedCount = 0;



  let savedCount = 0;



  let duplicateCount = 0;



  let emptyGroups = 0;



  groups.forEach(group => {



    group.rowNums.forEach(rowNum => {



      rawSheet.getRange(rowNum, 6).setValue("정리중");



    });



    const transcript = group.texts.join("\n");



    const candidates = extractSimpleReliableCandidates\_(transcript);



    if (candidates.length === 0) {



      group.rowNums.forEach(rowNum => {



        rawSheet.getRange(rowNum, 6).setValue("문장 없음");



      });



      emptyGroups++;



      return;



    }



    extractedCount += candidates.length;



    const toSave = [];



    candidates.forEach(sentence => {



      const cleaned = cleanupDisplaySentence\_(sentence);



      const key = normalizeSentenceKey\_(cleaned);



      if (!key) return;



      if (existingBank[key]) {



        duplicateCount++;



        return;



      }



      toSave.push({



        videoId: group.videoId,



        sentence: cleaned



      });



      // 현재 실행에서 이미 선택한 문장의 중복 저장을 방지한다.



      existingBank[key] = "\_\_PENDING\_\_";



    });



    if (toSave.length > 0) {



      const result = appendSentencesToLearningSheets\_(



        bankSheet,



        repeatSheet,



        writingSheet,



        todaySheet,



        toSave



      );



      savedCount += result.savedCount;



      result.savedItems.forEach(item => {



        existingBank[normalizeSentenceKey\_(item.sentence)] = item.sentenceId;



      });



    }



    group.rowNums.forEach(rowNum => {



      rawSheet.getRange(rowNum, 6).setValue("완료");



    });



    processedGroups++;



  });



  syncDataRowVisuals\_();



  syncCheckboxesToDataRows\_();



  ensureLearningConditionalFormatting\_();



  SpreadsheetApp.flush();



  SpreadsheetApp.getUi().alert(



    "자막 자동 정리 완료\n\n" +



    "처리한 영상: " + processedGroups + "\n" +



    "추출 및 정리된 문장 : " + extractedCount + "\n" +



    "새로 저장한 문장: " + savedCount + "\n" +



    "이미 문장은행에 있던 문장: " + duplicateCount + "\n" +



    "추출 문장 없음: " + emptyGroups



  );



}



/\* =========================================================



   학습용 영어 문장 후보 추출



   추출 및 필터링 기준



   1) 문장부호(. ! ?)로 끝나는 영어 문장을 우선 추출한다.



   2) 시간코드와 줄바꿈은 문장 구분 기준으로 사용하지 않는다.



   3) 화자 표시(>>)는 문장 경계로 처리하지 않고 제거한다.



   4) 문장부호 없이 반복되는 영어 표현을 추가로 탐지한다.



   5) 한국어가 포함된 자막에서는 반복 영어 예문을 탐지한다.



   6) 정규화된 키를 이용해 동일 문장을 제외하고 유사 문장 일괄 제거는 하지 않는다.



\========================================================= \*/



function extractSimpleReliableCandidates\_(text) {



  const prepared = prepareTranscriptSimple\_(text);



  if (!prepared) return [];



  const result = [];



  const seen = {};



  function add(sentence) {



    let cleaned = cleanupDisplaySentence\_(sentence);



    // 하나의 후보에 연속 반복된 문장이 있으면 축약 가능 여부를 검사한다.



    // 반복 구간이 후보 토큰의 70% 이상을 차지할 때만 축약한다.



    cleaned = collapseRepeatedCandidate\_(cleaned);



    const key = normalizeSentenceKey\_(cleaned);



    if (!key) return;



    if (!isUsableSentenceSimple\_(cleaned)) return;



    if (seen[key]) return;



    seen[key] = true;



    result.push(cleaned);



  }



  // 문장부호로 구분된 영어 문장 후보를 추출한다.



  extractPunctuatedSentencesSimple\_(prepared).forEach(add);



  const hasKorean = /[가-힣ㄱ-ㅎㅏ-ㅣ]/.test(prepared);



  // 한국어가 없는 자막에서 반복 발화를 탐지한다.



  // 한국어가 포함된 자막에는 서로 다른 예문의 결합 가능성을 고려해 이 탐지 방법을 적용하지 않는다.



  if (!hasKorean) {



    extractRepeatedSentencesSimple\_(prepared).forEach(add);



  }



  // 한국어가 포함된 자막에서 반복 영어 예문을 추출한다.



  if (hasKorean) {



    extractRepeatedMixedEnglishSimple\_(prepared).forEach(add);



  }



  return result;



}



/\* =========================================================



   전처리



   문장 추출 전에 자막 블록을 연결한다.



   시간코드, 줄바꿈, 화자 표시는 문장 구분 기준에서 제외한다.



\========================================================= \*/



function prepareTranscriptSimple\_(text) {



  let s = String(text || "")



    .replace(/\r\n?/g, "\n")



    .replace(/\u00A0/g, " ");



  // 시간코드(분:초 또는 시:분:초) 제거



  s = s.replace(



    /(^|\n)\s\*\d{1,2}:\d{2}(?::\d{2})?\s\*(?=\n|$)/g,



    "\n"



  );



  // 구분선 제거



  s = s.replace(/(^|\n)\s\*-{3,}\s\*(?=\n|$)/g, "\n");



  // 음악·박수·웃음 등의 자막 표기 제거



  s = s.replace(/\\[(?:music|applause|laughter|음악|박수|웃음)\\]/gi, " ");



  // 화자 표시(>>)를 공백으로 치환한다.



  s = s.replace(/>>+/g, " ");



  // 자막의 줄바꿈을 공백으로 변환하여 문장을 연결한다.



  s = s



    .replace(/\n+/g, " ")



    .replace(/\s+/g, " ")



    .replace(/\s+([,.!?;:])/g, "$1")



    .trim();



  return s;



}



/\* =========================================================



   문장부호 기반 영어 문장 추출



   문장부호로 구분된 구간을 기준으로 후보를 추출한다.



   한국어가 포함된 구간에서는 영어 부분을 별도로 추출한다.



\========================================================= \*/



function extractPunctuatedSentencesSimple\_(text) {



  const matches = String(text || "").match(/[^.!?]+[.!?]+/g) || [];



  const out = [];



  matches.forEach(chunk => {



    const cleaned = extractEnglishFromPunctuatedChunk\_(chunk);



    if (cleaned) out.push(cleaned);



  });



  return out;



}



function extractEnglishFromPunctuatedChunk\_(chunk) {



  let s = String(chunk || "").trim();



  if (!s || !/[A-Za-z]/.test(s)) return "";



  // 한국어가 없는 후보는 정리 후 반환한다.



  if (!/[가-힣ㄱ-ㅎㅏ-ㅣ]/.test(s)) {



    return cleanupDisplaySentence\_(s);



  }



  // 한국어를 구분자로 사용하여 영어 부분을 추출한다.



  const parts = s



    .replace(/[가-힣ㄱ-ㅎㅏ-ㅣ]+/g, "\n")



    .split(/\n+/)



    .map(x => cleanupDisplaySentence\_(x))



    .filter(x => /[A-Za-z]/.test(x));



  if (parts.length === 0) return "";



  // 문장부호로 끝나는 영어 조각을 우선 선택한다.



  const withPunctuation = parts.filter(x => /[.!?]$/.test(x));



  const pool = withPunctuation.length > 0 ? withPunctuation : parts;



  pool.sort((a, b) => wordCount\_(b) - wordCount\_(a));



  return pool[0] || "";



}



/\* =========================================================



   순수 영어 자막의 연속 반복 문장 탐지



   예:



   I'm on my way I'm on my way I'm on my way



   → I'm on my way



\========================================================= \*/



function extractRepeatedSentencesSimple\_(text) {



  const englishOnly = String(text || "")



    .replace(/[가-힣ㄱ-ㅎㅏ-ㅣ]+/g, " ")



    .replace(/\s+/g, " ")



    .trim();



  if (!englishOnly) return [];



  const repeats = findMaximalImmediateRepeats\_(englishOnly);



  const out = [];



  repeats.forEach(item => {



    const sentence = cleanupDisplaySentence\_(item.sentence);



    const wc = wordCount\_(sentence);



    // 단어 수가 3개 미만이거나 30개를 초과하는 후보는 제외한다.



    if (wc < 3 || wc > 30) return;



    if (!looksLikeSentenceStartSimple\_(sentence)) return;



    out.push(sentence);



  });



  return out;



}



/\* =========================================================



   한국어 포함 자막의 반복 영어 예문 추출



   한국어를 기준으로 나눈 영어 조각 중 동일한 조각이



   2회 이상 등장하면 문장 후보로 선택한다.



\========================================================= \*/



function extractRepeatedMixedEnglishSimple\_(text) {



  if (!/[가-힣ㄱ-ㅎㅏ-ㅣ]/.test(text)) return [];



  const parts = String(text || "")



    .replace(/[가-힣ㄱ-ㅎㅏ-ㅣ]+/g, "\n")



    .split(/\n+/)



    .map(x => cleanupDisplaySentence\_(x))



    .filter(x => /[A-Za-z]/.test(x));



  const countMap = {};



  const displayMap = {};



  const resultMap = {};



  parts.forEach(part => {



    const cleaned = stripTeachingCuesSimple\_(part);



    const key = normalizeSentenceKey\_(cleaned);



    const wc = wordCount\_(cleaned);



    if (key && wc >= 2 && wc <= 20) {



      countMap[key] = (countMap[key] || 0) + 1;



      displayMap[key] = cleaned;



    }



    // 하나의 영어 조각 안에서 발생하는 연속 반복도 검사한다.



    // 예: I'm used to traveling I'm used to traveling I'm used to traveling



    findMaximalImmediateRepeats\_(cleaned).forEach(item => {



      const sentence = cleanupDisplaySentence\_(item.sentence);



      const sentenceKey = normalizeSentenceKey\_(sentence);



      const sentenceWc = wordCount\_(sentence);



      if (!sentenceKey || sentenceWc < 2 || sentenceWc > 20) return;



      if (!looksLikeSentenceStartSimple\_(sentence)) return;



      resultMap[sentenceKey] = sentence;



    });



  });



  // 한국어 사이에서 동일한 영어 조각이 두 번 이상 나타나면 후보에 추가한다.



  Object.keys(countMap).forEach(key => {



    if (countMap[key] < 2) return;



    const sentence = displayMap[key];



    if (!looksLikeSentenceStartSimple\_(sentence)) return;



    resultMap[key] = sentence;



  });



  return Object.keys(resultMap).map(key => resultMap[key]);



}



function stripTeachingCuesSimple\_(text) {



  let s = cleanupDisplaySentence\_(text);



  // 일반 영어 단어는 제거 대상에서 제외한다.



  // 사전에 정의한 학습 지시 표현만 제거한다.



  const cues = [



    "one more time",



    "repeat after me"



  ];



  cues.forEach(cue => {



    const esc = escapeRegex\_(cue);



    s = s.replace(new RegExp("^(?:" + esc + ")[,:-]?\\\s+", "i"), "");



    s = s.replace(new RegExp("\\\s+(?:" + esc + ")[.!,:-]?$", "i"), "");



  });



  return cleanupDisplaySentence\_(s);



}



function looksLikeSentenceStartSimple\_(sentence) {



  const words = normalizeSentenceKey\_(sentence).split(" ").filter(Boolean);



  if (words.length === 0) return false;



  const first = words[0];



  const badStarters = new Set([



    "of", "to", "in", "on", "at", "for", "with", "from", "about",



    "and", "but", "or", "because", "than", "as"



  ]);



  return !badStarters.has(first);



}



function isUsableSentenceSimple\_(sentence) {



  const s = cleanupDisplaySentence\_(sentence);



  const wc = wordCount\_(s);



  if (!s) return false;



  if (!/[A-Za-z]/.test(s)) return false;



  if (/[가-힣]/.test(s)) return false;



  if (/https?:\\/\\//i.test(s) || /www\\./i.test(s)) return false;



  if (wc < 2 || wc > 40) return false;



  return true;



}



/\* =========================================================



   즉시 반복 탐지



\========================================================= \*/



function findMaximalImmediateRepeats\_(text) {



  const tokens = tokenizeEnglish\_(text);



  const out = [];



  if (tokens.length < 6) return out;



  let i = 0;



  while (i < tokens.length) {



    const remain = tokens.length - i;



    let best = null;



    // 반복 비교 단위를 최대 30개 토큰으로 제한한다.



    const maxSize = Math.min(30, Math.floor(remain / 2));



    for (let size = 3; size <= maxSize; size++) {



      if (!tokenSlicesEqual\_(tokens, i, i + size, size)) continue;



      let count = 2;



      while (



        i + (count + 1) \* size <= tokens.length &&



        tokenSlicesEqual\_(tokens, i, i + count \* size, size)



      ) {



        count++;



      }



      const coverage = size \* count;



      if (



        !best ||



        coverage > best.coverage ||



        (coverage === best.coverage && size > best.size)



      ) {



        best = {



          size: size,



          count: count,



          coverage: coverage



        };



      }



    }



    if (!best) {



      i++;



      continue;



    }



    const sentence = tokens.slice(i, i + best.size).join(" ");



    out.push({



      sentence: cleanupDisplaySentence\_(sentence),



      count: best.count



    });



    i += best.coverage;



  }



  return out;



}



function collapseRepeatedCandidate\_(sentence) {



  const cleaned = cleanupDisplaySentence\_(sentence);



  if (!cleaned) return "";



  const tokens = tokenizeEnglish\_(cleaned);



  // 토큰 수가 6개 미만인 후보는 반복 축약을 수행하지 않는다.



  if (tokens.length < 6) return cleaned;



  const repeats = findMaximalImmediateRepeats\_(cleaned);



  if (repeats.length === 0) {



    return cleaned;



  }



  // 후보에서 가장 많은 토큰을 포함하는 연속 반복 패턴을 선택한다.



  let best = null;



  repeats.forEach(item => {



    const unitTokens = tokenizeEnglish\_(item.sentence);



    const repeatedTokenCount = unitTokens.length \* item.count;



    if (



      !best ||



      repeatedTokenCount > best.repeatedTokenCount



    ) {



      best = {



        sentence: item.sentence,



        count: item.count,



        repeatedTokenCount: repeatedTokenCount



      };



    }



  });



  if (!best || best.count < 2) {



    return cleaned;



  }



  // 부분적인 반복만 발견된 경우 원본 후보를 유지한다.



  // 반복 구간의 토큰 비율이 70% 이상일 때만 축약한다.



  const coverage = best.repeatedTokenCount / tokens.length;



  if (coverage < 0.7) {



    return cleaned;



  }



  return cleanupDisplaySentence\_(best.sentence);



}



/\* =========================================================



   영어 토큰화



\========================================================= \*/



function tokenizeEnglish\_(text) {



  return String(text || "").match(



    /[A-Za-z]+(?:['’][A-Za-z]+)?|\d+(?::\d{2})?/g



  ) || [];



}



function tokenSlicesEqual\_(tokens, aStart, bStart, size) {



  for (let i = 0; i < size; i++) {



    if (



      normalizeToken\_(tokens[aStart + i]) !==



      normalizeToken\_(tokens[bStart + i])



    ) {



      return false;



    }



  }



  return true;



}



function normalizeToken\_(token) {



  return String(token || "")



    .toLowerCase()



    .replace(/[‘’]/g, "'")



    .replace(/[“”]/g, '"');



}



/\* =========================================================



   문장 정리 / 정규화



\========================================================= \*/



function cleanupDisplaySentence\_(sentence) {



  return String(sentence || "")



    .replace(/[“”]/g, '"')



    .replace(/[‘’]/g, "'")



    .replace(/\s+/g, " ")



    .replace(/\s+([,.!?;:])/g, "$1")



    .replace(/^[,;:\\-–—\s]+/, "")



    .replace(/[,;:\\-–—\s]+$/, "")



    .trim();



}



function normalizeSentenceKey\_(sentence) {



  return cleanupDisplaySentence\_(sentence)



    .toLowerCase()



    .replace(/[.!?]+$/g, "")



    .replace(/[^a-z0-9' ]+/g, " ")



    .replace(/\s+/g, " ")



    .trim();



}



function wordCount\_(text) {



  return tokenizeEnglish\_(text).length;



}



function escapeRegex\_(text) {



  return String(text || "").replace(/[.\*+?^${}()|[\\]\\\\]/g, "\\\\$&");



}



/\* =========================================================



   문장은행 + 학습 시트 동시 추가



\========================================================= \*/



function appendSentencesToLearningSheets\_(bankSheet, repeatSheet, writingSheet, todaySheet, items) {



  if (!items || items.length === 0) {



    return { savedCount: 0, savedItems: [] };



  }



  const bankRows = [];



  const repeatRows = [];



  const writingRows = [];



  const todayRows = [];



  const savedItems = [];



  const bankStart = getNextRealRow\_(bankSheet, 1);



  const repeatStart = getNextRealRow\_(repeatSheet, 1);



  const writingStart = getNextRealRow\_(writingSheet, 1);



  const todayStart = getNextRealRow\_(todaySheet, 2);



  items.forEach(item => {



    const sentenceId = "S" + Utilities.getUuid().slice(0, 8);



    const addedAt = new Date();



    const bankRow = bankStart + bankRows.length;



    const repeatRow = repeatStart + repeatRows.length;



    const writingRow = writingStart + writingRows.length;



    const todayRow = todayStart + todayRows.length;



    // 문장은행: A:J



    bankRows.push([



      sentenceId,                                                     // A 문장ID



      item.videoId,                                                   // B 영상ID



      item.sentence,                                                  // C 영어문장



      '=GOOGLETRANSLATE(C' + bankRow + ',"en","ko")',               // D 한국어 뜻



      "",                                                            // E 중요문형



      "",                                                            // F 카테고리



      addedAt,                                                        // G 추가일시



      "",                                                            // H 최초학습일



      "",                                                            // I 비고



      false                                                          // J 삭제



    ]);



    // 말하기연습: A:M



    repeatRows.push([



      "P" + Utilities.getUuid().slice(0, 8),



      sentenceId,



      item.videoId,



      '=XLOOKUP(B' + repeatRow + ',문장은행!A:A,문장은행!C:C,"")',



      '=XLOOKUP(B' + repeatRow + ',문장은행!A:A,문장은행!D:D,"")',



      false,



      false,



      false,



      false,



      false,



      false,



      '=COUNTIF(F' + repeatRow + ':K' + repeatRow + ',TRUE)/6',



      ""



    ]);



    // 작문연습: A:J



    writingRows.push([



      "W" + Utilities.getUuid().slice(0, 8),



      sentenceId,



      '=XLOOKUP(B' + writingRow + ',문장은행!A:A,문장은행!F:F,"")',



      '=XLOOKUP(B' + writingRow + ',문장은행!A:A,문장은행!D:D,"")',



      "",



      '=IF(G' + writingRow + '=TRUE,XLOOKUP(B' + writingRow + ',문장은행!A:A,문장은행!C:C,""),"")',



      false,



      "",



      false,



      ""



    ]);



    // 학습기록: A:J



    todayRows.push([



      addedAt,



      sentenceId,



      '=XLOOKUP(B' + todayRow + ',문장은행!A:A,문장은행!C:C,"")',



      '=XLOOKUP(B' + todayRow + ',문장은행!A:A,문장은행!D:D,"")',



      "신규 문장",



      false,



      false,



      false,



      false,



      ""



    ]);



    savedItems.push({



      sentenceId: sentenceId,



      sentence: item.sentence



    });



  });



  bankSheet.getRange(bankStart, 1, bankRows.length, 10).setValues(bankRows);



  bankSheet.getRange(bankStart, 7, bankRows.length, 1).setNumberFormat("yyyy-mm-dd hh:mm");



  bankSheet.getRange(bankStart, 8, bankRows.length, 1).setNumberFormat("yyyy-mm-dd");



  bankSheet.getRange(bankStart, 10, bankRows.length, 1).insertCheckboxes();



  repeatSheet.getRange(repeatStart, 1, repeatRows.length, 13).setValues(repeatRows);



  repeatSheet.getRange(repeatStart, 6, repeatRows.length, 6).insertCheckboxes();



  repeatSheet.getRange(repeatStart, 12, repeatRows.length, 1).setNumberFormat("0%");



  writingSheet.getRange(writingStart, 1, writingRows.length, 10).setValues(writingRows);



  writingSheet.getRange(writingStart, 7, writingRows.length, 1).insertCheckboxes();



  writingSheet.getRange(writingStart, 9, writingRows.length, 1).insertCheckboxes();



  todaySheet.getRange(todayStart, 1, todayRows.length, 10).setValues(todayRows);



  todaySheet.getRange(todayStart, 6, todayRows.length, 4).insertCheckboxes();



  todaySheet.getRange(todayStart, 1, todayRows.length, 1).setNumberFormat("yyyy-mm-dd");



  // 새로 추가한 데이터 행에 시트별 배경색을 적용한다.



  styleBankRows\_(bankSheet, bankStart, bankRows.length);



  styleRepeatRows\_(repeatSheet, repeatStart, repeatRows.length);



  styleWritingRows\_(writingSheet, writingStart, writingRows.length);



  styleTodayRows\_(todaySheet, todayStart, todayRows.length);



  ensureLearningConditionalFormatting\_();



  applyBankCategoryValidation\_(SpreadsheetApp.getActiveSpreadsheet());



  applyBankDeleteCheckboxes\_(bankSheet);



  return {



    savedCount: bankRows.length,



    savedItems: savedItems



  };



}



/\* =========================================================



   마지막 자막 입력 되돌리기



   기준:



   - 자막입력에서 마지막으로 데이터가 입력된 1개 행



   - 그 행의 영상ID로 생성된 문장은행/말하기연습/작문연습/오늘학습 데이터까지 함께 삭제



\========================================================= \*/



function rollbackLastSubtitleInput() {



  const ss = SpreadsheetApp.getActiveSpreadsheet();



  const ui = SpreadsheetApp.getUi();



  const cfg = trackerConfig\_().sheets;



  const rawSheet = ss.getSheetByName(cfg.raw);



  const bankSheet = ss.getSheetByName(cfg.bank);



  const repeatSheet = ss.getSheetByName(cfg.repeat);



  const writingSheet = ss.getSheetByName(cfg.writing);



  const todaySheet = ss.getSheetByName(cfg.today);



  if (!rawSheet || !bankSheet || !repeatSheet || !writingSheet || !todaySheet) {



    ui.alert("필수 시트를 찾을 수 없습니다.");



    return;



  }



  const lastRawRow = findLastDataRowInRange\_(rawSheet, 1, 7);



  if (lastRawRow < 2) {



    ui.alert("자막입력에 되돌릴 입력이 없습니다.");



    return;



  }



  const rawValues = rawSheet.getRange(lastRawRow, 1, 1, 7).getValues()[0];



  const videoId = String(rawValues[0] || "").trim();



  const title = String(rawValues[1] || "").trim();



  const rawPreview = String(rawValues[4] || "").trim().replace(/\s+/g, " ").slice(0, 80);



  // 같은 영상ID가 여러 자막입력 행에 있으면 어느 학습 데이터가 마지막 행에서 생성됐는지



  // 안전하게 구분할 수 없으므로 자동 삭제하지 않는다.



  if (videoId) {



    const rawLast = rawSheet.getLastRow();



    if (rawLast >= 2) {



      const ids = rawSheet.getRange(2, 1, rawLast - 1, 1).getValues().flat();



      const sameIdCount = ids.filter(v => String(v || "").trim() === videoId).length;



      if (sameIdCount > 1) {



        ui.alert(



          "되돌리기 중단",



          "마지막 행의 영상ID(" + videoId + ")가 자막입력의 여러 행에 사용되고 있습니다.\n\n" +



          "이 경우 관련 학습 데이터가 어느 행에서 생성됐는지 안전하게 구분할 수 없어 자동 삭제하지 않습니다.",



          ui.ButtonSet.OK



        );



        return;



      }



    }



  }



  const sentenceIds = videoId



    ? collectSentenceIdsByVideoId\_(bankSheet, videoId)



    : new Set();



  const label = title || rawPreview || (videoId ? videoId : "마지막 자막 입력");



  const answer = ui.alert(



    "마지막 자막 입력 되돌리기",



    "자막입력의 마지막 입력 1개를 되돌립니다.\n\n" +



    "대상: " + label + "\n" +



    (videoId ? "영상ID: " + videoId + "\n" : "") +



    "연결된 문장: " + sentenceIds.size + "개\n\n" +



    "자막입력 마지막 행과 그 입력에서 생성된 문장은행·말하기연습·작문연습·학습기록 데이터가 함께 삭제됩니다.\n" +



    "계속하시겠습니까?",



    ui.ButtonSet.YES_NO



  );



  if (answer !== ui.Button.YES) return;



  if (sentenceIds.size > 0) {



    deleteRowsBySentenceIds\_(repeatSheet, 2, sentenceIds);



    deleteRowsBySentenceIds\_(writingSheet, 2, sentenceIds);



    deleteRowsBySentenceIds\_(todaySheet, 2, sentenceIds);



    deleteRowsBySentenceIds\_(bankSheet, 1, sentenceIds);



  } else if (videoId) {



    deleteRowsByExactValue\_(repeatSheet, 3, videoId);



    deleteRowsByExactValue\_(bankSheet, 2, videoId);



  }



  rawSheet.deleteRow(lastRawRow);



  syncDataRowVisuals\_();



  SpreadsheetApp.flush();



  ui.alert(



    "되돌리기 완료\n\n" +



    "마지막 자막 입력 1개와 연결된 학습 데이터 " + sentenceIds.size + "개를 삭제했습니다."



  );



}



function collectSentenceIdsByVideoId\_(bankSheet, videoId) {



  const result = new Set();



  const lastRow = bankSheet.getLastRow();



  if (lastRow < 2 || !videoId) return result;



  const values = bankSheet.getRange(2, 1, lastRow - 1, 2).getValues();



  values.forEach(row => {



    const sentenceId = String(row[0] || "").trim();



    const rowVideoId = String(row[1] || "").trim();



    if (sentenceId && rowVideoId === videoId) {



      result.add(sentenceId);



    }



  });



  return result;



}



function deleteRowsByExactValue\_(sheet, col, value) {



  const lastRow = sheet.getLastRow();



  if (lastRow < 2) return;



  const target = String(value || "").trim();



  if (!target) return;



  const values = sheet.getRange(2, col, lastRow - 1, 1).getValues();



  const rows = [];



  values.forEach((row, i) => {



    if (String(row[0] || "").trim() === target) rows.push(i + 2);



  });



  rows.sort((a, b) => b - a).forEach(rowNum => sheet.deleteRow(rowNum));



}



/\* =========================================================



   선택 문장 완전 삭제



\========================================================= \*/



function deleteSelectedBankSentences() {



  const ss = SpreadsheetApp.getActiveSpreadsheet();



  const ui = SpreadsheetApp.getUi();



  const cfg = trackerConfig\_().sheets;



  const bankSheet = ss.getSheetByName(cfg.bank);



  const repeatSheet = ss.getSheetByName(cfg.repeat);



  const writingSheet = ss.getSheetByName(cfg.writing);



  const todaySheet = ss.getSheetByName(cfg.today);



  if (!bankSheet || !repeatSheet || !writingSheet || !todaySheet) {



    ui.alert("필수 시트를 찾을 수 없습니다.");



    return;



  }



  const lastRow = bankSheet.getLastRow();



  if (lastRow < 2) {



    ui.alert("문장은행에 삭제할 문장이 없습니다.");



    return;



  }



  const rows = bankSheet.getRange(2, 1, lastRow - 1, 10).getValues();



  const selected = [];



  rows.forEach((row, i) => {



    const sentenceId = String(row[0] || "").trim();



    const sentence = String(row[2] || "").trim();



    const checked = row[9] === true;



    if (sentenceId && checked) {



      selected.push({



        sentenceId: sentenceId,



        sentence: sentence,



        bankRow: i + 2



      });



    }



  });



  if (selected.length === 0) {



    ui.alert("문장은행 J열에서 삭제할 문장을 먼저 체크하세요.");



    return;



  }



  const preview = selected



    .slice(0, 5)



    .map(item => "• " + item.sentence)



    .join("\n");



  const more = selected.length > 5



    ? "\n외 " + (selected.length - 5) + "개"



    : "";



  const answer = ui.alert(



    "선택 문장 완전 삭제",



    "선택한 " + selected.length + "개 문장을 삭제합니다.\n\n" +



    preview + more + "\n\n" +



    "연결된 말하기연습·작문연습·학습기록 행도 함께 삭제됩니다.\n" +



    "계속하시겠습니까?",



    ui.ButtonSet.YES_NO



  );



  if (answer !== ui.Button.YES) return;



  const ids = new Set(selected.map(item => item.sentenceId));



  deleteRowsBySentenceIds\_(repeatSheet, 2, ids);



  deleteRowsBySentenceIds\_(writingSheet, 2, ids);



  deleteRowsBySentenceIds\_(todaySheet, 2, ids);



  selected



    .map(item => item.bankRow)



    .sort((a, b) => b - a)



    .forEach(rowNum => bankSheet.deleteRow(rowNum));



  syncDataRowVisuals\_();



  ensureLearningConditionalFormatting\_();



  SpreadsheetApp.flush();



  ui.alert("삭제 완료\n\n" + selected.length + "개 문장과 연결 학습 데이터를 함께 삭제했습니다.");



}



function deleteRowsBySentenceIds\_(sheet, sentenceIdCol, ids) {



  const lastRow = sheet.getLastRow();



  if (lastRow < 2) return;



  const values = sheet



    .getRange(2, sentenceIdCol, lastRow - 1, 1)



    .getValues();



  const rowsToDelete = [];



  values.forEach((row, i) => {



    const sentenceId = String(row[0] || "").trim();



    if (ids.has(sentenceId)) rowsToDelete.push(i + 2);



  });



  rowsToDelete



    .sort((a, b) => b - a)



    .forEach(rowNum => sheet.deleteRow(rowNum));



}



/\* =========================================================



   학습상태 계산



\========================================================= \*/



function updateLearningStateBySentenceId\_(ss, sentenceId) {



  const repeatSheet = ss.getSheetByName(trackerConfig\_().sheets.repeat);



  const writingSheet = ss.getSheetByName(trackerConfig\_().sheets.writing);



  const todaySheet = ss.getSheetByName(trackerConfig\_().sheets.today);



  if (!repeatSheet || !writingSheet || !todaySheet) return;



  let speakingStarted = false;



  let speakingCompleted = false;



  let recording = false;



  let writingDone = false;



  const repeatRow = findRowByValue\_(repeatSheet, 2, sentenceId);



  if (repeatRow) {



    const repeatValues = repeatSheet



      .getRange(repeatRow, 6, 1, 6)



      .getValues()[0];



    // 1\~5회 중 하나라도 체크하면 말하기 학습 시작



    speakingStarted = repeatValues



      .slice(0, 5)



      .some(v => v === true);



    // 1\~5회가 모두 체크되어야 말하기 완료



    speakingCompleted = repeatValues



      .slice(0, 5)



      .every(v => v === true);



    // K열 최종녹음



    recording = repeatValues[5] === true;



  }



  const writingRow = findRowByValue\_(writingSheet, 2, sentenceId);



  if (writingRow) {



    writingDone =



      writingSheet.getRange(writingRow, 9).getValue() === true;



  }



  // 전체 학습완료:



  // 말하기 5회 완료 + 작문 완료 + 최종녹음 완료



  const completed =



    speakingCompleted &&



    writingDone &&



    recording;



  let status = "신규 문장";



  if (completed) {



    status = "학습완료";



  } else if (



    speakingStarted ||



    writingDone ||



    recording



  ) {



    status = "학습 중";



  }



  const todayLastRow = todaySheet.getLastRow();



  if (todayLastRow < 2) return;



  const ids = todaySheet



    .getRange(2, 2, todayLastRow - 1, 1)



    .getValues();



  ids.forEach((row, index) => {



    if (



      String(row[0] || "").trim() !== sentenceId



    ) return;



    const targetRow = index + 2;



    todaySheet



      .getRange(targetRow, 5, 1, 5)



      .setValues([[



        status,



        speakingCompleted,



        writingDone,



        recording,



        completed



      ]]);



    todaySheet



      .getRange(targetRow, 6, 1, 4)



      .insertCheckboxes();



  });



}



function setFirstStudyDateIfNeeded\_(ss, sentenceId) {



  const bankSheet = ss.getSheetByName(trackerConfig\_().sheets.bank);



  if (!bankSheet) return;



  const row = findRowByValue\_(bankSheet, 1, sentenceId);



  if (!row) return;



  const dateCell = bankSheet.getRange(row, 8);



  if (!dateCell.getValue()) {



    dateCell.setValue(new Date());



    dateCell.setNumberFormat("yyyy-mm-dd");



  }



}



function recalculateAllLearningStates() {



  const ss = SpreadsheetApp.getActiveSpreadsheet();



  const todaySheet = ss.getSheetByName(trackerConfig\_().sheets.today);



  if (!todaySheet) {



    SpreadsheetApp.getUi().alert("학습기록 시트를 찾을 수 없습니다.");



    return;



  }



  const lastRow = todaySheet.getLastRow();



  if (lastRow < 2) {



    SpreadsheetApp.getUi().alert("다시 계산할 학습 문장이 없습니다.");



    return;



  }



  const ids = todaySheet



    .getRange(2, 2, lastRow - 1, 1)



    .getValues()



    .flat()



    .filter(Boolean);



  [...new Set(ids)].forEach(sentenceId => {



    updateLearningStateBySentenceId\_(ss, String(sentenceId));



  });



  SpreadsheetApp.flush();



  SpreadsheetApp.getUi().alert("학습상태 다시 계산 완료");



}



/\* =========================================================



   데이터 행 시각화



   원칙:



   - 빈 데이터 행: 배경색 없음



   - 실제 데이터가 들어간 행: 시트 용도에 맞는 색 적용



   - 말하기연습 체크 완료: 조건부서식으로 초록색



\========================================================= \*/



function syncDataRowVisuals\_() {



  const ss = SpreadsheetApp.getActiveSpreadsheet();



  const cfg = trackerConfig\_().sheets;



  const rawSheet = ss.getSheetByName(cfg.raw);



  const bankSheet = ss.getSheetByName(cfg.bank);



  const repeatSheet = ss.getSheetByName(cfg.repeat);



  const writingSheet = ss.getSheetByName(cfg.writing);



  const todaySheet = ss.getSheetByName(cfg.today);



  if (rawSheet) {



    clearDataAreaBackgrounds\_(rawSheet, 7);



    styleExistingRowsByAnyData\_(rawSheet, 1, 7, styleRawRows\_);



  }



  if (bankSheet) {



    clearDataAreaBackgrounds\_(bankSheet, 10);



    styleExistingRowsByKey\_(bankSheet, 1, styleBankRows\_);



  }



  if (repeatSheet) {



    clearDataAreaBackgrounds\_(repeatSheet, 13);



    styleExistingRowsByKey\_(repeatSheet, 1, styleRepeatRows\_);



  }



  if (writingSheet) {



    clearDataAreaBackgrounds\_(writingSheet, 10);



    styleExistingRowsByKey\_(writingSheet, 1, styleWritingRows\_);



  }



  if (todaySheet) {



    clearDataAreaBackgrounds\_(todaySheet, 10);



    styleExistingRowsByKey\_(todaySheet, 2, styleTodayRows\_);



  }



}



function styleExistingRowsByAnyData\_(sheet, startCol, numCols, styler) {



  const lastRow = sheet.getLastRow();



  if (lastRow < 2) return;



  const values = sheet.getRange(2, startCol, lastRow - 1, numCols).getValues();



  let runStart = 0;



  let runLength = 0;



  const flushRun = () => {



    if (runLength > 0) {



      styler(sheet, runStart, runLength);



      runStart = 0;



      runLength = 0;



    }



  };



  values.forEach((row, i) => {



    const actualRow = i + 2;



    const hasData = row\.some(value => String(value === null ? "" : value).trim() !== "");



    if (hasData) {



      if (runLength === 0) runStart = actualRow;



      runLength++;



    } else {



      flushRun();



    }



  });



  flushRun();



}



function styleExistingRowsByKey\_(sheet, keyCol, styler) {



  const lastRow = sheet.getLastRow();



  if (lastRow < 2) return;



  const values = sheet.getRange(2, keyCol, lastRow - 1, 1).getValues();



  let runStart = 0;



  let runLength = 0;



  const flushRun = () => {



    if (runLength > 0) {



      styler(sheet, runStart, runLength);



      runStart = 0;



      runLength = 0;



    }



  };



  values.forEach((row, i) => {



    const actualRow = i + 2;



    const hasKey = String(row[0] || "").trim() !== "";



    if (hasKey) {



      if (runLength === 0) runStart = actualRow;



      runLength++;



    } else {



      flushRun();



    }



  });



  flushRun();



}



function clearDataAreaBackgrounds\_(sheet, numCols) {



  if (!sheet) return;



  const maxRows = Math.max(sheet.getMaxRows(), 2);



  sheet.getRange(2, 1, maxRows - 1, numCols).setBackground(null);



}



function clearRowBackground\_(sheet, row, numCols) {



  sheet.getRange(row, 1, 1, numCols).setBackground(null);



}



function rowHasAnyData\_(sheet, row, startCol, numCols) {



  return sheet



    .getRange(row, startCol, 1, numCols)



    .getValues()[0]



    .some(value => String(value === null ? "" : value).trim() !== "");



}



function styleRawRows\_(sheet, startRow, numRows) {



  if (!sheet || numRows <= 0) return;



  const c = trackerStyleConfig\_();



  sheet.getRange(startRow, 1, numRows, 7).setBackground(c.base);



  sheet.getRange(startRow, 4, numRows, 1).setBackground(c.blue);



  sheet.getRange(startRow, 5, numRows, 1).setBackground(c.rawText);



}



function styleBankRows\_(sheet, startRow, numRows) {



  if (!sheet || numRows <= 0) return;



  const c = trackerStyleConfig\_();



  sheet.getRange(startRow, 1, numRows, 10).setBackground(c.base);



  sheet.getRange(startRow, 7, numRows, 1).setBackground(c.input);



  sheet.getRange(startRow, 10, numRows, 1).setBackground(c.delete);



}



function styleRepeatRows\_(sheet, startRow, numRows) {



  if (!sheet || numRows <= 0) return;



  const c = trackerStyleConfig\_();



  sheet.getRange(startRow, 1, numRows, 13).setBackground(c.base);



  sheet.getRange(startRow, 6, numRows, 6).setBackground(c.input);



  sheet.getRange(startRow, 12, numRows, 1).setBackground(c.gray);



}



function styleWritingRows\_(sheet, startRow, numRows) {



  if (!sheet || numRows <= 0) return;



  const c = trackerStyleConfig\_();



  sheet.getRange(startRow, 1, numRows, 10).setBackground(c.base);



  sheet.getRange(startRow, 5, numRows, 1).setBackground(c.input);



  sheet.getRange(startRow, 6, numRows, 1).setBackground(c.blue);



  sheet.getRange(startRow, 7, numRows, 1).setBackground(c.input);



  sheet.getRange(startRow, 8, numRows, 1).setBackground(c.blue);



  sheet.getRange(startRow, 9, numRows, 1).setBackground(c.input);



}



function styleTodayRows\_(sheet, startRow, numRows) {



  if (!sheet || numRows <= 0) return;



  const c = trackerStyleConfig\_();



  sheet.getRange(startRow, 1, numRows, 10).setBackground(c.base);



  sheet.getRange(startRow, 5, numRows, 5).setBackground(c.gray);



}



function ensureLearningConditionalFormatting\_() {



  const ss = SpreadsheetApp.getActiveSpreadsheet();



  const repeatSheet = ss.getSheetByName(trackerConfig\_().sheets.repeat);



  if (!repeatSheet) return;



  const maxRows = Math.max(repeatSheet.getMaxRows(), 2);



  const existingRules = repeatSheet.getConditionalFormatRules();



  const keptRules = existingRules.filter(rule => {



    return !rule.getRanges().some(range => {



      const firstCol = range.getColumn();



      const lastCol = range.getLastColumn();



      const lastRow = range.getLastRow();



      return firstCol <= 11 && lastCol >= 6 && lastRow >= 2;



    });



  });



  const repeatDoneRule = SpreadsheetApp.newConditionalFormatRule()



    .whenFormulaSatisfied('=AND($B2<>"",F2=TRUE)')



    .setBackground(trackerStyleConfig\_().done)



    .setRanges([repeatSheet.getRange(2, 6, maxRows - 1, 5)])



    .build();



  const recordingDoneRule = SpreadsheetApp.newConditionalFormatRule()



    .whenFormulaSatisfied('=AND($B2<>"",K2=TRUE)')



    .setBackground(trackerStyleConfig\_().done)



    .setRanges([repeatSheet.getRange(2, 11, maxRows - 1, 1)])



    .build();



  repeatSheet.setConditionalFormatRules(keptRules.concat([repeatDoneRule, recordingDoneRule]));



}



function findLastDataRowInRange\_(sheet, startCol, numCols) {



  const lastRow = sheet.getLastRow();



  if (lastRow < 2) return 0;



  const values = sheet.getRange(2, startCol, lastRow - 1, numCols).getValues();



  for (let i = values.length - 1; i >= 0; i--) {



    const hasData = values[i].some(value => String(value === null ? "" : value).trim() !== "");



    if (hasData) return i + 2;



  }



  return 0;



}



/\* =========================================================



   조회 보조 함수



\========================================================= \*/



function getExistingBank\_(sheet) {



  const lastRow = sheet.getLastRow();



  const result = {};



  if (lastRow < 2) return result;



  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();



  values.forEach(row => {



    const sentenceId = row[0];



    const sentence = row[2];



    if (sentenceId && sentence) {



      result[normalizeSentenceKey\_(sentence)] = sentenceId;



    }



  });



  return result;



}



function getNextRealRow\_(sheet, keyCol) {



  const lastRow = sheet.getLastRow();



  if (lastRow < 2) return 2;



  const values = sheet.getRange(2, keyCol, lastRow - 1, 1).getValues();



  let lastDataRow = 1;



  values.forEach((row, i) => {



    if (row[0] !== "" && row[0] !== null) {



      lastDataRow = i + 2;



    }



  });



  return lastDataRow + 1;



}



function findRowByValue\_(sheet, col, value) {



  const lastRow = sheet.getLastRow();



  if (lastRow < 2) return 0;



  const values = sheet.getRange(2, col, lastRow - 1, 1).getValues();



  for (let i = 0; i < values.length; i++) {



    if (String(values[i][0] || "").trim() === String(value || "").trim()) {



      return i + 2;



    }



  }



  return 0;



}



/\* =========================================================



   구조 점검



\========================================================= \*/



function validateTrackerStructure() {



  const ss = SpreadsheetApp.getActiveSpreadsheet();



  const cfg = trackerConfig\_().sheets;



  const requiredSheets = [



    cfg.raw,



    cfg.bank,



    cfg.repeat,



    cfg.writing,



    cfg.today,



    cfg.dashboard,



    cfg.reset,



    cfg.settings



  ];



  const missing = requiredSheets.filter(name => !ss.getSheetByName(name));



  if (missing.length > 0) {



    SpreadsheetApp.getUi().alert("없는 시트:\n\n" + missing.join("\n"));



    return;



  }



  SpreadsheetApp.getUi().alert(



    "시트 구조 정상\n\n" +



    "자막입력 → 문장은행 → 말하기연습/작문연습 → 학습기록 → 학습대시보드"



  );



}



function refreshLearningDashboard() {



  const ss = SpreadsheetApp.getActiveSpreadsheet();



  const dashboardSheet = ss.getSheetByName(trackerConfig\_().sheets.dashboard);



  if (!dashboardSheet) {



    SpreadsheetApp.getUi().alert("학습대시보드 시트를 찾을 수 없습니다.");



    return;



  }



  SpreadsheetApp.flush();



  ss.setActiveSheet(dashboardSheet);



  SpreadsheetApp.getUi().alert(



    "학습대시보드를 새로고침했습니다.\n\n" +



    "카테고리별 수집 문장수·학습 시작·완료율을 확인하세요."



  );



}



/\* =========================================================



   기존 버튼 호환용



   기존 clearRawSubtitleData 연결 버튼을 지원하기 위해



   rollbackLastSubtitleInput 함수를 호출한다.



\========================================================= \*/



function clearRawSubtitleData() {



  rollbackLastSubtitleInput();



}



/\* =========================================================



   전체 학습데이터 초기화



\========================================================= \*/



function resetAllLearningData() {



  const ss = SpreadsheetApp.getActiveSpreadsheet();



  const ui = SpreadsheetApp.getUi();



  const first = ui.alert(



    "전체 학습데이터 초기화",



    "학습 데이터를 모두 삭제하고 골조만 남깁니다.\n\n" +



    "삭제 대상:\n" +



    "자막입력 / 문장은행 / 말하기연습 / 작문연습 / 학습기록\n\n" +



    "계속하시겠습니까?",



    ui.ButtonSet.YES_NO



  );



  if (first !== ui.Button.YES) return;



  const second = ui.prompt(



    "최종 확인",



    "정말 초기화하려면 아래 입력칸에 정확히 '초기화'라고 입력하세요.",



    ui.ButtonSet.OK_CANCEL



  );



  if (second.getSelectedButton() !== ui.Button.OK) return;



  if (String(second.getResponseText() || "").trim() !== "초기화") {



    ui.alert("입력값이 일치하지 않아 초기화를 취소했습니다.");



    return;



  }



  const cfg = trackerConfig\_().sheets;



  const rawSheet = ss.getSheetByName(cfg.raw);



  const bankSheet = ss.getSheetByName(cfg.bank);



  const repeatSheet = ss.getSheetByName(cfg.repeat);



  const writingSheet = ss.getSheetByName(cfg.writing);



  const todaySheet = ss.getSheetByName(cfg.today);



  if (!rawSheet || !bankSheet || !repeatSheet || !writingSheet || !todaySheet) {



    ui.alert("필수 시트를 찾을 수 없어 초기화를 중단했습니다.");



    return;



  }



  clearDataRows\_(rawSheet, 7);



  clearDataRows\_(bankSheet, 10);



  clearDataRows\_(repeatSheet, 13);



  clearDataRows\_(writingSheet, 10);



  clearDataRows\_(todaySheet, 10);



  clearDataAreaBackgrounds\_(rawSheet, 7);



  clearDataAreaBackgrounds\_(bankSheet, 10);



  clearDataAreaBackgrounds\_(repeatSheet, 13);



  clearDataAreaBackgrounds\_(writingSheet, 10);



  clearDataAreaBackgrounds\_(todaySheet, 10);



  applyCheckboxesPreservingValues\_(repeatSheet, 6, 6);



  applyCheckboxesPreservingValues\_(writingSheet, 7, 1);



  applyCheckboxesPreservingValues\_(writingSheet, 9, 1);



  applyCheckboxesPreservingValues\_(todaySheet, 6, 4);



  applyBankDeleteCheckboxes\_(bankSheet);



  ensureLearningConditionalFormatting\_();



  SpreadsheetApp.flush();



  ui.alert(



    "전체 초기화 완료\n\n" +



    "시트 구조와 설정은 유지하고 학습 데이터만 삭제했습니다.\n" +



    "이제 자막입력부터 새로 시작할 수 있습니다."



  );



}



function clearDataRows\_(sheet, numCols) {



  const maxRows = Math.max(sheet.getMaxRows(), 2);



  sheet.getRange(2, 1, maxRows - 1, numCols).clearContent();



}
