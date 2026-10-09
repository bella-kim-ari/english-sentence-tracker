# English Sentence Tracker

**User Behavior & Survey Analytics Project**

A data analytics portfolio project designed to understand English learners' needs through user behavior tracking and survey analysis.

The project consists of a Google Sheets-based English learning workbook, a landing page, and a data collection pipeline using GA4 and Google Apps Script.

**Tech Stack:** HTML, CSS, JavaScript, Google Apps Script, Google Sheets, Google Analytics 4, Git, GitHub

*Detailed project information is provided in Korean below.*

---

## 1. 프로젝트 소개

English Sentence Tracker는 유튜브 영어 자막에서 학습할 문장을 정리하고, 말하기와 작문 연습을 기록할 수 있도록 만든 Google Sheets 기반 학습 워크북입니다.

개인적인 영어 학습의 불편함을 해결하기 위해 제작했으며, 이후 사용자 행동과 설문 응답을 통해 실제 학습자의 요구와 제품의 개선 방향을 검증하는 **데이터 분석 프로젝트**로 확장했습니다.

**핵심 목표:** 사용자 행동 데이터와 설문 응답 데이터를 수집·분석하여 제품의 유용성과 개선 방향을 도출하는 것입니다.

## 2. 주요 기능

| 구성 요소 | 주요 기능 |
|---|---|
| 학습 워크북 템플릿| 자막 문장 추출 자동화, 문장은행 관리, 말하기·작문 연습, 학습현황 대시보드 |
| 랜딩페이지 | 워크북 소개 및 주요 기능 안내 |
| 설문조사 | 영어 학습 경험, 불편함, 관심 기능, 사용 의향 조사 (9문항) |
| GA4 | 페이지 방문, 스크롤, CTA 클릭, 설문 참여 과정 등의 행동 이벤트 추적 |
| Google Apps Script | 학습 워크북 자동화 및 설문 응답 저장 |

## 3. 데이터 수집 구조

```text
사용자
  ↓
랜딩페이지
  ├── GA4
  │     └── 사용자 행동 이벤트 수집
  │
  └── 설문조사
        ↓
      JavaScript
        ↓
      Google Apps Script
        ↓
      Google Sheets (responses)
        ↓
      데이터 정제 및 분석
        ↓
      사용자 요구 파악 및 개선 방향 도출
```

설문 완료 후 사용자는 Google Sheets 학습 워크북의 사본을 만들 수 있습니다.

GA4 행동 데이터와 설문 응답 데이터는 별도로 수집하며, 공통 식별자가 없는 경우 집계 수준에서 비교합니다.

## 4. 기술 스택

| 기술 | 역할 |
|---|---|
| HTML / CSS / JavaScript | 랜딩페이지 및 설문 인터페이스 |
| Google Apps Script | 설문 데이터 저장 및 학습 기능 자동화 |
| Google Sheets | 설문 응답 데이터 저장 및 학습 워크북 |
| Google Analytics 4 | 사용자 행동 이벤트 수집 |
| Git / GitHub | 소스 코드 및 버전 관리 |
| Cloudflare Pages | 웹사이트 배포 |

## 5. 프로젝트 구조

```text
english-learning-landing/
├── apps-script/
│   ├── English_Sentence_Tracker_Apps_Script.gs
│   └── survey_Apps_Script.gs
├── public/
│   ├── images/
│   ├── index.html
│   ├── script.js
│   └── style.css
└── README.md
```

`apps-script/`에는 학습 워크북 자동화와 설문 응답 저장을 위한 두 개의 독립적인 스크립트가 포함되어 있습니다.

## 6. 데이터 분석 방향

- **사용자 행동 분석:** GA4 이벤트를 활용한 랜딩페이지 탐색 및 설문 참여 과정 분석
- **설문 응답 분석:** 영어 학습 목적, 주요 불편함, 관심 기능 및 사용 의향 분석
- **개선 방향 도출:** 분석 결과를 바탕으로 사용자 요구와 제품 개선 과제 도출

실제 데이터 수집 후 주요 분석 결과와 시각화 자료를 추가할 예정입니다.

---

**Project:** English Sentence Tracker  
**Focus:** Product Analytics · GA4 · User Behavior Analysis · Survey Analysis
