
document.addEventListener("DOMContentLoaded", () => {
  // GA4 이벤트 전송
  const track = (eventName, parameters = {}) => {
    if (typeof gtag === "function") {
      gtag("event", eventName, parameters);
    }
  };

  // 기존 랜딩페이지 섹션 도달 이벤트
  const sectionEvents = [
    { id: "flow", eventName: "view_learning_flow" },
    { id: "dashboard", eventName: "view_dashboard" },
    { id: "how-to-use", eventName: "view_how_to_use" },
    { id: "data-management", eventName: "view_data_management" },
    { id: "closing", eventName: "view_closing" }
  ];

  const sentSectionEvents = new Set();

  const sectionObserver = new IntersectionObserver((entries) => {
    entries.forEach((entry) => {
      if (!entry.isIntersecting) return;

      const matched = sectionEvents.find(
        (item) => item.id === entry.target.id
      );

      if (!matched || sentSectionEvents.has(matched.eventName)) return;

      track(matched.eventName);
      sentSectionEvents.add(matched.eventName);
    });
  }, { threshold: 0.4 });

  sectionEvents.forEach(({ id }) => {
    const section = document.getElementById(id);
    if (section) sectionObserver.observe(section);
  });

  const openSurveyButton = document.getElementById("open-survey");
  const surveySection = document.getElementById("survey");
  const surveyForm = document.getElementById("survey-form");
  const templateCopyButton = document.getElementById("template-copy-button");
  const submitButton = surveyForm?.querySelector(".survey-submit-button");
  const surveySuccess = document.getElementById("survey-success");

  const SURVEY_API_URL =
    "https://script.google.com/macros/s/AKfycbxANPk-3T4FToVLvwStg7LWPW4-CGN0l1_nhztt5tckaAQAA7nbDMiKmPbgkcudcesk/exec";

  const questions = surveyForm
    ? Array.from(
        surveyForm.querySelectorAll(
          ".survey-question[data-question-number]"
        )
      )
    : [];

  const answeredOnce = new Set();
  const viewedOnce = new Set();

  let surveyOpened = false;
  let attemptNumber = 0;
  let lastAnsweredQuestion = 0;
  let submitting = false;

  const questionNumber = (question) =>
    Number(question?.dataset.questionNumber || 0);

  const questionAnswered = (question) => {
    if (!question) return false;

    const textarea = question.querySelector("textarea");

    if (textarea) {
      return textarea.value.trim().length > 0;
    }

    return Boolean(
      question.querySelector(
        'input[type="radio"]:checked, input[type="checkbox"]:checked'
      )
    );
  };

  const firstMissingRequiredQuestion = () => {
  for (const question of questions) {
    if (
      questionNumber(question) <= 8 &&
      !questionAnswered(question)
    ) {
      return questionNumber(question);
    }
  }

  return 0;
};

const missingRequiredCount = () =>
  questions.filter(
    (question) =>
      questionNumber(question) <= 8 &&
      !questionAnswered(question)
  ).length;





  const answeredCount = () =>
    questions.filter(questionAnswered).length;

  // 문항별 화면 노출 추적
  const questionObserver = new IntersectionObserver((entries) => {
    if (!surveyOpened) return;

    entries.forEach((entry) => {
      const number = questionNumber(entry.target);

      if (!entry.isIntersecting || viewedOnce.has(number)) return;

      viewedOnce.add(number);

      track("survey_question_view", {
        question_number: number
      });
    });
  }, { threshold: 0.4 });

  questions.forEach((question) => {
    questionObserver.observe(question);
  });

  // 설문 열기
  if (openSurveyButton && surveySection) {
    openSurveyButton.addEventListener("click", () => {
      track("template_cta_click");

      surveySection.hidden = false;

      surveySection.scrollIntoView({
        behavior: "smooth",
        block: "start"
      });

      if (!surveyOpened) {
        surveyOpened = true;
        track("survey_start");
      }
    });
  }

  // 템플릿 복사 버튼
  if (templateCopyButton) {
    templateCopyButton.addEventListener("click", () => {
      track("template_copy_click");
    });
  }

  if (!surveyForm || !submitButton) return;

  // 문항별 최초 응답 추적
  const recordAnswer = (target) => {
    const question = target.closest(
      ".survey-question[data-question-number]"
    );

    if (!question) return;

    const number = questionNumber(question);

    if (!questionAnswered(question)) return;

    lastAnsweredQuestion = number;

    if (!answeredOnce.has(number)) {
      answeredOnce.add(number);

      track("survey_question_answered", {
        question_number: number
      });
    }
  };

  // 체크박스 선택 및 상호 배타적 항목 처리
  surveyForm.addEventListener("change", (event) => {
    if (event.target.matches('input[type="checkbox"]')) {
      const field = event.target.name;
        // Q7: 최대 2개 선택 제한
    if (
      field === "interested_features" &&
      event.target.checked
    ) {
      const selected = surveyForm.querySelectorAll(
        'input[name="interested_features"]:checked'
      );

      const noFeature =
        "특별히 도움이 될 것 같은 기능이 없다";

      if (
        selected.length > 2 &&
        event.target.value !== noFeature
      ) {
        event.target.checked = false;

        alert(
          "최대 2개까지 선택할 수 있습니다.\n" +
          "다른 기능을 선택하려면 기존 선택을 해제해 주세요."
        );

        return;
      }
    }

      const exclusiveValues = {
        study_management: [
          "영상을 시청하고 별도로 기록하지 않는다.",
          "유튜브로 영어 공부를 하지 않는다."
        ],
        study_pain_points: [
          "특별히 불편한 점은 없다.",
          "유튜브로 영어 공부를 해본 적이 없다."
        ],
        interested_features: [
          "특별히 도움이 될 것 같은 기능이 없다"
        ]
      };

      const exclusives = exclusiveValues[field] || [];

      if (event.target.checked && exclusives.length) {
        surveyForm
          .querySelectorAll(
            `input[type="checkbox"][name="${field}"]`
          )
          .forEach((input) => {
            if (input === event.target) return;

            if (
              exclusives.includes(event.target.value) ||
              exclusives.includes(input.value)
            ) {
              input.checked = false;
            }
          });
      }
    }

    recordAnswer(event.target);
  });

  surveyForm.addEventListener("input", (event) => {
    if (event.target.matches("textarea")) {
      recordAnswer(event.target);
    }
  });

  // 제출 버튼 클릭 및 필수 문항 누락 추적
  submitButton.addEventListener("click", () => {
    if (submitting) return;

    attemptNumber += 1;

    track("survey_submit_attempt", {
      attempt_number: attemptNumber,
      answered_count: answeredCount(),
      last_answered_question: lastAnsweredQuestion
    });

    const missing = missingRequiredCount();

    if (missing > 0) {
      track("survey_validation_error", {
        attempt_number: attemptNumber,
        first_missing_question: firstMissingRequiredQuestion(),
        missing_count: missing
      });
    }
  });

  // 설문 제출
  surveyForm.addEventListener("submit", async (event) => {
    event.preventDefault();

    if (submitting) return;

    // Enter 키로 제출하는 경우도 추적
    if (attemptNumber === 0) {
      attemptNumber = 1;

      track("survey_submit_attempt", {
        attempt_number: attemptNumber,
        answered_count: answeredCount(),
        last_answered_question: lastAnsweredQuestion
      });
    }

   const firstMissing = firstMissingRequiredQuestion();

if (firstMissing > 0) {
  alert(`Q${firstMissing} 문항에 응답해 주세요.`);

  const missingQuestion = questions.find(
    (question) => questionNumber(question) === firstMissing
  );

  missingQuestion?.scrollIntoView({
    behavior: "smooth",
    block: "center"
  });

  return;
}

submitting = true;

    submitButton.disabled = true;
    submitButton.textContent = "제출 중...";

    const formData = new FormData(surveyForm);

    // 확정된 9개 설문 문항 데이터
    const fields = {
      source: formData.get("source") || "",
      study_purpose: formData.get("study_purpose") || "",
      study_methods: formData.getAll("study_methods").join(" | "),
      study_management: formData.getAll("study_management").join(" | "),
      study_device: formData.get("study_device") || "",
      study_pain_points: formData.getAll("study_pain_points").join(" | "),
      interested_features: formData.getAll("interested_features").join(" | "),
      usage_intent: formData.get("usage_intent") || "",
      feedback: formData.get("feedback") || ""
    };

    let hiddenFrame;
    let postForm;

    try {
      // Google Apps Script 웹앱으로 설문 데이터 전송
      hiddenFrame = document.createElement("iframe");
      hiddenFrame.name = "survey-submit-frame";
      hiddenFrame.style.display = "none";

      document.body.appendChild(hiddenFrame);

      postForm = document.createElement("form");
      postForm.method = "POST";
      postForm.action = SURVEY_API_URL;
      postForm.target = hiddenFrame.name;
      postForm.style.display = "none";

      Object.entries(fields).forEach(([name, value]) => {
        const input = document.createElement("input");

        input.type = "hidden";
        input.name = name;
        input.value = value;

        postForm.appendChild(input);
      });

      document.body.appendChild(postForm);
      postForm.submit();

      // 기존 전송 방식 유지
      // 실제 서버 저장 성공 여부는 별도로 확인해야 함
      await new Promise((resolve) => {
        setTimeout(resolve, 1500);
      });

      track("survey_complete");

      submitButton.hidden = true;

      if (surveySuccess) {
        surveySuccess.hidden = false;

        surveySuccess.scrollIntoView({
          behavior: "smooth",
          block: "center"
        });
      }
    } catch (error) {
      console.error("Survey submission failed:", error);

      submitButton.disabled = false;
      submitButton.textContent = "설문 제출하기";
      submitting = false;

      alert(
        "설문 제출 중 문제가 발생했습니다. 잠시 후 다시 시도해 주세요."
      );
    } finally {
      postForm?.remove();
      hiddenFrame?.remove();
    }
  });
});
