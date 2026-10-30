/* =====================================================================
 *  사이트 설정 파일 — 사이트에 보이는 모든 글과 항목은 여기서만 고치면 됩니다.
 *  · 따옴표("…") 안의 글자만 바꾸세요.
 *  · 항목을 늘리려면 { … } 한 덩어리를 복사해 붙여 넣고, 줄이려면 지우면 됩니다.
 *  · 저장 후 브라우저에서 새로고침(F5)하면 바로 반영됩니다.
 *  · 화면 순서: 첫 화면 → 통계 → 프로그램 소개 → 교육 일정(주차·달력)
 *               → 수강 안내 → 참여하기 → FAQ → 공지사항 → 교수자(푸터)
 *  · 52주 교육 일정은 엑셀(data/schedule.xlsx)에서 만들어집니다.
 *    엑셀을 고친 뒤 `python tools/excel_to_config.py` 를 실행하면 아래 주차 목록이 다시 써집니다.
 * ===================================================================== */
window.SITE_CONFIG = {
  /* ── 기본 정보 ─────────────────────────────────────────── */
  site: {
    university: "모산학술연구소",
    department: "",
    courseTitle: "모산 with AI",
    version: "v1.0", // 제목 오른쪽에 빨간 글씨로 표시. 비워 두면 표시하지 않습니다.
    logoImage: "", // 로고 이미지 경로. 비워 두면 아래 이모지가 로고로 쓰입니다.
    logoEmoji: "🌿",
    // 첫 화면·브라우저 탭의 제목 옆에 붙는 소속 줄. 비워 두면 위의 "대학교 학과"가 표시됩니다.
    brandSub: "모산학술연구소 · 시니어 AI 아카데미",
    // 헤더 왼쪽 위 제목의 윗줄. 비워 두면 위의 소속 줄(brandSub)이 표시됩니다.
    headerSub: "모산과 함께하는 AI 집중 교육 프로그램",
  },

  /* ── 헤더 메뉴 (id는 아래 각 섹션의 id와 같아야 합니다) ── */
  nav: [
    { id: "about", label: "프로그램 소개" },
    { id: "curriculum", label: "교육 일정" },
    { id: "guide", label: "수강 안내" },
    { id: "participate", label: "참여하기" },
    { id: "faq", label: "FAQ" },
    { id: "instructor", label: "교수자" },
  ],

  /* ── 첫 화면 ───────────────────────────────────────────── */
  hero: {
    badge: "2026–2027 시니어 AI 아카데미 · 연간 온라인 과정",
    subtitle: "모산과 함께하는 *AI 집중 교육* 프로그램",
    // 오른쪽 곡선 액자 안의 글(*…*는 굵게)과 모서리 색 상자
    artText: "AI와 함께 여는 *새로운 내일.*",
    artBadge: { value: "52", label: "WEEKS" },
    description:
      "AI 역량을 강화하려는 사람을 위한 1년 52주 AI 교육 과정입니다. 매주 목요일 저녁 화상회의로 만나 AI 기초부터 프롬프트, 콘텐츠 제작, 프로젝트, 자동화, 바이브코딩 웹앱까지 강의·시연·실습으로 차근차근 익히고, 마지막에는 나만의 AI 웹앱을 완성해 발표합니다.",
    buttons: [
      // href가 "#"으로 시작하면 페이지 안 이동, 주소(https://…)면 새 창으로 열립니다.
      { label: "수강 신청", href: "#apply", primary: true },
      { label: "교육 일정 보기", href: "#calendar", primary: false },
      { label: "🎥 Zoom 수업 입장", href: "https://us05web.zoom.us/j/89978809823?pwd=RCMGuMJXdcaCIIamXPVHZyMlvLdCho.1", primary: false },
    ],
    // 한눈에 보기
    quickInfo: [
      { icon: "🗓️", label: "교육 기간", value: "2026. 12. 3 – 2027. 11. 25 (52주)" },
      { icon: "⏰", label: "수업 시간", value: "매주 목요일 19:00 – 21:00" },
      { icon: "💻", label: "교육 방식", value: "온라인 실시간 (화상회의)" },
      { icon: "👥", label: "교육 대상", value: "AI 역량을 강화하려는 사람" },
    ],
  },

  /* ── 공지사항 (자주 묻는 질문 아래. 카드를 누르면 내용이 펼쳐집니다) ──
   *   관리자 로그인 뒤 공지사항 섹션에서 추가·수정·삭제할 수 있습니다.
   *   pinned: true면 '중요' 표시와 함께 맨 위로. 그 밖에는 최근 날짜가 위로 옵니다. */
  notices: {
    id: "notices",
    eyebrow: "Notice",
    title: "공지사항",
    lead: "교육과 관련한 새 소식을 알려 드립니다.",
    items: [
      { title: "첫 수업은 12월 3일(목) 저녁 7시입니다", text: "운영은 12월 1일(화)부터 시작하며, 첫 정규 수업은 12월 3일(목) 19:00에 화상회의로 진행합니다.\nZoom 입장 버튼은 첫 화면과 교육 일정의 온라인 강의실 상자에 있습니다.", date: "2026-11-20", pinned: true },
      { title: "휴강·보강 안내 원칙", text: "명절·공휴일 및 강사·학습자 일정에 따라 휴강할 경우 보강일을 별도로 정해 공지하고 달력에 표시합니다.", date: "2026-11-20", pinned: false },
    ],
  },

  /* ── 숫자로 보는 강의 ──────────────────────────────────── */
  stats: [
    { value: 52, suffix: "회", label: "연간 정규 수업" },
    { value: 104, suffix: "시간", label: "총 교육시간" },
    { value: 6, suffix: "개", label: "교육 영역" },
    { value: 25, suffix: "명", label: "권장 인원 (20~25명)" },
  ],

  /* ── 프로그램 소개: 소개 + 특징 슬라이드 + 학습 목표 ── */
  about: {
    id: "about",
    eyebrow: "About",
    title: "프로그램 소개",
    lead: "AI와 함께 나의 역량을 새롭게 키우는 1년 과정입니다.",
    intro: [
      "「모산 with AI」는 모산학술연구소가 운영하는 시니어 AI 아카데미의 연간 온라인 교육 프로그램입니다. AI 역량을 강화하려는 사람이 생성형 AI를 생활과 일, 취미와 사회 참여에 직접 활용할 수 있도록 52주 동안 단계별로 안내합니다.",
      "매주 목요일 저녁 7시부터 9시까지 화상회의로 진행하며, 한 회는 개념 설명 · 시연 · 두 번의 실습 · 발표와 정리로 구성됩니다. 6개의 교육 영역을 차례로 거치며 주차마다 눈에 보이는 결과물을 하나씩 쌓아 갑니다.",
      "AI 기초와 디지털 활용에서 출발해 프롬프트 전략, 글·이미지·음성·영상 콘텐츠 제작, 개인 프로젝트, 업무 자동화와 AI 에이전트를 거쳐, 마지막에는 바이브코딩으로 나만의 웹앱을 만들어 최종 성과발표회에서 시연하고 수료합니다.",
    ],
    slidesTitle: "6개 교육 영역",
    slides: [
      { icon: "📱", title: "AI 기초와 디지털 활용 (1–8주)", text: "ChatGPT·Gemini·Claude 시작하기, 스마트폰 음성 AI, 사실 확인, 생활 문서와 생활 도우미 활용." },
      { icon: "💬", title: "프롬프트 전략과 AI 도구 (9–16주)", text: "역할·맥락·단계별 지시, 예시 기반 개선, 개인 AI 비서, PDF 분석과 NotebookLM 지식 정리." },
      { icon: "🎨", title: "AI 콘텐츠 제작 (17–26주)", text: "회고록 글쓰기, 이미지·카드뉴스·발표자료, 음성 내레이션과 음악, 영상 제작과 유튜브·블로그 운영." },
      { icon: "🧭", title: "AI 프로젝트 기획·제작 (27–34주)", text: "자신의 경험에서 주제를 찾아 기획서 작성, 자료 수집, 제작, 개선을 거쳐 중간 발표회에서 시연." },
      { icon: "⚙️", title: "AI 자동화와 에이전트 (35–43주)", text: "Make·n8n 기초, 뉴스 수집·이메일 요약·문서 생성 자동화, 시트 연동과 AI 에이전트 설계." },
      { icon: "🌐", title: "바이브코딩과 AI 웹앱 (44–52주)", text: "AI 코딩 도구로 웹앱 만들기, 큰 글씨·쉬운 화면 설계, 기능 구현과 배포, 최종 성과발표회와 수료." },
    ],
    goalsTitle: "교육 목표",
    goalsLead: "생성형 AI를 안전하고 주체적으로 활용해 생활·일·배움·사회 참여를 확장할 수 있는 실천 역량을 기릅니다.",
    goals: [
      "생성형 AI의 활용 분야와 한계를 이해하고 주요 도구를 비교해 쓸 수 있다.",
      "AI 답변을 출처로 검증하고 개인정보를 안전하게 다룰 수 있다.",
      "역할·맥락·조건을 갖춘 프롬프트로 원하는 결과를 얻을 수 있다.",
      "글·이미지·음성·음악·영상 콘텐츠를 AI로 기획하고 제작할 수 있다.",
      "자신의 경험을 바탕으로 AI 프로젝트를 기획하고 완성할 수 있다.",
      "반복 업무를 자동화 흐름으로 바꾸고 AI 에이전트를 설계할 수 있다.",
      "바이브코딩으로 시니어 친화적인 웹앱을 만들어 배포할 수 있다.",
    ],

    /* 인포그래픽 (프로그램 소개 맨 아래). 통째로 지우면 표시되지 않습니다. */
    infographic: {
      title: "한눈에 보는 모산 with AI",
      inputsTitle: "활용하는 것",
      inputs: [
        { icon: "🤖", title: "생성형 AI", text: "ChatGPT · Gemini · Claude · NotebookLM" },
        { icon: "🎨", title: "콘텐츠 도구", text: "Canva · 이미지·음성·음악·영상 생성 AI" },
        { icon: "⚙️", title: "자동화·코딩", text: "Make · n8n · Google Sheets · AI 코딩 도구" },
      ],
      outputsTitle: "만드는 것",
      outputs: [
        { icon: "📚", label: "프롬프트 라이브러리" },
        { icon: "🎬", label: "콘텐츠 포트폴리오" },
        { icon: "🌐", label: "나만의 AI 웹앱" },
      ],
      stepsTitle: "한 회 수업의 흐름 (120분)",
      steps: [
        { icon: "👋", title: "입장·안내", text: "10분" },
        { icon: "💡", title: "개념", text: "15분" },
        { icon: "🖥️", title: "시연", text: "15분" },
        { icon: "✍️", title: "실습", text: "25분 + 휴식 10분 + 30분" },
        { icon: "🎤", title: "발표·정리", text: "15분" },
      ],
      roadmapTitle: "52주 학습 흐름",
      roadmap: [
        { weeks: "1–8주", span: 8, title: "AI 기초", text: "AI 기초와 디지털 활용" },
        { weeks: "9–16주", span: 8, title: "프롬프트", text: "프롬프트 전략과 AI 도구" },
        { weeks: "17–26주", span: 10, title: "콘텐츠", text: "AI 콘텐츠 제작" },
        { weeks: "27–34주", span: 8, title: "프로젝트", text: "AI 프로젝트 기획·제작 · 중간 발표회" },
        { weeks: "35–43주", span: 9, title: "자동화", text: "AI 자동화와 에이전트" },
        { weeks: "44–52주", span: 9, title: "웹앱", text: "바이브코딩과 AI 웹앱 · 최종 발표·수료" },
      ],
      toolsTitle: "단계별 주요 도구",
      toolGroups: [
        { title: "AI 기초·프롬프트", items: ["ChatGPT", "Gemini", "Claude", "NotebookLM"] },
        { title: "콘텐츠 제작", items: ["Canva", "이미지 생성 AI", "TTS", "AI 음악·영상"] },
        { title: "자동화", items: ["Make", "n8n", "Google Sheets", "Airtable"] },
        { title: "웹앱", items: ["AI 코딩 도구", "HTML·CSS·JS", "정적 웹 배포"] },
      ],
    },
  },

  /* ── 교육 일정 (주차별 커리큘럼 + 달력) ──────────────────── */
  curriculum: {
    id: "curriculum",
    eyebrow: "Schedule",
    title: "교육 일정",
    lead: "엑셀 교육계획의 52주 일정이 그대로 반영되어 있습니다. 달력에서 날짜를 누르거나, 아래 주차를 펼쳐 학습 내용·실습 과제·결과물을 확인하세요.",
    weeksTitle: "52주 커리큘럼",
    calendarTitle: "교육 일정 캘린더",

    /* 교육 영역 — 달력·주차 목록의 색과 영역별 보기 버튼에 쓰입니다. name은 주차의 stage와 같아야 합니다. */
    stages: [
      { name: "AI 기초와 디지털 활용", short: "AI 기초", color: "#2f8fe0" },
      { name: "프롬프트 전략과 AI 도구", short: "프롬프트", color: "#12876c" },
      { name: "AI 콘텐츠 제작", short: "콘텐츠", color: "#e0752f" },
      { name: "AI 프로젝트 기획·제작", short: "프로젝트", color: "#8a5cd1" },
      { name: "AI 자동화와 에이전트", short: "자동화", color: "#c7417b" },
      { name: "바이브코딩과 AI 웹앱", short: "웹앱", color: "#b08900" },
    ],

    /* 수업 일정 기본값 — 각 주차에 date가 있으면 그 날짜를 씁니다(엑셀의 수업 날짜).
     *   submitUrl  : 과제 제출 버튼이 열 주소. "#classroom"이면 이 사이트의 과제 제출 창이 열립니다. */
    schedule: {
      firstClass: "2026-12-03",
      time: "19:00 – 21:00",
      location: "온라인 실시간",
      submitUrl: "#classroom",
    },

    /* 온라인 강의실(Zoom) — 교육 일정 위에 입장 상자가, 달력·주차·첫 화면에 입장 버튼이 표시됩니다.
     *   url은 Zoom 초대 링크(비밀번호 포함)입니다. 비워 두면 Zoom 안내를 표시하지 않습니다. */
    zoom: {
      title: "온라인 강의실 (Zoom)",
      url: "https://us05web.zoom.us/j/89978809823?pwd=RCMGuMJXdcaCIIamXPVHZyMlvLdCho.1",
      meetingId: "899 7880 9823",
      buttonLabel: "Zoom 수업 입장",
      tips: [
        "수업 10분 전(18:50)부터 입장할 수 있습니다.",
        "처음이면 Zoom 앱이 설치됩니다. 컴퓨터는 '브라우저에서 참가'로도 들어올 수 있습니다.",
        "입장할 때 이름을 실명으로 바꿔 주세요. 마이크는 끄고, 말할 때만 켭니다.",
      ],
    },

    /* 주차별 내용 (엑셀 '52주 상세일정' 시트에서 자동 생성)
     *   stage    : 교육 영역      topics   : 세부 학습 내용
     *   practice : 실습 활동·과제  outcome  : 주차별 결과물 */
    weeks: [
      /* WEEKS:BEGIN */
      {"week": 1, "title": "AI 시대와 새로운 기회", "tag": "", "stage": "AI 기초와 디지털 활용", "date": "2026-12-03", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["AI의 활용 분야와 한계 이해", "시니어 AI 활용 성공 사례 살펴보기", "개인 목표와 사전 디지털 역량 진단"], "practice": "AI로 해보고 싶은 일 3가지와 연간 학습 목표 작성", "outcome": "개인 AI 학습 목표서", "materials": [], "videos": []},
      {"week": 2, "title": "생성형 AI와 주요 도구 비교", "tag": "", "stage": "AI 기초와 디지털 활용", "date": "2026-12-10", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["생성형 AI·LLM·챗봇의 차이 이해", "ChatGPT·Gemini·Claude의 화면과 기능 비교", "무료·유료 기능과 요금 주의사항 확인"], "practice": "동일 질문을 2개 AI에 입력하고 답변 비교", "outcome": "AI 도구 비교표", "materials": [], "videos": []},
      {"week": 3, "title": "ChatGPT 시작 및 파일 활용", "tag": "", "stage": "AI 기초와 디지털 활용", "date": "2026-12-17", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["회원가입·로그인과 대화 기록 확인", "텍스트 질문 및 후속 질문 연습", "PDF·사진 첨부와 개인정보 삭제 원칙 학습"], "practice": "자신의 소개문과 공개 문서의 요약문 생성", "outcome": "첫 AI 대화 기록", "materials": [], "videos": []},
      {"week": 4, "title": "스마트폰 AI와 음성 대화", "tag": "", "stage": "AI 기초와 디지털 활용", "date": "2026-12-24", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["모바일 앱 설치와 접근성 설정", "음성 입력·대화로 질문하기", "카메라·사진 기반 사물 설명 및 번역"], "practice": "사진 1장을 분석하고 음성으로 후속 질문하기", "outcome": "스마트폰 AI 활용 카드", "materials": [], "videos": []},
      {"week": 5, "title": "AI 검색과 사실 확인", "tag": "", "stage": "AI 기초와 디지털 활용", "date": "2026-12-31", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["웹 검색과 출처 표시 이해", "AI 답변의 오류·환각 사례 확인", "정부·공공기관 자료와 교차 검증"], "practice": "관심 주제의 사실 5개를 출처로 확인", "outcome": "출처 검증표", "materials": [], "videos": []},
      {"week": 6, "title": "생활 문서 작성", "tag": "", "stage": "AI 기초와 디지털 활용", "date": "2027-01-07", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["공지문·이메일·민원 문의문 형식 익히기", "어조·분량·독자 조건 설정", "초안 검토 및 문장 교정"], "practice": "동호회 공지문과 정중한 이메일 작성", "outcome": "생활문서 2종", "materials": [], "videos": []},
      {"week": 7, "title": "AI 생활 도우미", "tag": "", "stage": "AI 기초와 디지털 활용", "date": "2027-01-14", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["여행·요리·취미·일정 계획에 AI 활용", "지출 계획 예시 만들기", "건강·금융 정보는 전문 자문 대체 불가 확인"], "practice": "개인 취미 또는 여행 1일 계획 수립", "outcome": "생활 활용 계획표", "materials": [], "videos": []},
      {"week": 8, "title": "AI 기초 복습·실습 발표", "tag": "발표", "stage": "AI 기초와 디지털 활용", "date": "2027-01-21", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["기초 기능 5가지 반복 실습", "안전한 개인정보 입력 체크리스트", "1인 3분 활용 사례 발표"], "practice": "자주 쓰는 AI 기능을 소개하는 안내서 제작", "outcome": "AI 기초 안내서", "materials": [], "videos": []},
      {"week": 9, "title": "프롬프트 기본 구조", "tag": "", "stage": "프롬프트 전략과 AI 도구", "date": "2027-01-28", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["역할·목표·상황·조건·출력 형식 이해", "모호한 질문과 구체적 질문 비교", "개선 전후 응답 차이 평가"], "practice": "하나의 질문을 3단계로 고도화", "outcome": "기본 프롬프트 5종", "materials": [], "videos": []},
      {"week": 10, "title": "역할 부여와 맥락 설계", "tag": "", "stage": "프롬프트 전략과 AI 도구", "date": "2027-02-04", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["교사·편집자·여행기획자 등 역할 설정", "대상 독자와 배경 정보 명시", "출력 분량과 언어 제약 넣기"], "practice": "자기 전문 분야의 전문가 프롬프트 작성", "outcome": "역할 프롬프트 모음", "materials": [], "videos": []},
      {"week": 11, "title": "단계별 작업 지시", "tag": "", "stage": "프롬프트 전략과 AI 도구", "date": "2027-02-11", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["복잡한 과제를 작은 단계로 나누기", "중간 점검 및 질문 요청 방식", "최종 답변의 평가 항목 만들기"], "practice": "퇴직 경험 기반 강의안을 단계별 지시로 제작", "outcome": "단계형 작업 프롬프트", "materials": [], "videos": []},
      {"week": 12, "title": "예시 기반 프롬프트와 반복 개선", "tag": "", "stage": "프롬프트 전략과 AI 도구", "date": "2027-02-18", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["원하는 예시 제시 및 모방 범위 지정", "표·체크리스트·JSON 기초 출력 요청", "피드백으로 2회 이상 개선"], "practice": "홍보문을 예시와 평가표에 맞춰 수정", "outcome": "개선 이력 포함 프롬프트", "materials": [], "videos": []},
      {"week": 13, "title": "개인화 AI 비서 만들기", "tag": "", "stage": "프롬프트 전략과 AI 도구", "date": "2027-02-25", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["개인 지침과 톤 설정", "반복 수행 업무 유형 정리", "기밀 정보 제외한 개인화 사례 적용"], "practice": "취미·독서·여행 중 한 분야 AI 비서 지침 설계", "outcome": "개인 AI 비서 설정서", "materials": [], "videos": []},
      {"week": 14, "title": "PDF·문서·표 분석 프롬프트", "tag": "", "stage": "프롬프트 전략과 AI 도구", "date": "2027-03-04", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["PDF 요약 및 질문 응답", "표 자료 분류·정리", "원본과 AI 결과의 숫자·인용 확인"], "practice": "공개 PDF 1개를 분석하여 요약과 질문 5개 작성", "outcome": "문서 분석 결과표", "materials": [], "videos": []},
      {"week": 15, "title": "NotebookLM 지식 정리", "tag": "", "stage": "프롬프트 전략과 AI 도구", "date": "2027-03-11", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["자료 업로드와 출처 기반 질문", "문서별 주제 비교", "개인 학습 자료와 요약 노트 만들기"], "practice": "공개 자료 2~3개로 학습 노트 제작", "outcome": "AI 지식 노트", "materials": [], "videos": []},
      {"week": 16, "title": "프롬프트 라이브러리 발표", "tag": "발표", "stage": "프롬프트 전략과 AI 도구", "date": "2027-03-18", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["업무·생활·취미 프롬프트 통합", "사용 상황과 주의사항 기록", "서로의 프롬프트 시연·평가"], "practice": "검증한 프롬프트 20개를 분류·제출", "outcome": "개인 프롬프트 라이브러리", "materials": [], "videos": []},
      {"week": 17, "title": "AI 글쓰기와 회고록", "tag": "", "stage": "AI 콘텐츠 제작", "date": "2027-03-25", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["인생 사건의 연대기 구성", "인터뷰형 질문으로 경험 수집", "사실과 창작 내용을 구분해 퇴고"], "practice": "회고록 한 장(800~1200자) 작성", "outcome": "개인 회고록 초고", "materials": [], "videos": []},
      {"week": 18, "title": "AI 이미지 생성 기초", "tag": "", "stage": "AI 콘텐츠 제작", "date": "2027-04-01", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["이미지 프롬프트의 피사체·구도·분위기", "생성 결과 개선", "저작권·초상권 확인"], "practice": "여행·취미 주제로 이미지 5장 제작", "outcome": "AI 이미지 작품집", "materials": [], "videos": []},
      {"week": 19, "title": "AI 포스터와 카드뉴스", "tag": "", "stage": "AI 콘텐츠 제작", "date": "2027-04-08", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["카드뉴스 메시지와 대상 설정", "Canva 레이아웃·글자 배치", "이미지와 문구의 가독성 개선"], "practice": "지역 행사 홍보 카드뉴스 4장 제작", "outcome": "홍보 카드뉴스", "materials": [], "videos": []},
      {"week": 20, "title": "AI 발표자료 제작", "tag": "", "stage": "AI 콘텐츠 제작", "date": "2027-04-15", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["발표 흐름과 핵심 메시지 정리", "AI로 10장 구성안 만들기", "이미지·도표 넣고 사실 검토"], "practice": "개인 전문 경험 발표자료 제작", "outcome": "10장 발표자료", "materials": [], "videos": []},
      {"week": 21, "title": "AI 음성과 내레이션", "tag": "", "stage": "AI 콘텐츠 제작", "date": "2027-04-22", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["텍스트를 자연스러운 낭독용으로 수정", "TTS 언어·속도·발음 조정", "음성 사용 동의와 사칭 금지 학습"], "practice": "1분 분량 개인 소개 나레이션 제작", "outcome": "오디오 파일", "materials": [], "videos": []},
      {"week": 22, "title": "AI 음악 창작", "tag": "", "stage": "AI 콘텐츠 제작", "date": "2027-04-29", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["장르·분위기·가사 주제 기획", "AI 작사·작곡 결과 비교", "사용권 및 게시 조건 확인"], "practice": "가족·여행을 주제로 짧은 노래 제작", "outcome": "AI 창작음악", "materials": [], "videos": []},
      {"week": 23, "title": "AI 영상 기획·스토리보드", "tag": "", "stage": "AI 콘텐츠 제작", "date": "2027-05-06", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["60초 영상 목적과 대상 설정", "6개 장면 스토리보드 작성", "영상에 쓸 이미지·음성 수집"], "practice": "개인 소개 영상 스토리보드 작성", "outcome": "6장면 영상 기획안", "materials": [], "videos": []},
      {"week": 24, "title": "AI 동영상 제작 실습", "tag": "", "stage": "AI 콘텐츠 제작", "date": "2027-05-13", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["장면 생성·편집과 전환 효과", "TTS 내레이션과 배경음악 삽입", "자막 정확도·가독성 점검"], "practice": "1분 자기소개 또는 취미 영상 편집", "outcome": "완성 영상 1편", "materials": [], "videos": []},
      {"week": 25, "title": "유튜브·블로그 콘텐츠 운영", "tag": "", "stage": "AI 콘텐츠 제작", "date": "2027-05-20", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["게시용 제목·설명·썸네일 작성", "콘텐츠 발행 일정 설계", "공개 범위와 댓글·개인정보 설정"], "practice": "블로그 글 1개와 쇼츠 게시안 작성", "outcome": "콘텐츠 운영 캘린더", "materials": [], "videos": []},
      {"week": 26, "title": "멀티미디어 작품 발표", "tag": "발표", "stage": "AI 콘텐츠 제작", "date": "2027-05-27", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["글·이미지·음성·영상 결과물 통합", "작품 설명서와 사용권 확인", "동료 피드백 후 수정"], "practice": "개인 콘텐츠 3종을 포트폴리오로 정리", "outcome": "AI 콘텐츠 포트폴리오", "materials": [], "videos": []},
      {"week": 27, "title": "AI 프로젝트 사례와 주제 발굴", "tag": "", "stage": "AI 프로젝트 기획·제작", "date": "2027-06-03", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["자서전·지역문화·여행·교육 사례 탐색", "자신의 경험·관심·자원 목록화", "실행 가능성·난이도 평가"], "practice": "프로젝트 후보 3개 작성 및 1개 선정", "outcome": "프로젝트 주제 선정표", "materials": [], "videos": []},
      {"week": 28, "title": "프로젝트 기획서 작성", "tag": "", "stage": "AI 프로젝트 기획·제작", "date": "2027-06-10", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["해결할 문제와 이용자 정의", "핵심 기능 및 범위 결정", "주차별 작업·성과 기준 수립"], "practice": "목표·대상·일정·산출물 포함 기획서 작성", "outcome": "프로젝트 기획서", "materials": [], "videos": []},
      {"week": 29, "title": "프로젝트 자료 수집과 정제", "tag": "", "stage": "AI 프로젝트 기획·제작", "date": "2027-06-17", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["공공·개인 자료의 출처 구분", "파일 이름·폴더 구조 통일", "개인정보·저작물 사용 가능 여부 확인"], "practice": "자료 10건 분류 및 출처 기록", "outcome": "자료 목록·출처표", "materials": [], "videos": []},
      {"week": 30, "title": "프로젝트 콘텐츠·화면 설계", "tag": "", "stage": "AI 프로젝트 기획·제작", "date": "2027-06-24", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["사용자 여정과 목차 설계", "주요 화면 또는 콘텐츠 흐름 구성", "AI로 시안 2개 생성·비교"], "practice": "프로젝트 구조도와 시안 제작", "outcome": "콘텐츠 설계서", "materials": [], "videos": []},
      {"week": 31, "title": "프로젝트 제작 1: 핵심 자료", "tag": "", "stage": "AI 프로젝트 기획·제작", "date": "2027-07-01", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["선정 주제의 원고·이미지·데이터 제작", "프롬프트 일관성 유지", "중간 결과물 품질 검사"], "practice": "핵심 콘텐츠 50% 이상 제작", "outcome": "중간 결과물 1차", "materials": [], "videos": []},
      {"week": 32, "title": "프로젝트 제작 2: 통합", "tag": "", "stage": "AI 프로젝트 기획·제작", "date": "2027-07-08", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["생성 콘텐츠 편집과 결합", "오류·중복·사실 관계 점검", "동료 시연 및 피드백 수집"], "practice": "시제품 완성 후 동료 1명 테스트", "outcome": "프로젝트 시제품", "materials": [], "videos": []},
      {"week": 33, "title": "프로젝트 개선과 완성", "tag": "", "stage": "AI 프로젝트 기획·제작", "date": "2027-07-15", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["사용성·가독성 수정", "출처·저작권 및 공개 여부 최종 점검", "사용 설명서와 소개문 작성"], "practice": "프로젝트 최종본 및 설명서 작성", "outcome": "완성 프로젝트", "materials": [], "videos": []},
      {"week": 34, "title": "중간 프로젝트 발표회", "tag": "발표", "stage": "AI 프로젝트 기획·제작", "date": "2027-07-22", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["5분 발표와 3분 질의응답", "동료 평가표 작성", "향후 자동화·웹앱 확장 아이디어 정리"], "practice": "프로젝트 시연 및 피드백 반영 계획 제출", "outcome": "중간 성과발표", "materials": [], "videos": []},
      {"week": 35, "title": "AI 자동화 개념과 설계", "tag": "", "stage": "AI 자동화와 에이전트", "date": "2027-07-29", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["트리거·작업·조건의 개념", "수동 업무를 흐름도로 변환", "자동 실행 위험과 승인 단계 파악"], "practice": "개인 반복 업무 3개를 자동화 후보로 작성", "outcome": "업무 자동화 후보표", "materials": [], "videos": []},
      {"week": 36, "title": "Make·n8n 기초", "tag": "", "stage": "AI 자동화와 에이전트", "date": "2027-08-05", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["자동화 서비스 계정 및 화면 소개", "모듈·노드 연결 원리", "테스트 실행과 실행 기록 확인"], "practice": "예제 트리거→메시지 저장 흐름 구성", "outcome": "첫 자동화 워크플로", "materials": [], "videos": []},
      {"week": 37, "title": "뉴스·정보 수집 자동화", "tag": "", "stage": "AI 자동화와 에이전트", "date": "2027-08-12", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["RSS·공개 정보 소스 선정", "키워드 필터와 중복 제거", "요약문을 저장하거나 전송하는 흐름 설계"], "practice": "관심 뉴스 3건을 자동 수집해 요약", "outcome": "뉴스 수집 자동화", "materials": [], "videos": []},
      {"week": 38, "title": "이메일 요약·알림", "tag": "", "stage": "AI 자동화와 에이전트", "date": "2027-08-19", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["메일 접근 권한 최소화", "특정 조건의 메일 분류·요약", "발송 전 승인 및 오류 알림 구성"], "practice": "테스트 메일을 대상으로 요약 흐름 구성", "outcome": "이메일 요약 시나리오", "materials": [], "videos": []},
      {"week": 39, "title": "문서 생성 자동화", "tag": "", "stage": "AI 자동화와 에이전트", "date": "2027-08-26", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["폼 입력→내용 정리→문서 초안 제작", "템플릿 필드 치환", "출력값 확인과 오탈자 검토"], "practice": "행사 공지문 자동 생성 시범", "outcome": "문서 자동화 템플릿", "materials": [], "videos": []},
      {"week": 40, "title": "Google Sheets·Airtable 연동", "tag": "", "stage": "AI 자동화와 에이전트", "date": "2027-09-02", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["테이블 열·행과 데이터 형식 이해", "입력값 자동 기록", "누락·중복 행 처리와 보안"], "practice": "간단한 독서기록 표 자동 업데이트", "outcome": "기록 자동화 시트", "materials": [], "videos": []},
      {"week": 41, "title": "AI 에이전트와 도구 사용", "tag": "", "stage": "AI 자동화와 에이전트", "date": "2027-09-09", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["에이전트·일반 챗봇 차이", "계획·검색·도구 실행의 역할", "권한·비용·확인 단계를 설계"], "practice": "개인 정보 도우미 에이전트 사용 시나리오 작성", "outcome": "에이전트 설계도", "materials": [], "videos": []},
      {"week": 42, "title": "자동화 개인 실습 프로젝트", "tag": "", "stage": "AI 자동화와 에이전트", "date": "2027-09-16", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["수집→분류→요약→저장 워크플로 구성", "실행 오류 점검", "개인정보·비용·실행 제한 설정"], "practice": "개인 자동화 흐름 1개를 처음부터 구현", "outcome": "작동하는 워크플로", "materials": [], "videos": []},
      {"week": 43, "title": "자동화 시연 및 개선", "tag": "발표", "stage": "AI 자동화와 에이전트", "date": "2027-09-23", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["자동화 결과 비교와 검증", "실패 시 재시도·알림 설정", "시연 영상 또는 설명서 제작"], "practice": "내 자동화 3분 시연 및 개선점 정리", "outcome": "자동화 프로젝트 완성", "materials": [], "videos": []},
      {"week": 44, "title": "웹앱과 바이브코딩 이해", "tag": "", "stage": "바이브코딩과 AI 웹앱", "date": "2027-09-30", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["웹사이트·웹앱 차이", "시니어용 웹앱 사례 살펴보기", "개발 도구와 접근성·보안 원칙"], "practice": "만들고 싶은 웹앱 아이디어 3개 작성", "outcome": "개인 웹앱 주제표", "materials": [], "videos": []},
      {"week": 45, "title": "첫 바이브코딩 실습", "tag": "", "stage": "바이브코딩과 AI 웹앱", "date": "2027-10-07", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["AI 코딩 도구 사용법", "요구사항을 기능 단위로 지시", "생성된 화면 미리보기 및 수정 요청"], "practice": "할 일 목록 또는 퀴즈 앱 예제 제작", "outcome": "첫 번째 미니 웹앱", "materials": [], "videos": []},
      {"week": 46, "title": "HTML·CSS·JavaScript 기초", "tag": "", "stage": "바이브코딩과 AI 웹앱", "date": "2027-10-14", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["HTML 구조·CSS 꾸미기·JS 동작 이해", "버튼과 입력창·목록 분석", "AI에게 코드 설명 요청"], "practice": "기존 예제의 제목·색상·버튼 동작 변경", "outcome": "기초 웹페이지", "materials": [], "videos": []},
      {"week": 47, "title": "웹앱 화면과 사용자 경험", "tag": "", "stage": "바이브코딩과 AI 웹앱", "date": "2027-10-21", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["주요 이용자 및 사용 흐름 정의", "큰 글씨·높은 대비·단순 메뉴", "화면 2~3개 구성"], "practice": "개인 웹앱 화면 설계 및 생성", "outcome": "웹앱 UI 시안", "materials": [], "videos": []},
      {"week": 48, "title": "웹앱 핵심 기능 구현", "tag": "", "stage": "바이브코딩과 AI 웹앱", "date": "2027-10-28", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["입력·검색·필터·저장 기능", "오류 메시지 작성", "AI 코드 생성 후 기능별 수동 테스트"], "practice": "핵심 기능 2개 이상 구현", "outcome": "동작하는 웹앱 초안", "materials": [], "videos": []},
      {"week": 49, "title": "웹앱 데이터·파일 활용", "tag": "", "stage": "바이브코딩과 AI 웹앱", "date": "2027-11-04", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["브라우저 로컬 저장 개념", "사진·CSV 가져오기 예제", "개인정보 수집 최소화와 데이터 백업"], "practice": "기록 저장과 다시 불러오기 구현", "outcome": "데이터 기능 추가본", "materials": [], "videos": []},
      {"week": 50, "title": "웹앱 테스트와 배포", "tag": "", "stage": "바이브코딩과 AI 웹앱", "date": "2027-11-11", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["기능 테스트표 작성", "모바일 화면·접근성 검토", "정적 웹앱 배포와 공유 주소 확인"], "practice": "공개 가능 데이터만 사용해 웹앱 배포", "outcome": "배포 URL·테스트표", "materials": [], "videos": []},
      {"week": 51, "title": "최종 프로젝트 통합·보완", "tag": "", "stage": "바이브코딩과 AI 웹앱", "date": "2027-11-18", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["웹앱 안내문·사용 매뉴얼 작성", "콘텐츠와 자동화 연계 가능성 검토", "발표 시나리오 연습"], "practice": "웹앱 오류 수정과 발표자료 완성", "outcome": "최종 포트폴리오", "materials": [], "videos": []},
      {"week": 52, "title": "최종 성과발표회와 수료", "tag": "발표", "stage": "바이브코딩과 AI 웹앱", "date": "2027-11-25", "time": "19:00 – 21:00", "location": "온라인 실시간", "topics": ["개인 웹앱 5분 시연", "연간 학습 성과 자기평가", "후속 동아리·봉사·창업 아이디어 공유"], "practice": "최종 웹앱 발표와 동료 평가", "outcome": "최종 발표·수료", "materials": [], "videos": []},
      /* WEEKS:END */
    ],

    /* 달력 일정 — 수업 외에 달력에 표시할 일정(휴강·보강·특강·행사 등).
     *   관리자 로그인 뒤 달력에서 날짜를 눌러 추가·수정·삭제할 수 있습니다. */
    events: [
      { date: "2026-12-01", title: "연간 교육 운영 시작", type: "운영", text: "계획 시작일입니다. 첫 정규 수업은 12월 3일(목)입니다." },
      { date: "2027-09-16", title: "추석 연휴 (휴강·보강 검토)", type: "확인", text: "42주차 수업일이 추석 연휴와 겹칩니다. 휴강 여부와 보강일을 정해 공지합니다." },
    ],

    // 최종 프로젝트 예시
    projectsTitle: "1년 동안 쌓는 결과물",
    projectsLead: "주차마다 결과물을 하나씩 만들어 연말에는 개인 AI 포트폴리오가 완성됩니다.",
    projects: [
      "개인 AI 학습 목표서",
      "AI 도구 비교표",
      "AI 기초 안내서",
      "개인 프롬프트 라이브러리",
      "AI 지식 노트",
      "개인 회고록 초고",
      "AI 이미지 작품집",
      "홍보 카드뉴스",
      "10장 발표자료",
      "AI 창작음악",
      "완성 영상 1편",
      "AI 콘텐츠 포트폴리오",
      "완성 프로젝트",
      "작동하는 자동화 워크플로",
      "배포된 개인 웹앱",
      "최종 포트폴리오",
    ],
    projectsNote: "실습은 개인정보를 가리거나 공개·가상 자료를 이용하며, 유료 AI 도구 사용 여부는 사전에 안내합니다.",
  },

  /* ── 포트폴리오: 우수 과제물 (메뉴에서는 빠져 있습니다. 과제물이 쌓이면 nav에 다시 넣으세요) ── */
  portfolio: {
    id: "portfolio",
    eyebrow: "Portfolio",
    title: "수강생 작품",
    lead: "수강생들이 AI로 직접 만든 작품을 소개합니다.",
    buttonLabel: "작품 보기",
    emptyText: "아직 등록된 작품이 없습니다. 첫 결과물이 나오면 이곳에 소개합니다.",
    items: [],
  },

  /* ── 수강 안내: 운영 방식 · 실습 도구 · 수업 구성 · 참고자료 · 준비물 ── */
  guide: {
    id: "guide",
    eyebrow: "Guide",
    title: "수강 안내",
    lead: "모든 수업은 온라인 실시간 화상회의로 진행합니다. 강의 · 시연 · 실습을 한 회 안에서 함께 하고, 매주 결과물을 만들어 공유합니다.",

    toolsTitle: "주요 실습 도구",
    /* 도구 카드를 누르면 detail 정보가 팝업으로 열립니다.
     *   summary 소개 · url 바로가기 · links 여러 링크 · pricing 요금제(plan/price/note)
     *   uses 활용 방법 · weeks 관련 주차 · tip 주의할 점 · pricingUrl 공식 요금 페이지
     *   요금은 2026년 10월 기준 참고값입니다. */
    pricingNote: "※ 요금은 2026년 10월 기준 참고값이며 환율·국가·프로모션에 따라 달라질 수 있습니다. 결제 전 공식 요금 페이지를 꼭 확인하세요.",
    tools: [
      {
        icon: "💬", name: "ChatGPT", category: "생성형 AI",
        text: "질문·요약·문서 작성 등 AI 기초 실습의 중심 도구입니다.",
        detail: {
          maker: "OpenAI",
          url: "https://chatgpt.com",
          pricingUrl: "https://chatgpt.com/pricing",
          summary: "가장 널리 쓰이는 대화형 AI입니다. 말하듯 질문하면 답을 주고, 글쓰기·요약·번역·사진 설명·음성 대화까지 한곳에서 할 수 있습니다.",
          pricing: [
            { plan: "무료", price: "$0", note: "기본 모델, 사용량 제한, 광고가 표시될 수 있음" },
            { plan: "Go", price: "월 $8", note: "더 넉넉한 사용량 (국가별 가격 다름)" },
            { plan: "Plus", price: "월 $20", note: "최신 모델, 심층 리서치, 이미지·영상 생성 확대" },
            { plan: "Pro", price: "월 $100 / $200", note: "전문가용 대용량 사용" },
          ],
          uses: ["궁금한 것을 말하듯 묻고 후속 질문으로 다듬기", "공지문·이메일·인사말 초안 작성 후 고치기", "PDF·사진을 올려 요약하거나 설명 듣기", "스마트폰 앱에서 음성으로 대화하기"],
          weeks: "2·3·4·6주차",
          tip: "개인정보(주민번호·계좌·병원 기록)는 입력하지 마세요. 수업은 무료 플랜으로 진행합니다.",
        },
      },
      {
        icon: "🔮", name: "Gemini", category: "생성형 AI",
        text: "다른 AI와 답변을 비교하고 검색 기반 질문에 활용합니다.",
        detail: {
          maker: "Google",
          url: "https://gemini.google.com",
          pricingUrl: "https://one.google.com/about/google-ai-plans/",
          summary: "구글의 대화형 AI입니다. 구글 검색·지도·유튜브·지메일과 연결되어 최신 정보를 찾거나 구글 서비스와 함께 쓰기 좋습니다.",
          pricing: [
            { plan: "무료", price: "0원", note: "구글 계정만 있으면 사용" },
            { plan: "Google AI Plus", price: "월 약 11,000원", note: "사용량 확대, 저장공간 추가" },
            { plan: "Google AI Pro", price: "월 약 29,000원", note: "고급 모델·심층 리서치, NotebookLM 한도 확대" },
          ],
          uses: ["같은 질문을 ChatGPT와 비교해 보기", "유튜브 영상·웹 페이지 내용 요약", "구글 지도와 함께 여행 계획 세우기", "사진을 찍어 사물·글자 설명 듣기"],
          weeks: "2·5·7주차",
          tip: "답변의 출처 링크를 눌러 사실 여부를 꼭 확인하세요.",
        },
      },
      {
        icon: "✨", name: "Claude", category: "생성형 AI",
        text: "긴 글 작성·문서 분석·코드 생성에 활용합니다.",
        detail: {
          maker: "Anthropic",
          url: "https://claude.ai",
          pricingUrl: "https://claude.com/pricing",
          summary: "자연스럽고 정돈된 글쓰기와 긴 문서 분석에 강한 대화형 AI입니다. 웹페이지·문서를 바로 만들어 보여 주는 기능도 있어 바이브코딩 실습에 활용합니다.",
          pricing: [
            { plan: "무료", price: "$0", note: "하루 사용량 제한" },
            { plan: "Pro", price: "월 $20 (연간 결제 시 월 $17)", note: "사용량 확대, Claude Code 포함" },
            { plan: "Max", price: "월 $100 / $200", note: "대용량 사용" },
          ],
          uses: ["회고록·소개글처럼 긴 글을 함께 쓰고 다듬기", "긴 PDF·보고서를 올려 핵심 정리", "말로 설명해 간단한 웹페이지·퀴즈 만들기", "표·체크리스트 형태로 정리 요청"],
          weeks: "2·11·17·45주차",
          tip: "결과물은 그대로 쓰지 말고 사실과 표현을 한 번 더 확인하세요.",
        },
      },
      {
        icon: "📒", name: "NotebookLM", category: "지식 정리",
        text: "자료를 올려 출처 기반으로 질문하고 학습 노트를 만듭니다.",
        detail: {
          maker: "Google",
          url: "https://notebooklm.google.com",
          pricingUrl: "https://one.google.com/about/google-ai-plans/",
          summary: "내가 올린 자료(PDF·웹페이지·유튜브·메모)만 근거로 답하는 AI 노트입니다. 답변마다 출처가 표시되고, 자료를 대화형 오디오(팟캐스트)로 만들어 주기도 합니다.",
          pricing: [
            { plan: "무료", price: "0원", note: "구글 계정으로 사용, 하루 사용량 제한" },
            { plan: "Google AI 구독", price: "Plus·Pro 요금제에 포함", note: "노트·자료 수, 오디오 개요 횟수 확대" },
          ],
          uses: ["공공기관 자료·PDF 여러 개를 올려 비교 질문", "강의 자료로 요약 노트·퀴즈 만들기", "자료를 오디오 개요로 만들어 산책하며 듣기", "책·기사 내용을 출처와 함께 정리"],
          weeks: "15주차",
          tip: "저작권이 있는 자료는 개인 학습용으로만 쓰고 결과물을 공개하지 마세요.",
        },
      },
      {
        icon: "📌", name: "Pinterest", category: "디자인 참고",
        text: "카드뉴스·포스터·발표자료에 쓸 템플릿과 디자인 아이디어 이미지를 찾아 내려받습니다.",
        detail: {
          maker: "Pinterest",
          url: "https://kr.pinterest.com/",
          summary: "전 세계 사람들이 올린 이미지를 주제별로 모아 보는 시각 검색 서비스입니다. '카드뉴스 템플릿', '포스터 디자인'처럼 검색하면 참고할 디자인을 한눈에 볼 수 있고, 마음에 드는 이미지는 보드에 저장하거나 내려받을 수 있습니다.",
          pricing: [
            { plan: "개인 계정", price: "무료", note: "검색·저장(핀)·보드 만들기·이미지 내려받기 모두 무료 (광고 표시)" },
            { plan: "비즈니스 계정", price: "무료", note: "홍보·통계 기능 (광고 집행은 별도 비용)" },
          ],
          uses: ["'카드뉴스 템플릿', '발표자료 디자인'으로 검색해 참고 이미지 찾기", "마음에 드는 이미지를 내 보드에 저장(핀)해 모아 두기", "이미지를 눌러 오른쪽 위 ⋯ → '이미지 다운로드'로 내려받기", "모은 이미지를 AI에게 보여 주고 비슷한 분위기로 만들어 달라고 하기"],
          weeks: "18·19·20·47주차",
          tip: "Pinterest의 이미지는 대부분 다른 사람의 저작물입니다. 내려받은 이미지는 개인 학습·참고용으로만 쓰고, 그대로 홍보물이나 웹앱에 올리지 마세요.",
        },
      },
      {
        icon: "🎵", name: "음성·음악·영상 AI", category: "콘텐츠 제작",
        text: "내레이션, 노래, 짧은 영상을 AI로 만들어 봅니다.",
        detail: {
          maker: "ElevenLabs · Suno 등",
          url: "https://elevenlabs.io",
          pricingUrl: "https://elevenlabs.io/pricing",
          links: [{ label: "ElevenLabs (음성)", url: "https://elevenlabs.io" }, { label: "Suno (음악)", url: "https://suno.com" }],
          summary: "글을 자연스러운 목소리로 읽어 주는 음성 AI(ElevenLabs), 가사와 분위기만 정하면 노래를 만들어 주는 음악 AI(Suno) 등을 체험합니다. 만든 음성·음악은 Canva에서 영상으로 엮습니다.",
          pricing: [
            { plan: "ElevenLabs 무료", price: "$0", note: "월 10,000 크레딧, 상업적 이용 불가" },
            { plan: "ElevenLabs Starter", price: "월 약 $5~6", note: "상업적 이용 가능" },
            { plan: "Suno 무료", price: "$0", note: "생성 횟수 제한, 상업적 이용 불가" },
            { plan: "Suno Pro", price: "월 약 $8~10", note: "생성량 확대, 상업적 이용 가능" },
          ],
          uses: ["1분 자기소개 내레이션 만들기", "가족·여행 주제로 짧은 노래 만들기", "이미지·음성·음악을 묶어 1분 영상 완성", "자막 넣고 유튜브·블로그에 올리기"],
          weeks: "21·22·23·24주차",
          tip: "다른 사람의 목소리를 흉내 내거나 허락 없이 쓰면 안 됩니다. 무료 플랜 결과물은 개인 용도로만 쓰세요.",
        },
      },
      {
        icon: "📊", name: "Google Sheets", category: "데이터",
        text: "입력값을 자동으로 기록하는 표를 만듭니다.",
        detail: {
          maker: "Google",
          url: "https://sheets.google.com",
          pricingUrl: "https://workspace.google.com/pricing",
          summary: "인터넷에서 쓰는 무료 엑셀입니다. 여러 사람이 함께 고칠 수 있고, 구글 설문지 응답을 자동으로 모으거나 Gemini에게 표 정리를 부탁할 수 있습니다.",
          pricing: [
            { plan: "개인", price: "0원", note: "구글 계정으로 무료 (저장공간 15GB 공유)" },
            { plan: "Google Workspace", price: "1인당 월 유료", note: "회사·단체용 (개인 수업에는 필요 없음)" },
          ],
          uses: ["독서·운동·지출 기록표 만들기", "구글 설문지 응답을 표로 자동 정리", "AI에게 함수·정리 방법 물어보기", "표를 공유해 동호회 명단 함께 관리"],
          weeks: "40주차",
          tip: "링크 공유를 켜면 주소를 아는 누구나 볼 수 있으니 개인정보가 담긴 표는 조심하세요.",
        },
      },
      {
        icon: "⌨️", name: "AI 코딩 도구", category: "바이브코딩",
        text: "말로 지시해 나만의 웹앱을 만들고 배포합니다.",
        detail: {
          maker: "Claude · Google AI Studio 등",
          url: "https://claude.ai",
          pricingUrl: "https://claude.com/pricing",
          links: [{ label: "Claude", url: "https://claude.ai" }, { label: "Google AI Studio", url: "https://aistudio.google.com" }],
          summary: "코드를 몰라도 원하는 화면과 기능을 말로 설명하면 AI가 웹앱을 만들어 줍니다(바이브코딩). 결과를 바로 미리 보고, 고칠 점을 말로 요청하며 완성합니다.",
          pricing: [
            { plan: "Claude 무료", price: "$0", note: "간단한 웹페이지·퀴즈 만들기 체험" },
            { plan: "Claude Pro", price: "월 $20", note: "Claude Code로 본격적인 웹앱 제작" },
            { plan: "Google AI Studio", price: "무료", note: "구글 계정으로 웹앱 생성·미리보기" },
          ],
          uses: ["할 일 목록·퀴즈 같은 미니 웹앱 만들기", "제목·색·버튼을 말로 바꿔 달라고 하기", "큰 글씨·쉬운 메뉴로 보기 편한 화면 설계", "완성한 웹앱을 인터넷 주소로 공유"],
          weeks: "44~52주차",
          tip: "처음에는 작은 기능 하나부터 시작하세요. 개인정보를 받는 웹앱은 만들지 않습니다.",
        },
      },
      {
        icon: "🐙", name: "GitHub", category: "코드 저장·관리",
        text: "만든 웹앱 파일을 인터넷 저장소에 보관하고 버전을 관리합니다.",
        detail: {
          maker: "GitHub (Microsoft)",
          url: "https://github.com",
          pricingUrl: "https://github.com/pricing",
          summary: "웹앱을 이루는 파일(HTML·CSS·JavaScript 등)을 인터넷 저장소에 올려 보관하는 서비스입니다. 고칠 때마다 기록이 남아 예전 상태로 되돌릴 수 있고, Vercel·Railway와 연결하면 저장소에 올리는 것만으로 웹앱이 자동 배포됩니다.",
          pricing: [
            { plan: "Free", price: "$0", note: "개인 사용 무료, 공개·비공개 저장소 모두 가능" },
            { plan: "Pro", price: "월 약 $4", note: "고급 기능 추가 (수업에는 필요 없음)" },
            { plan: "Copilot", price: "무료 / Pro 월 $10~", note: "AI 코딩 도우미는 별도 요금 (수업에는 필요 없음)" },
          ],
          uses: ["무료 계정 만들고 첫 저장소(Repository) 만들기", "AI가 만든 웹앱 파일을 저장소에 올리기 (Add file → Upload files)", "파일을 고친 기록(Commit) 보고 예전으로 되돌리기", "Vercel·Railway와 연결해 자동 배포하기"],
          weeks: "50~52주차",
          tip: "비밀번호·API 키 같은 비밀 정보는 저장소에 올리지 마세요. 공개 저장소는 누구나 볼 수 있습니다.",
        },
      },
      {
        icon: "▲", name: "Vercel", category: "웹 배포",
        text: "GitHub에 올린 웹앱을 클릭 몇 번으로 인터넷 주소에 공개합니다.",
        detail: {
          maker: "Vercel",
          url: "https://vercel.com",
          pricingUrl: "https://vercel.com/pricing",
          summary: "GitHub 저장소를 연결하면 웹앱을 인터넷에 공개하고 '○○.vercel.app' 주소를 만들어 주는 배포 서비스입니다. 파일을 고쳐 GitHub에 올리면 사이트도 자동으로 새로 고쳐집니다. 화면만 있는 정적 웹앱을 올리기에 가장 쉽습니다.",
          pricing: [
            { plan: "Hobby", price: "$0", note: "개인·비상업용 무료, 월 데이터 전송량 한도 있음" },
            { plan: "Pro", price: "1인당 월 $20", note: "상업용 사이트, 사용량 크레딧 포함 (초과분 별도 과금)" },
            { plan: "Enterprise", price: "별도 문의", note: "기업용" },
          ],
          uses: ["GitHub 계정으로 로그인하기", "저장소를 골라 'Deploy' 눌러 웹앱 공개하기", "만들어진 주소를 가족·동료에게 공유하기", "파일을 고쳐 다시 올리면 자동으로 반영되는지 확인"],
          weeks: "50주차",
          tip: "수업 실습은 무료(Hobby)로 충분합니다. 수익을 내는 사이트는 유료(Pro)를 써야 합니다.",
        },
      },
      {
        icon: "🚆", name: "Railway", category: "웹 배포·서버",
        text: "데이터 저장이 필요한 웹앱을 서버와 데이터베이스까지 함께 배포합니다.",
        detail: {
          maker: "Railway",
          url: "https://railway.com",
          pricingUrl: "https://railway.com/pricing",
          summary: "서버 프로그램과 데이터베이스가 필요한 웹앱을 올려 운영하는 클라우드 서비스입니다. GitHub 저장소를 연결하면 자동으로 배포되고, 이 사이트처럼 신청서·출석 기록을 저장하는 웹앱(PostgreSQL)을 운영할 수 있습니다.",
          pricing: [
            { plan: "체험(Trial)", price: "$5 크레딧 1회", note: "가입 시 한 번 제공, 약 30일 후 만료" },
            { plan: "Free", price: "$0", note: "매월 $1 크레딧 (아주 작은 실습용)" },
            { plan: "Hobby", price: "월 $5", note: "$5 사용량 포함, 초과분은 사용한 만큼 과금" },
            { plan: "Pro", price: "월 $20", note: "팀·운영용, 사용량 별도" },
          ],
          uses: ["GitHub 계정으로 로그인하고 저장소 연결하기", "웹앱과 데이터베이스(PostgreSQL)를 함께 만들기", "환경 변수(비밀번호 등)를 안전하게 설정하기", "사용량·요금 화면에서 비용 확인하기"],
          weeks: "50주차",
          tip: "사용한 만큼 요금이 나오는 방식이라, 실습이 끝나면 쓰지 않는 서비스는 꼭 정지·삭제하세요. 결제 카드를 등록하기 전에 요금 화면을 확인하세요.",
        },
      },
    ],

    // 막대 그래프 상자 — percent에는 숫자, gradingUnit은 숫자 뒤에 붙는 단위, gradingTotal은 막대 100%에 해당하는 값
    gradingTitle: "한 회 수업 구성 (120분)",
    gradingUnit: "분",
    gradingTotal: 30,
    grading: [
      { label: "화상회의 입장·안내", percent: 10 },
      { label: "개념 설명", percent: 15 },
      { label: "시연", percent: 15 },
      { label: "실습 ①", percent: 25 },
      { label: "휴식", percent: 10 },
      { label: "실습 ②", percent: 30 },
      { label: "발표·정리", percent: 15 },
    ],
    gradingNote: "※ 수업 내용과 진행 상황에 따라 시간 배분은 조정될 수 있습니다.",

    // 교재와 참고자료 (href를 비우면 링크 없이 이름만 표시)
    referencesTitle: "참고자료",
    referencesLead: "강사가 제공하는 강의자료와 각 도구의 공식 도움말을 중심으로 진행합니다.",
    references: [
      { label: "ChatGPT", href: "https://chatgpt.com" },
      { label: "Google Gemini", href: "https://gemini.google.com" },
      { label: "Claude", href: "https://claude.ai" },
      { label: "NotebookLM", href: "https://notebooklm.google.com" },
      { label: "Canva", href: "https://www.canva.com" },
      { label: "Make", href: "https://www.make.com" },
      { label: "n8n", href: "https://n8n.io" },
    ],
    referencesNote: "AI 도구는 화면과 요금이 자주 바뀌므로 수업 때 최신 화면으로 다시 안내합니다.",

    prepTitle: "수강 준비물",
    prep: [
      { icon: "💻", title: "컴퓨터 또는 태블릿", text: "카메라·마이크가 되는 기기 (실습은 컴퓨터를 권장합니다)" },
      { icon: "📱", title: "스마트폰", text: "스마트폰 AI·음성 대화 실습(4주차)에 사용" },
      { icon: "📧", title: "Google 계정", text: "Gemini, NotebookLM, Google Sheets 실습에 사용" },
      { icon: "🎧", title: "이어폰·헤드셋", text: "화상회의 음성을 또렷하게 듣고 말하기 위해 권장" },
      { icon: "🖥️", title: "보조모니터 (권장)", text: "화상회의와 실습 화면을 나눠 보면 훨씬 편합니다. 아래 추천 제품을 참고하세요." },
    ],

    /* 보조모니터 안내 — 수강 준비물 아래에 표시됩니다. 통째로 지우면 보이지 않습니다.
     *   image : 제품 사진 경로 (images/monitors 폴더)   specs : 주요 사양   point : 이런 분께 추천
     *   url   : '가격 비교 보기' 버튼이 열 주소 */
    monitorsTitle: "효율적인 수강을 위한 *보조모니터*",
    monitorsLead: "한 화면에는 화상회의(강사 시연), 다른 화면에는 AI 도구를 띄워 두면 시연을 보면서 바로 따라 할 수 있습니다. 노트북에 케이블 하나로 연결하는 포터블 모니터를 추천합니다.",
    monitorTips: [
      { icon: "👀", title: "왼쪽 화면", text: "화상회의 — 강사 화면 공유와 시연 보기" },
      { icon: "✍️", title: "오른쪽 화면", text: "실습 — ChatGPT·Canva 등 AI 도구 창" },
      { icon: "🔌", title: "연결", text: "USB-C 케이블 하나로 영상·전원 연결 (노트북 USB-C가 영상 출력을 지원해야 합니다. 없으면 mini HDMI 사용)" },
      { icon: "⚙️", title: "설정", text: "윈도우: Win + P → '확장' 선택 · 맥: 시스템 설정 → 디스플레이에서 배치" },
    ],
    monitors: [
      {
        brand: "LG전자",
        name: "그램 +view 3세대 17MT70",
        badge: "17형 · WQXGA",
        image: "images/monitors/lg-17mt70.jpg",
        specs: ["17형 2560×1600 (16:10)", "IPS 무광 · 350nit", "USB-C · mini HDMI", "약 780g", "화면 자동 회전 (LG Switch 앱)"],
        point: "가장 큰 17형 화면 — 글씨가 크게 보여 화상회의와 문서 보기에 편합니다.",
        url: "https://search.shopping.naver.com/search/all?query=LG%20%EA%B7%B8%EB%9E%A8%EB%B7%B0%2017MT70",
      },
      {
        brand: "제우스랩",
        name: "P16K",
        badge: "16형 · 2.5K",
        image: "images/monitors/zeuslap-p16k.jpg",
        specs: ["16형 2560×1600 (16:10)", "IPS · 500nit · sRGB 100%", "최대 144Hz (USB-C) / 120Hz (HDMI)", "USB-C 2개 · mini HDMI", "30W PD 충전기 포함"],
        point: "밝고 선명한 2.5K 화면을 부담 없는 가격에 — 처음 보조모니터를 쓰는 분께.",
        url: "https://search.shopping.naver.com/search/all?query=%EC%A0%9C%EC%9A%B0%EC%8A%A4%EB%9E%A9%20P16K",
      },
      {
        brand: "제우스랩",
        name: "P16KT (터치)",
        badge: "16형 · 2.5K · 터치",
        image: "images/monitors/zeuslap-p16kt.jpg",
        specs: ["16형 2560×1600 (16:10)", "IPS · 500nit · sRGB 100%", "터치스크린 (윈도우 노트북에서 동작)", "USB-C 2개 · mini HDMI", "30W PD 충전기 포함"],
        point: "화면을 손가락으로 눌러 조작 — 마우스가 아직 익숙하지 않은 분께 편리합니다.",
        url: "https://search.shopping.naver.com/search/all?query=%EC%A0%9C%EC%9A%B0%EC%8A%A4%EB%9E%A9%20P16KT",
      },
    ],
    monitorsNote: "※ 특정 제품 구매를 강요하지 않습니다. 가지고 계신 모니터나 TV(HDMI)로도 충분합니다. 사양·가격은 판매처와 시기에 따라 다를 수 있으니 구매 전 확인해 주세요.",
  },

  /* ── 참여하기: 투표 + 수강 신청서 ──────────────────────── */
  participate: {
    id: "participate",
    eyebrow: "Join",
    title: "참여하기",
    lead: "여러분의 의견을 들려주고, 수강 신청서를 작성해 주세요.",

    /* 설문 — 여러 개를 둘 수 있습니다. 관리자 로그인 뒤 '참여하기'에서 추가·수정·삭제하세요. */
    pollsTitle: "설문",
    polls: [
      {
        id: "p1",
        title: "가장 기대되는 교육 영역은?",
        help: "하나를 골라 주세요. 다시 눌러 바꿀 수 있습니다.",
        options: [
          "AI 기초와 디지털 활용",
          "프롬프트 전략과 AI 도구",
          "AI 콘텐츠 제작 (글·이미지·음성·영상)",
          "AI 프로젝트 기획·제작",
          "AI 자동화와 에이전트",
          "바이브코딩과 AI 웹앱",
        ],
        createdAt: "2026-10-30T12:00",
        open: true,
      },
    ],

    apply: {
      title: "수강 신청서",
      notice: "지금은 시범 운영 중이라 신청서가 이 기기(브라우저)에만 저장됩니다.",
      endpoint: "",
      submitLabel: "신청서 제출하기",
      successTitle: "신청서가 접수되었습니다 🌿",
      successText: "작성해 주셔서 고맙습니다. 확인 후 연락드리겠습니다.",
      fields: [
        { name: "name", label: "이름", type: "text", required: true, placeholder: "홍길동" },
        { name: "studentId", label: "생년월일", type: "text", required: true, placeholder: "예) 19600315 (숫자 8자리)", pattern: "^\\d{8}$", patternMsg: "생년월일은 숫자 8자리로 입력해 주세요." },
        { name: "department", label: "이전 직업·경력", type: "text", required: false, placeholder: "예) 교사 30년, 공무원" },
        { name: "grade", label: "AI 사용 경험", type: "select", required: true, options: ["처음입니다", "몇 번 써 봤습니다", "자주 사용합니다"] },
        { name: "email", label: "이메일", type: "email", required: true, placeholder: "name@example.com" },
        { name: "phone", label: "연락처", type: "tel", required: true, placeholder: "010-0000-0000", pattern: "^[0-9-]{9,13}$", patternMsg: "연락처는 숫자와 -만 입력해 주세요." },
        { name: "motivation", label: "신청 동기", type: "textarea", required: true, placeholder: "AI로 해 보고 싶은 일을 자유롭게 적어 주세요." },
        { name: "agree", label: "개인정보 수집·이용에 동의합니다.", type: "checkbox", required: true },
      ],
    },
  },

  /* ── 내 강의실: 로그인 · 출석 · 과제 제출 ──────────────── */
  classroom: {
    id: "classroom",
    eyebrow: "My Class",
    title: "내 강의실",
    lead: "로그인하면 출석을 체크하고 과제를 제출할 수 있습니다.",
    notice: "지금은 시범 운영 중이라 출석·제출 기록이 이 기기(브라우저)에만 저장되고, 파일은 서버로 전송되지 않습니다.",
    // 수강 코드와 수강생 명단은 원문 대신 지문(해시)으로 저장됩니다.
    // 직접 고치지 말고 관리자 화면(오른쪽 위 자물쇠) → '수강생 명단'에서 바꾸세요.
    accessCodeHash: "",
    rosterHashes: [],
    testMode: true,
    fileAccept: ".pdf,.docx,.hwp,.hwpx,.pptx,.xlsx,.zip,.jpg,.png,.mp3,.mp4",
    fileMaxMB: 20,
  },

  /* ── 안내 팝업 ───────────────────────────────────────── */
  popup: {
    delaySeconds: 2,
    hideTodayLabel: "오늘 하루 보지 않기",
  },
  popups: [
    {
      id: "pop1",
      enabled: true,
      badge: "수강생 모집",
      title: "모산 with AI 1기 수강생을 모집합니다",
      text: "2026년 12월 3일(목) 개강 · 매주 목요일 저녁 7시 · 온라인 화상회의 · 권장 인원 20~25명. 정원이 차면 조기 마감될 수 있습니다.",
      buttonLabel: "수강 신청하러 가기",
      buttonHref: "#apply",
    },
  ],

  /* ── 첫 방문 환영 효과 ─────────────────────────────────── */
  welcome: {
    fireworks: true,
    message: "모산 with AI에 오신 것을 환영합니다 🎉",
  },

  /* ── FAQ ───────────────────────────────────────────────── */
  faq: {
    id: "faq",
    eyebrow: "FAQ",
    title: "자주 묻는 질문",
    lead: "수강을 고민하시는 분들이 자주 묻는 내용을 모았습니다.",
    items: [
      { q: "컴퓨터를 잘 못 다뤄도 들을 수 있나요?", a: "네. 1~8주차는 AI 기초와 디지털 활용으로, 회원가입·로그인부터 스마트폰 음성 대화까지 천천히 함께 해 봅니다. 수업마다 시연을 보고 바로 따라 하는 실습 시간이 있습니다." },
      { q: "수업은 어떻게 진행되나요?", a: "매주 목요일 19:00~21:00에 화상회의로 진행합니다. 입장·안내 10분, 개념 15분, 시연 15분, 실습 25분, 휴식 10분, 실습 30분, 발표·정리 15분으로 구성됩니다." },
      { q: "명절이나 공휴일에는 어떻게 하나요?", a: "명절·공휴일과 강사·학습자 일정에 따라 휴강할 수 있으며, 이때는 보강일을 따로 정해 공지하고 교육 일정 캘린더에 표시합니다." },
      { q: "유료 AI 도구를 사야 하나요?", a: "기본 실습은 무료 기능으로 진행합니다. 유료 기능이 필요한 경우에는 미리 안내하고, 무료로 할 수 있는 대안을 함께 알려 드립니다." },
      { q: "개인정보가 걱정됩니다.", a: "실습은 개인정보를 가리거나 공개·가상 자료를 이용합니다. 3주차와 8주차에 AI에 개인정보를 입력할 때의 원칙과 체크리스트를 따로 배웁니다." },
      { q: "수료하면 무엇이 남나요?", a: "주차마다 만든 결과물(프롬프트 라이브러리, 회고록, 카드뉴스, 영상, 자동화 워크플로, 개인 웹앱 등)이 개인 포트폴리오로 남고, 52주차 최종 성과발표회에서 수료합니다." },
    ],
  },

  /* ── 교수자 (페이지 끝 푸터) ───────────────────────────── */
  instructor: {
    id: "instructor",
    eyebrow: "Instructor",
    title: "교수자",
    name: "박정원 교수",
    nameSub: "朴正元 | Park Jeong Weon",
    role: "모산학술연구소 · 한국외국어대학교 중국학대학 중국언어문화학부",
    roleSub: "Division of Chinese Language, Literature and Culture",
    photo: "images/professor.jpg",
    bio: [
      "AI 기반 교육 콘텐츠 제작, 데이터 분석, 데이터 시각화 및 웹플랫폼 개발을 연구·교육하며, 생성형 AI를 활용한 교육 및 연구 방법론을 개발하고 있습니다.",
      "주요 관심 분야는 생성형 AI, AI 데이터 분석, 디지털인문학, 웹플랫폼, AI 자동화, 데이터 큐레이션 및 시각화이며, ChatGPT, Claude Code, Codex, Google AI Studio, Make, n8n 등 다양한 AI 도구를 활용한 교육 및 연구를 진행합니다.",
    ],
    career: [
      "국가교육위원회 AI특별위원회 위원",
      "교육부 AI인재양성추진단",
    ],
    worksTitle: "저술",
    works: [
      { title: "데이터 큐레이팅", href: "https://www.upaper.net/auraweon/1144732" },
      { title: "AI 데이터 자동화: N8N(기업업무편)", href: "https://auraweon.upaper.kr/content/1204663" },
      { title: "AI 데이터 자동화: OPAL", href: "https://auraweon.upaper.kr/content/1204909" },
      { title: "AI 미디어 큐레이션", href: "https://auraweon.upaper.kr/content/1220294" },
      { title: "AI 대시보드", href: "https://auraweon.upaper.kr/content/1221085" },
      { title: "AI 데이터 자동화", href: "https://auraweon.upaper.kr/content/1224709" },
    ],
    contacts: [
      { icon: "📱", label: "mobile", value: "010-9131-6127", href: "tel:010-9131-6127" },
      { icon: "✉️", label: "email", value: "park9626@hanmail.net", href: "mailto:park9626@hanmail.net" },
      { icon: "🌐", label: "APP", value: "www.kletter.kr", href: "https://www.kletter.kr" },
      { icon: "🌐", label: "APP", value: "www.kteacher.kr", href: "https://www.kteacher.kr" },
    ],
  },

  /* ── 관리자 ────────────────────────────────────────────── */
  admin: {
    // 비어 있으면 자물쇠를 처음 누를 때 비밀번호를 만들게 됩니다. 배포 전에 꼭 만들어 두세요.
    passwordHash: "",
  },

  /* ── 맨 아래 한 줄 ─────────────────────────────────────── */
  footer: {
    copyright: "Copyright 2026 ⓒ 모산 with AI · 모산학술연구소, All Rights Reserved.",
    links: [
      { icon: "✉️", label: "AI 뉴스레터 빌더", text: "www.kletter.kr", href: "https://www.kletter.kr" },
      { icon: "🎓", label: "AI 강의콘텐츠 빌더", text: "www.kteacher.kr", href: "https://www.kteacher.kr" },
    ],
  },
};
