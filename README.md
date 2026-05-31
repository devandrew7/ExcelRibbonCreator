# 🎬 Excel Custom Ribbon Creator

> **WYSIWYG 실시간 시뮬레이터를 기반으로 엑셀 Fluent UI 표준 규격의 Custom UI XML 및 VBA 매크로 소스 코드를 생성하고 최종 파일 패키징(.xlsm)까지 지원하는 초고속 리본 개발 플랫폼입니다.**

---

## ✨ 핵심 하이라이트 (Core Highlights)

*   **실시간 WYSIWYG 리본 시뮬레이터**: 마우스 드래그 앤 드롭 및 4방향 D-Pad 정밀 컨트롤 이동을 통해 엑셀 프로그램과 동일한 레이아웃을 웹에서 실시간 설계 및 수정합니다.
*   **자주 사용하는 VBA 매크로 라이브러리**: 중앙 하단에 탑재된 VBA 프리셋(시트 보호, PDF 내보내기, 열 너비 맞춤, 중복 제거, 시트 취합 등)을 즉석에서 검토하고 복사해 속성창에서 컨트롤에 연결할 수 있습니다.
*   **지능형 3단 세로 적층 Stacking**: 소형 컨트롤들을 만나는 즉시 최대 3개 단위의 세로 열로 자동 묶어 정렬하고, 엑셀 표준 박스 스키마인 `<box boxStyle="vertical">`으로 완벽히 컴파일합니다.
*   **로컬 M365 고해상도 SVG 아이콘 스캐너**: 파워셸(PowerShell) 연동 스크립트를 통해 로컬 폴더 경로의 모든 SVG 자산을 동적 2단 Sub-tab 형태로 자동 렌더링하고 이미지 태그로 무설정 맵핑합니다.
*   **토글-라디오 배타적 선택 제어**: 토글 버튼들에 그룹명을 지정하여 동일 그룹 내 오직 단 하나만 활성화 상태(`checked = true`)를 가지는 라디오 버튼 조작과 물리적 쑥 들어간 눌림 효과(`.pressed` 스타일)를 가동합니다.
*   **브라우저 내장 패키징 빌더**: 다운로드 단추 클릭 시 서버 연동 없이 즉석에서 XML과 매크로 정보가 패키징된 순정 엑셀 통합 문서(`.xlsm`)를 인코딩 내보내기합니다.

---

## 🛠️ 개발 기술 스택 (Tech Stack)

*   **Frontend**: Vanilla HTML5, JavaScript ES6+, Vanilla CSS3 (HSL Tailored Theme, Glassmorphism, Micro-animations)
*   **스캐너 연동**: PowerShell Script (`generate_icons_map.ps1`)
*   **패키징**: JSZip Module Integration

---

## 🚀 빠른 시작 (Quick Start)

본 프로젝트는 100% 클라이언트 사이드 정적 웹 애플리케이션으로, 복잡한 빌드 도구나 Node.js 설치가 전혀 필요 없습니다.

1.  이 저장소를 클론하거나 다운로드합니다:
    ```bash
    git clone https://github.com/devandrew7/ExcelRibbonCreator.git
    ```
2.  프로젝트 루트 디렉토리에 위치한 `index.html` 파일을 크롬, 엣지, 사파리 등 현대적인 웹 브라우저에서 더블 클릭하여 실행합니다.

---

## 📂 로컬 M365 고해상도 아이콘 폴더 자동 동기화

로컬 SVG 아이콘 파일들을 프로젝트의 동적 아이콘 팝업 선택창에 동기화 반영하는 방법입니다.

1.  프로젝트 디렉토리의 `icons/2024-microsoft-365-content-icons/` 경로 하위에 자유롭게 아이콘 폴더 및 SVG 파일들을 정리합니다.
    *   *예: `icons/2024-microsoft-365-content-icons/Microsoft Blue/48x48 Light Blue Icon/Folder.svg`*
2.  **`generate_icons_map.ps1` 파워셸 파일**을 마우스 우클릭한 후 **[PowerShell로 실행]**을 선택하여 자동 스캔을 마칩니다.
3.  웹 브라우저를 새로고침하면, **스캔 완료된 계층 폴더명 그대로 2단계 동적 하위 탭(Level 1, Level 2 sub-tabs)이 피커 모달에 100% 조립 반영**됩니다.

---

## 📖 문서 정보 링크 (Documentation Links)

본 프로젝트의 상세 아키텍처 정보와 단계별 실무 가이드는 아래 전용 문서를 참고하십시오:

*   📘 **[개발 기획서 (PLANNING.md)](PLANNING.md)**: 데이터 모델 설계 구조, 글로벌 상태 객체(`ribbonState`), 시스템 아키텍처, 렌더링 파이프라인 및 향후 개발 로드맵 명세서.
*   📗 **[사용 설명서 (USER_GUIDE.md)](USER_GUIDE.md)**: 드래그 앤 드롭 컨트롤 배치, D-Pad 4방향 단추 조작법, 토글-라디오 단일선택 그룹화, 다운로드한 엑셀 파일에 VBA 매크로 소스코드를 안전하게 매핑 활성화하는 엑셀 실적용 튜토리얼 가이드.

---

## 📄 라이선스 (License)

본 프로젝트는 오픈소스로 배포되며 자유로운 사용, 수정 및 기여를 환영합니다.
구체적인 라이선스 유형은 별도 협의에 따릅니다.
