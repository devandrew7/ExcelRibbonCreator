/**
 * Excel Custom Ribbon Creator - MS Office Native imageMso dataset & SVGs
 * Defines standard Excel-supported built-in icons with outline vector rendering.
 */

// --- 1. Excel Native imageMso Groups and Korean Labels ---
const OFFICE_MSO_ICONS = {
    "파일 (File)": {
        "FileNew": { label: "새로 만들기", desc: "새 문서나 파일을 생성합니다." },
        "FileOpen": { label: "열기", desc: "기존 문서를 불러옵니다." },
        "FileSave": { label: "저장", desc: "현재 통합 문서를 저장합니다." },
        "FileSaveAs": { label: "다른 이름으로 저장", desc: "현재 통합 문서를 다른 이름으로 복사해 저장합니다." },
        "FilePrint": { label: "인쇄", desc: "현재 활성 시트를 출력합니다." },
        "PrintPreview": { label: "인쇄 미리보기", desc: "출력 상태를 사전에 조회합니다." },
        "Close": { label: "닫기", desc: "현재 활성 통합 문서를 닫습니다." },
        "Exit": { label: "끝내기", desc: "엑셀 응용 프로그램을 종료합니다." }
    },
    "편집 (Edit)": {
        "Copy": { label: "복사", desc: "선택 영역을 클립보드에 복사합니다." },
        "Cut": { label: "잘라내기", desc: "선택 영역을 클립보드로 이동합니다." },
        "Paste": { label: "붙여넣기", desc: "클립보드 내용을 선택 영역에 덮어씁니다." },
        "Undo": { label: "실행 취소", desc: "마지막 작업을 뒤로 돌립니다." },
        "Redo": { label: "다시 실행", desc: "마지막으로 취소한 작업을 재실행합니다." },
        "Delete": { label: "삭제", desc: "선택한 표나 셀 내용을 지웁니다." },
        "FindDialog": { label: "찾기", desc: "통합 문서 내에서 텍스트를 탐색합니다." },
        "ReplaceDialog": { label: "바꾸기", desc: "찾은 텍스트를 다른 텍스트로 대체합니다." }
    },
    "서식 (Format)": {
        "Bold": { label: "굵게", desc: "글꼴을 굵게 지정합니다." },
        "Italic": { label: "기울임꼴", desc: "글꼴을 기울입니다." },
        "Underline": { label: "밑줄", desc: "글꼴 아래에 밑줄을 긋습니다." },
        "FontColor": { label: "글꼴 색", desc: "글자 색상을 지정합니다." },
        "FontDialog": { label: "글꼴 서식", desc: "전체 글꼴 속성 대화상자를 호출합니다." },
        "FormatPainter": { label: "서식 복사", desc: "선택한 항목의 서식을 복사해 다른 곳에 적용합니다." },
        "ClearFormats": { label: "서식 지우기", desc: "선택 영역의 모든 서식을 초기화합니다." }
    },
    "정렬 (Align)": {
        "AlignLeft": { label: "왼쪽 맞춤", desc: "텍스트를 왼쪽에 정렬합니다." },
        "AlignCenter": { label: "가운데 맞춤", desc: "텍스트를 가운데에 정렬합니다." },
        "AlignRight": { label: "오른쪽 맞춤", desc: "텍스트를 오른쪽에 정렬합니다." },
        "AlignJustify": { label: "양끝 맞춤", desc: "텍스트를 양끝 가장자리에 맞춥니다." },
        "MergeCells": { label: "셀 병합", desc: "선택된 여러 셀을 하나의 셀로 합칩니다." },
        "UnmergeCells": { label: "셀 분할", desc: "병합되어 있는 셀들을 원래대로 나눕니다." },
        "WrapText": { label: "텍스트 줄 바꿈", desc: "셀 너비를 넘어가는 긴 텍스트를 여러 줄로 표시합니다." }
    },
    "데이터 (Data)": {
        "SortAscending": { label: "오름차순 정렬", desc: "데이터를 오름차순으로 정렬합니다." },
        "SortDescending": { label: "내림차순 정렬", desc: "데이터를 내림차순으로 정렬합니다." },
        "Filter": { label: "필터", desc: "표 데이터의 필터를 설정합니다." },
        "AutoFilter": { label: "자동 필터", desc: "데이터 열에 자동 필터를 활성화합니다." },
        "CalculateNow": { label: "지금 계산", desc: "전체 통합 문서의 수식을 강제로 계산합니다." },
        "Refresh": { label: "새로고침", desc: "연결된 외부 데이터 소스를 갱신합니다." },
        "Group": { label: "그룹", desc: "선택한 셀 범위를 그룹으로 묶어 요약할 수 있게 합니다." },
        "Ungroup": { label: "그룹 해제", desc: "설정되어 있는 셀 그룹을 해제합니다." }
    },
    "삽입 (Insert)": {
        "TableInsert": { label: "표 삽입", desc: "데이터 영역에 엑셀 표준 표를 삽입합니다." },
        "ChartInsert": { label: "차트 삽입", desc: "시각적 분석용 표면 차트를 개설합니다." },
        "PivotTableInsert": { label: "피벗 테이블", desc: "복잡한 분석 데이터를 피벗으로 취합합니다." },
        "HyperlinkInsert": { label: "하이퍼링크", desc: "웹 주소나 문서 내 링크를 지정합니다." },
        "PictureInsert": { label: "그림 삽입", desc: "PC나 클라우드에서 그림을 문서에 삽입합니다." },
        "ShapeInsert": { label: "도형 삽입", desc: "원, 직사각형, 화살표 등 도형을 삽입합니다." },
        "TextBoxInsert": { label: "텍스트 상자", desc: "원하는 위치에 자유롭게 글을 쓰는 텍스트 상자를 삽입합니다." }
    },
    "검토 및 보안 (Review)": {
        "SheetProtect": { label: "시트 보호", desc: "시트 내용 수정을 잠급니다." },
        "SheetUnprotect": { label: "시트 보호 해제", desc: "시트 잠금을 해제하고 수정을 허용합니다." },
        "WorkbookProtect": { label: "통합 문서 보호", desc: "시트 추가/삭제 등의 문서 구조 변경을 제한합니다." },
        "WorkbookUnprotect": { label: "통합 문서 보호 해제", desc: "통합 문서 보호 구조 잠금을 해제합니다." },
        "ReviewCheckSpelling": { label: "맞춤법 검사", desc: "문장의 맞춤법 오류를 스캔합니다." },
        "Mail": { label: "이메일 전송", desc: "메일 시스템을 열어 첨부/발송합니다." },
        "CommentInsert": { label: "새 메모", desc: "선택한 셀에 새 메모를 추가합니다." },
        "CommentDelete": { label: "메모 삭제", desc: "선택한 셀의 메모를 지웁니다." }
    },
    "개발 도구 (Developer)": {
        "MacroPlay": { label: "매크로 실행", desc: "VBA 매크로 목록을 열고 실행합니다." },
        "VisualBasic": { label: "VBA 에디터", desc: "Alt+F11 매크로 편집창을 띄웁니다." },
        "DeveloperTools": { label: "개발 도구", desc: "개발 기능 제어 속성을 다룹니다." },
        "MacroRecord": { label: "매크로 기록", desc: "작업 과정을 VBA 코드로 기록합니다." },
        "AddInManager": { label: "추가 기능", desc: "엑셀 추가 기능 목록을 관리합니다." }
    },
    "기타 (Misc)": {
        "HappyFace": { label: "웃는 얼굴", desc: "RibbonX 기본 시각 테스트 아이콘입니다." },
        "Help": { label: "도움말", desc: "도움말 센터 및 상세 안내를 호출합니다." },
        "Settings": { label: "설정", desc: "엑셀 및 추가기능 옵션을 조절합니다." },
        "Options": { label: "옵션", desc: "엑셀 프로그램 전체 환경 설정을 엽니다." },
        "Properties": { label: "속성", desc: "통합 문서의 메타데이터 및 속성 정보를 엽니다." }
    }
};

// --- 2. Unified Vector SVG Path Registry for Mso Icons ---
function getIconSvg(iconKey, size = 18) {
    if (!iconKey) return "";

    const SVG_PATHS = {
        // --- 파일 ---
        "FileNew": `<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" fill="none" stroke="currentColor" stroke-width="2"/><polyline points="14 2 14 8 20 8" fill="none" stroke="currentColor" stroke-width="2"/><line x1="12" y1="18" x2="12" y2="12" stroke="currentColor" stroke-width="2"/><line x1="9" y1="15" x2="15" y2="15" stroke="currentColor" stroke-width="2"/>`,
        "FileOpen": `<path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/><path d="M2 10h20" fill="none" stroke="currentColor" stroke-width="1.5"/>`,
        "FileSave": `<path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z" fill="none" stroke="currentColor" stroke-width="2"/><polyline points="17 21 17 13 7 13 7 21" fill="none" stroke="currentColor" stroke-width="2"/><polyline points="7 3 7 8 15 8" fill="none" stroke="currentColor" stroke-width="2"/>`,
        "FileSaveAs": `<path d="M17 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l4 4v7" fill="none" stroke="currentColor" stroke-width="2"/><polyline points="15 21 15 13 7 13 7 21" fill="none" stroke="currentColor" stroke-width="2"/><path d="M17.5 13.5l3.5-3.5 1.5 1.5-3.5 3.5z" fill="currentColor"/>`,
        "FilePrint": `<path d="M6 9V2h12v7M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2" fill="none" stroke="currentColor" stroke-width="2"/><rect x="6" y="14" width="12" height="8" fill="none" stroke="currentColor" stroke-width="2"/>`,
        "PrintPreview": `<path d="M4 18h10M4 14h6" stroke="currentColor" stroke-width="2"/><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" fill="none" stroke="currentColor" stroke-width="2"/><circle cx="15" cy="13" r="4" fill="none" stroke="currentColor" stroke-width="1.5"/><line x1="18" y1="16" x2="21" y2="19" stroke="currentColor" stroke-width="2"/>`,
        "Close": `<line x1="18" y1="6" x2="6" y2="18" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><line x1="6" y1="6" x2="18" y2="18" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>`,
        "Exit": `<path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/><polyline points="16 17 21 12 16 7" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/><line x1="21" y1="12" x2="9" y2="12" stroke="currentColor" stroke-width="2"/>`,
        
        // --- 편집 ---
        "Copy": `<rect x="9" y="9" width="13" height="13" rx="2" ry="2" fill="none" stroke="currentColor" stroke-width="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" fill="none" stroke="currentColor" stroke-width="2"/>`,
        "Cut": `<circle cx="6" cy="6" r="3" fill="none" stroke="currentColor" stroke-width="2"/><circle cx="6" cy="18" r="3" fill="none" stroke="currentColor" stroke-width="2"/><line x1="9.8" y1="8.5" x2="20" y2="17" stroke="currentColor" stroke-width="2"/><line x1="9.8" y1="15.5" x2="20" y2="7" stroke="currentColor" stroke-width="2"/>`,
        "Paste": `<path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2" fill="none" stroke="currentColor" stroke-width="2"/><rect x="8" y="2" width="8" height="4" rx="1" fill="none" stroke="currentColor" stroke-width="2"/>`,
        "Undo": `<path d="M3 7v6h6" fill="none" stroke="currentColor" stroke-width="2"/><path d="M21 17a9 9 0 0 0-9-9 9 9 0 0 0-6 2.3L3 13" fill="none" stroke="currentColor" stroke-width="2"/>`,
        "Redo": `<path d="M21 7v6h-6" fill="none" stroke="currentColor" stroke-width="2"/><path d="M3 17a9 9 0 0 1 9-9 9 9 0 0 1 6 2.3l3 2.7" fill="none" stroke="currentColor" stroke-width="2"/>`,
        "Delete": `<polyline points="3 6 5 6 21 6" fill="none" stroke="currentColor" stroke-width="2"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" fill="none" stroke="currentColor" stroke-width="2"/>`,
        "FindDialog": `<circle cx="11" cy="11" r="8" fill="none" stroke="currentColor" stroke-width="2"/><line x1="21" y1="21" x2="16.65" y2="16.65" stroke="currentColor" stroke-width="2"/>`,
        "ReplaceDialog": `<circle cx="9" cy="9" r="6" fill="none" stroke="currentColor" stroke-width="2"/><line x1="18" y1="18" x2="13.2" y2="13.2" stroke="currentColor" stroke-width="2"/><path d="M16 5h5M18 3l3 2-3 2" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>`,
        
        // --- 서식 ---
        "Bold": `<path d="M6 4h6a4 4 0 0 1 4 4 4 4 0 0 1-4 4H6z" fill="none" stroke="currentColor" stroke-width="3.5"/><path d="M6 12h7a4 4 0 0 1 4 4 4 4 0 0 1-4 4H6z" fill="none" stroke="currentColor" stroke-width="3.5"/><line x1="6" y1="4" x2="6" y2="20" stroke="currentColor" stroke-width="3.5"/>`,
        "Italic": `<line x1="19" y1="4" x2="10" y2="20" stroke="currentColor" stroke-width="3.5" stroke-linecap="round"/><line x1="14" y1="4" x2="20" y2="4" stroke="currentColor" stroke-width="2.5"/><line x1="8" y1="20" x2="14" y2="20" stroke="currentColor" stroke-width="2.5"/>`,
        "Underline": `<path d="M6 3v7a6 6 0 0 0 12 0V3" fill="none" stroke="currentColor" stroke-width="2.5"/><line x1="4" y1="19" x2="20" y2="19" stroke="currentColor" stroke-width="3"/>`,
        "FontColor": `<path d="M5 16l6-12h2l6 12" fill="none" stroke="currentColor" stroke-width="2"/><line x1="7" y1="12" x2="17" y2="12" stroke="currentColor" stroke-width="1.5"/><rect x="3" y="19" width="18" height="4" fill="currentColor"/>`,
        "FontDialog": `<path d="M5 16l6-12h2l6 12" fill="none" stroke="currentColor" stroke-width="2"/><circle cx="18" cy="18" r="4" fill="none" stroke="currentColor" stroke-width="1.5"/><line x1="7" y1="12" x2="17" y2="12" stroke="currentColor" stroke-width="1.5"/>`,
        "FormatPainter": `<path d="M18 14V6a4 4 0 0 0-8 0v8" fill="none" stroke="currentColor" stroke-width="2"/><rect x="6" y="14" width="8" height="6" rx="1" fill="none" stroke="currentColor" stroke-width="2"/><line x1="10" y1="2" x2="10" y2="5" stroke="currentColor" stroke-width="2"/>`,
        "ClearFormats": `<path d="M18 14V6a4 4 0 0 0-8 0v8" fill="none" stroke="currentColor" stroke-width="2"/><line x1="6" y1="20" x2="14" y2="14" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><line x1="14" y1="20" x2="6" y2="14" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>`,
        
        // --- 정렬 ---
        "AlignLeft": `<line x1="17" y1="10" x2="3" y2="10" stroke="currentColor" stroke-width="2"/><line x1="21" y1="6" x2="3" y2="6" stroke="currentColor" stroke-width="2"/><line x1="21" y1="14" x2="3" y2="14" stroke="currentColor" stroke-width="2"/><line x1="15" y1="18" x2="3" y2="18" stroke="currentColor" stroke-width="2"/>`,
        "AlignCenter": `<line x1="18" y1="10" x2="6" y2="10" stroke="currentColor" stroke-width="2"/><line x1="21" y1="6" x2="3" y2="6" stroke="currentColor" stroke-width="2"/><line x1="21" y1="14" x2="3" y2="14" stroke="currentColor" stroke-width="2"/><line x1="16" y1="18" x2="8" y2="18" stroke="currentColor" stroke-width="2"/>`,
        "AlignRight": `<line x1="21" y1="10" x2="7" y2="10" stroke="currentColor" stroke-width="2"/><line x1="21" y1="6" x2="3" y2="6" stroke="currentColor" stroke-width="2"/><line x1="21" y1="14" x2="3" y2="14" stroke="currentColor" stroke-width="2"/><line x1="21" y1="18" x2="9" y2="18" stroke="currentColor" stroke-width="2"/>`,
        "AlignJustify": `<line x1="21" y1="10" x2="3" y2="10" stroke="currentColor" stroke-width="2"/><line x1="21" y1="6" x2="3" y2="6" stroke="currentColor" stroke-width="2"/><line x1="21" y1="14" x2="3" y2="14" stroke="currentColor" stroke-width="2"/><line x1="21" y1="18" x2="3" y2="18" stroke="currentColor" stroke-width="2"/>`,
        "MergeCells": `<rect x="3" y="3" width="18" height="18" rx="2" fill="none" stroke="currentColor" stroke-width="2"/><polyline points="8 12 5 12 5 16" fill="none" stroke="currentColor" stroke-width="1.5"/><polyline points="16 12 19 12 19 8" fill="none" stroke="currentColor" stroke-width="1.5"/><line x1="3" y1="12" x2="21" y2="12" stroke="currentColor" stroke-width="1.5" stroke-dasharray="2,2"/>`,
        "UnmergeCells": `<rect x="3" y="3" width="18" height="18" rx="2" fill="none" stroke="currentColor" stroke-width="2"/><line x1="3" y1="12" x2="21" y2="12" stroke="currentColor" stroke-width="1.5"/><line x1="12" y1="3" x2="12" y2="21" stroke="currentColor" stroke-width="1.5"/>`,
        "WrapText": `<path d="M3 6h18M3 12h12a4 4 0 0 1 0 8H7" fill="none" stroke="currentColor" stroke-width="2"/><polyline points="10 17 7 20 10 23" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>`,
        
        // --- 데이터 ---
        "SortAscending": `<line x1="4" y1="6" x2="14" y2="6" stroke="currentColor" stroke-width="2"/><line x1="4" y1="12" x2="11" y2="12" stroke="currentColor" stroke-width="2"/><line x1="4" y1="18" x2="8" y2="18" stroke="currentColor" stroke-width="2"/><polyline points="16 14 19 11 22 14" fill="none" stroke="currentColor" stroke-width="2"/><line x1="19" y1="11" x2="19" y2="21" stroke="currentColor" stroke-width="2"/>`,
        "SortDescending": `<line x1="4" y1="6" x2="14" y2="6" stroke="currentColor" stroke-width="2"/><line x1="4" y1="12" x2="11" y2="12" stroke="currentColor" stroke-width="2"/><line x1="4" y1="18" x2="8" y2="18" stroke="currentColor" stroke-width="2"/><polyline points="16 18 19 21 22 18" fill="none" stroke="currentColor" stroke-width="2"/><line x1="19" y1="11" x2="19" y2="21" stroke="currentColor" stroke-width="2"/>`,
        "Filter": `<polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/>`,
        "AutoFilter": `<polygon points="18 3 2 3 9 11 9 17 12 19 12 11 18 3" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/><polygon points="18 12 21 15 24 12 18 12" fill="currentColor"/>`,
        "CalculateNow": `<rect x="4" y="2" width="16" height="20" rx="2" fill="none" stroke="currentColor" stroke-width="2"/><line x1="8" y1="6" x2="16" y2="6" stroke="currentColor" stroke-width="1.5"/><circle cx="8" cy="11" r="1" fill="currentColor"/><circle cx="12" cy="11" r="1" fill="currentColor"/><circle cx="16" cy="11" r="1" fill="currentColor"/><circle cx="8" cy="15" r="1" fill="currentColor"/><circle cx="12" cy="15" r="1" fill="currentColor"/><circle cx="16" cy="15" r="1" fill="currentColor"/><circle cx="8" cy="19" r="1" fill="currentColor"/><circle cx="12" cy="19" r="1" fill="currentColor"/><circle cx="16" cy="19" r="1" fill="currentColor"/>`,
        "Refresh": `<path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l.73-1.19" fill="none" stroke="currentColor" stroke-width="2"/>`,
        "Group": `<rect x="3" y="3" width="7" height="7" rx="1" fill="none" stroke="currentColor" stroke-width="2"/><rect x="14" y="3" width="7" height="7" rx="1" fill="none" stroke="currentColor" stroke-width="2"/><rect x="3" y="14" width="7" height="7" rx="1" fill="none" stroke="currentColor" stroke-width="2"/><rect x="14" y="14" width="7" height="7" rx="1" fill="none" stroke="currentColor" stroke-width="2"/><line x1="10" y1="6" x2="14" y2="6" stroke="currentColor" stroke-width="1.5" stroke-dasharray="2,2"/><line x1="6" y1="10" x2="6" y2="14" stroke="currentColor" stroke-width="1.5" stroke-dasharray="2,2"/>`,
        "Ungroup": `<rect x="3" y="3" width="7" height="7" rx="1" fill="none" stroke="currentColor" stroke-width="2"/><rect x="14" y="3" width="7" height="7" rx="1" fill="none" stroke="currentColor" stroke-width="2"/><rect x="3" y="14" width="7" height="7" rx="1" fill="none" stroke="currentColor" stroke-width="2"/><rect x="14" y="14" width="7" height="7" rx="1" fill="none" stroke="currentColor" stroke-width="2"/><line x1="10" y1="6" x2="14" y2="6" stroke="currentColor" stroke-width="2"/><line x1="6" y1="10" x2="6" y2="14" stroke="currentColor" stroke-width="2"/>`,
        
        // --- 삽입 ---
        "TableInsert": `<rect x="3" y="3" width="18" height="18" rx="2" fill="none" stroke="currentColor" stroke-width="2"/><line x1="3" y1="9" x2="21" y2="9" stroke="currentColor" stroke-width="1.5"/><line x1="3" y1="15" x2="21" y2="15" stroke="currentColor" stroke-width="1.5"/><line x1="9" y1="3" x2="9" y2="21" stroke="currentColor" stroke-width="1.5"/><line x1="15" y1="3" x2="15" y2="21" stroke="currentColor" stroke-width="1.5"/>`,
        "ChartInsert": `<rect x="3" y="3" width="18" height="18" rx="2" fill="none" stroke="currentColor" stroke-width="2"/><line x1="7" y1="17" x2="7" y2="12" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"/><line x1="12" y1="17" x2="12" y2="7" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"/><line x1="17" y1="17" x2="17" y2="10" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"/>`,
        "PivotTableInsert": `<rect x="3" y="3" width="18" height="18" rx="2" fill="none" stroke="currentColor" stroke-width="2"/><line x1="9" y1="3" x2="9" y2="21" stroke="currentColor" stroke-width="1.5"/><line x1="3" y1="9" x2="21" y2="9" stroke="currentColor" stroke-width="1.5"/><circle cx="15" cy="15" r="2.5" fill="none" stroke="currentColor" stroke-width="1.5"/>`,
        "HyperlinkInsert": `<path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" fill="none" stroke="currentColor" stroke-width="2"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71" fill="none" stroke="currentColor" stroke-width="2"/>`,
        "PictureInsert": `<rect x="3" y="3" width="18" height="18" rx="2" fill="none" stroke="currentColor" stroke-width="2"/><circle cx="8.5" cy="8.5" r="1.5" fill="currentColor"/><polyline points="21 15 16 10 5 21" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/>`,
        "ShapeInsert": `<rect x="3" y="3" width="7" height="7" rx="1" fill="none" stroke="currentColor" stroke-width="2"/><circle cx="17.5" cy="6.5" r="3.5" fill="none" stroke="currentColor" stroke-width="2"/><polygon points="12 21 6 12 18 12" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/>`,
        "TextBoxInsert": `<rect x="3" y="3" width="18" height="18" rx="2" fill="none" stroke="currentColor" stroke-width="2" stroke-dasharray="3,3"/><text x="12" y="16" font-family="serif" font-size="14" text-anchor="middle" fill="currentColor" font-weight="bold">A</text>`,
        
        // --- 검토 및 보안 ---
        "SheetProtect": `<rect x="3" y="3" width="18" height="18" rx="2" fill="none" stroke="currentColor" stroke-width="2"/><line x1="3" y1="9" x2="21" y2="9" stroke="currentColor" stroke-width="1.5"/><rect x="9" y="13" width="6" height="5" rx="1" fill="currentColor"/><path d="M10 13v-2a2 2 0 0 1 4 0v2" fill="none" stroke="currentColor" stroke-width="1.2"/>`,
        "SheetUnprotect": `<rect x="3" y="3" width="18" height="18" rx="2" fill="none" stroke="currentColor" stroke-width="2"/><line x1="3" y1="9" x2="21" y2="9" stroke="currentColor" stroke-width="1.5"/><rect x="9" y="13" width="6" height="5" rx="1" fill="currentColor"/><path d="M10 11v-2a2 2 0 0 1 4 0" fill="none" stroke="currentColor" stroke-width="1.2"/>`,
        "WorkbookProtect": `<rect x="3" y="3" width="18" height="18" rx="2" fill="none" stroke="currentColor" stroke-width="2"/><line x1="3" y1="7" x2="21" y2="7" stroke="currentColor" stroke-width="1.5"/><rect x="9" y="11" width="6" height="5" rx="1" fill="currentColor"/><path d="M10 11v-2a2 2 0 0 1 4 0v2" fill="none" stroke="currentColor" stroke-width="1.2"/>`,
        "WorkbookUnprotect": `<rect x="3" y="3" width="18" height="18" rx="2" fill="none" stroke="currentColor" stroke-width="2"/><line x1="3" y1="7" x2="21" y2="7" stroke="currentColor" stroke-width="1.5"/><rect x="9" y="11" width="6" height="5" rx="1" fill="currentColor"/><path d="M10 9v-2a2 2 0 0 1 4 0" fill="none" stroke="currentColor" stroke-width="1.2"/>`,
        "ReviewCheckSpelling": `<path d="M9 11l3 3 7-7" fill="none" stroke="currentColor" stroke-width="2"/><text x="6" y="16" font-family="sans-serif" font-size="12" font-weight="bold" fill="currentColor">ABC</text>`,
        "Mail": `<path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z" fill="none" stroke="currentColor" stroke-width="2"/><polyline points="22,6 12,13 2,6" stroke="currentColor" stroke-width="2" fill="none"/>`,
        "CommentInsert": `<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z" fill="none" stroke="currentColor" stroke-width="2"/><line x1="12" y1="7" x2="12" y2="13" stroke="currentColor" stroke-width="2"/><line x1="9" y1="10" x2="15" y2="10" stroke="currentColor" stroke-width="2"/>`,
        "CommentDelete": `<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z" fill="none" stroke="currentColor" stroke-width="2"/><line x1="9" y1="7" x2="15" y2="13" stroke="currentColor" stroke-width="2"/><line x1="15" y1="7" x2="9" y2="13" stroke="currentColor" stroke-width="2"/>`,
        
        // --- 개발 도구 ---
        "MacroPlay": `<polygon points="6 3 20 12 6 21 6 3" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/><circle cx="12" cy="12" r="9" fill="none" stroke="currentColor" stroke-width="1.5" stroke-dasharray="2,2"/>`,
        "VisualBasic": `<rect x="3" y="3" width="18" height="18" rx="2" fill="none" stroke="currentColor" stroke-width="2"/><text x="7" y="11" font-family="Consolas, monospace" font-size="7" font-weight="bold" fill="currentColor">Sub</text><text x="7" y="18" font-family="Consolas, monospace" font-size="7" font-weight="bold" fill="currentColor">VBA</text>`,
        "DeveloperTools": `<path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94z" fill="none" stroke="currentColor" stroke-width="2"/>`,
        "MacroRecord": `<circle cx="12" cy="12" r="10" fill="none" stroke="currentColor" stroke-width="2"/><circle cx="12" cy="12" r="4" fill="red"/>`,
        "AddInManager": `<rect x="3" y="3" width="18" height="18" rx="2" fill="none" stroke="currentColor" stroke-width="2"/><path d="M12 7v10M7 12h10" stroke="currentColor" stroke-width="2"/><path d="M9 9l6 6m0-6l-6 6" stroke="currentColor" stroke-width="1.2"/>`,
        
        // --- 기타 ---
        "HappyFace": `<circle cx="12" cy="12" r="10" fill="none" stroke="currentColor" stroke-width="2"/><path d="M8 14s1.5 2 4 2 4-2 4-2" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><circle cx="9" cy="9" r="1" fill="currentColor"/><circle cx="15" cy="9" r="1" fill="currentColor"/>`,
        "Help": `<circle cx="12" cy="12" r="10" fill="none" stroke="currentColor" stroke-width="2"/><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><line x1="12" y1="17" x2="12.01" y2="17" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"/>`,
        "Settings": `<circle cx="12" cy="12" r="3" fill="none" stroke="currentColor" stroke-width="2"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/>`,
        "Options": `<circle cx="12" cy="12" r="3" fill="none" stroke="currentColor" stroke-width="2"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/>`,
        "Properties": `<rect x="4" y="3" width="16" height="18" rx="2" fill="none" stroke="currentColor" stroke-width="2"/><line x1="7" y1="7" x2="17" y2="7" stroke="currentColor" stroke-width="1.5"/><line x1="7" y1="11" x2="17" y2="11" stroke="currentColor" stroke-width="1.5"/><line x1="7" y1="15" x2="13" y2="15" stroke="currentColor" stroke-width="1.5"/>`
    };

    const markup = SVG_PATHS[iconKey] || `<circle cx="12" cy="12" r="10" fill="none" stroke="currentColor" stroke-width="2"></circle><text x="12" y="15" font-family="sans-serif" font-size="8" text-anchor="middle" fill="currentColor" font-weight="bold">Mso</text>`;
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 24 24" class="mso-svg-icon" style="display:inline-block; vertical-align:middle;">${markup}</svg>`;
}

// Export modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { OFFICE_MSO_ICONS, getIconSvg };
} else {
    window.OFFICE_MSO_ICONS = OFFICE_MSO_ICONS;
    window.getIconSvg = getIconSvg;
}
