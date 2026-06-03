/**
 * Excel Custom Ribbon Creator - Extended Core Application Logic Script (v4)
 * Implements M365 local directory dynamic folder mapping, 5-group emoji sub-tabs,
 * unified image/emoji renderer, tab reordering, and light theme state binding.
 */

// --- 1. Global State Management ---
let ribbonState = {
    tabs: []
};

let selectedElement = null; // { type: 'tab'|'group'|'control'|'separator', path: [tabIdx, grpIdx, ctrlIdx] }

// --- Predefined VBA Actions Presets ---
const VBA_PRESETS = [
    {
        id: "protect_sheet",
        label: "🔒 시트 보호 및 해제 토글",
        description: "현재 시트를 암호('1234')로 잠그거나 보호 해제하는 토글 동작을 수행합니다.",
        code: `Sub %ACTION_NAME%(control As IRibbonControl)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    On Error GoTo ErrHandler
    If ws.ProtectContents Then
        ws.Unprotect Password:="1234"
        MsgBox "시트 보호가 성공적으로 해제되었습니다.", vbInformation, "엑셀 리본 빌더"
    Else
        ws.Protect Password:="1234", AllowFiltering:=True, AllowFormattingCells:=True
        MsgBox "시트가 암호('1234')로 성공적으로 보호되었습니다.", vbInformation, "엑셀 리본 빌더"
    End If
    Exit Sub
ErrHandler:
    MsgBox "에러가 발생했습니다: " & Err.Description, vbCritical, "오류"
End Sub`
    },
    {
        id: "export_pdf",
        label: "📄 선택 영역 PDF로 내보내기",
        description: "현재 마우스로 드래그 선택한 영역을 즉시 고품질 PDF 문서로 내보내어 저장합니다.",
        code: `Sub %ACTION_NAME%(control As IRibbonControl)
    Dim savePath As Variant
    Dim targetRange As Range
    
    Set targetRange = Selection
    If targetRange Is Nothing Then
        MsgBox "선택 영역이 없습니다. 영역을 드래그한 뒤 실행해 주세요.", vbExclamation, "엑셀 리본 빌더"
        Exit Sub
    End If
    
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="선택영역_출력_" & Format(Now, "yyyymmdd_hhnnss"), _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="PDF 파일로 저장할 위치 지정")
        
    If savePath <> False Then
        On Error GoTo ErrHandler
        targetRange.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=savePath, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=True
        MsgBox "성공적으로 PDF 파일이 생성되었습니다.", vbInformation, "엑셀 리본 빌더"
    End If
    Exit Sub
ErrHandler:
    MsgBox "PDF 생성 중 에러가 발생했습니다: " & Err.Description, vbCritical, "오류"
End Sub`
    },
    {
        id: "autofit_columns",
        label: "↔️ 열 너비 및 행 높이 자동 맞춤",
        description: "현재 활성화된 전체 시트 데이터 영역의 모든 열 너비와 행 높이를 자동으로 적정 크기로 맞춥니다.",
        code: `Sub %ACTION_NAME%(control As IRibbonControl)
    On Error Resume Next
    Application.ScreenUpdating = False
    
    With ActiveSheet.UsedRange
        .Columns.AutoFit
        .Rows.AutoFit
    End With
    
    Application.ScreenUpdating = True
    MsgBox "선택한 시트의 전체 열 너비와 행 높이를 자동 조정했습니다.", vbInformation, "엑셀 리본 빌더"
End Sub`
    },
    {
        id: "remove_duplicates",
        label: "🗑️ 선택 영역 중복 데이터 일괄 제거",
        description: "선택한 표 데이터 영역 내에서 첫 번째 열 기준 중복 값을 제거하고 고유 값만 남깁니다.",
        code: `Sub %ACTION_NAME%(control As IRibbonControl)
    Dim targetRange As Range
    Set targetRange = Selection
    
    If targetRange.Cells.Count <= 1 Then
        MsgBox "중복을 제거할 표 영역을 드래그 선택해 주세요.", vbExclamation, "엑셀 리본 빌더"
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    targetRange.RemoveDuplicates Columns:=1, Header:=xlYes
    MsgBox "선택한 데이터 범위에서 첫 번째 열 기준 중복된 데이터를 지우고 고유 값만 남겼습니다.", vbInformation, "엑셀 리본 빌더"
    Exit Sub
ErrHandler:
    MsgBox "중복값 제거 처리 중 에러: " & Err.Description, vbCritical, "오류"
End Sub`
    },
    {
        id: "merge_sheets",
        label: "📑 여러 워크시트 데이터 일괄 취합",
        description: "현재 통합 문서 내의 모든 시트에 들어 있는 데이터를 하나의 시트로 긁어모아 일괄 취합합니다.",
        code: `Sub %ACTION_NAME%(control As IRibbonControl)
    Dim ws As Worksheet
    Dim mergeWs As Worksheet
    Dim nextRow As Long
    Dim lastRow As Long
    Dim isFirst As Boolean
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    ' "통합취합데이터" 시트 생성
    Set mergeWs = Worksheets.Add(Before:=Worksheets(1))
    mergeWs.Name = "통합취합데이터_" & Format(Now, "hhnnss")
    isFirst = True
    
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> mergeWs.Name Then
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            If lastRow > 1 Then
                If isFirst Then
                    ' 헤더포함 복사
                    ws.Rows("1:" & lastRow).Copy mergeWs.Range("A1")
                    isFirst = False
                Else
                    ' 데이터만 복사
                    nextRow = mergeWs.Cells(mergeWs.Rows.Count, "A").End(xlUp).Row + 1
                    ws.Rows("2:" & lastRow).Copy mergeWs.Range("A" & nextRow)
                End If
            End If
        End If
    Next ws
    
    Application.ScreenUpdating = True
    MsgBox "모든 워크시트의 데이터를 취합 완료했습니다! 새 시트를 확인하세요.", vbInformation, "엑셀 리본 빌더"
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "데이터 취합 중 오류가 발생했습니다: " & Err.Description, vbCritical, "오류"
End Sub`
    }
];
let currentIconTargetControl = null;
let currentIconSourceTab = "mso"; // "mso", "emoji"

// Dynamic M365 Picker Selected states
let m365SelectedLevel1 = ""; // e.g., "Microsoft Blue"
let m365SelectedLevel2 = ""; // e.g., "48x48 Light Blue Icon"

// Emoji Selected state
let emojiSelectedGroup = "smileys"; // "smileys", "nature", "people", "objects", "symbols"

// --- 2. Advanced Multi-Component Preset Demos ---
const DEMO_WORKPLACE = {
    tabs: [
        {
            id: "customTab_Automate",
            label: "업무 자동화 도구",
            visible: true,
            isStandardTab: false,
            idMso: "",
            groups: [
                {
                    id: "group_DataCompile",
                    label: "데이터 전처리",
                    visible: true,
                    controls: [
                        {
                            id: "btn_MergeFiles",
                            type: "button",
                            label: "파일 취합",
                            size: "large",
                            imageMso: "icons/2024-microsoft-365-content-icons/Microsoft Blue/48x48 Light Blue Icon/Folder Open.svg",
                            onAction: "btn_MergeFiles_Click",
                            enabled: true,
                            visible: true
                        },
                        {
                            id: "btn_CleanDuplicates",
                            type: "button",
                            label: "중복값 제거",
                            size: "normal",
                            imageMso: "icons/2024-microsoft-365-content-icons/Microsoft Blue/48x48 Light Blue Icon/Dismiss Circle.svg",
                            onAction: "btn_CleanDuplicates_Click",
                            enabled: true,
                            visible: true
                        },
                        {
                            id: "sep_1",
                            type: "separator"
                        },
                        {
                            id: "chk_BackupFirst",
                            type: "checkbox",
                            label: "자동 백업",
                            onAction: "chk_BackupFirst_Click",
                            enabled: true,
                            visible: true,
                            checked: true
                        },
                        {
                            id: "edit_FilePrefix",
                            type: "editbox",
                            label: "결과 파일 접두사",
                            onAction: "edit_FilePrefix_Change",
                            enabled: true,
                            visible: true,
                            text: "COMP_"
                        }
                    ]
                },
                {
                    id: "group_QuickAutomation",
                    label: "스마트 유틸리티",
                    visible: true,
                    controls: [
                        {
                            id: "btn_SendEmail",
                            type: "togglebutton",
                            label: "부서원 알림",
                            size: "large",
                            imageMso: "icons/2024-microsoft-365-content-icons/Microsoft Blue/48x48 Light Blue Icon/Mail.svg",
                            onAction: "btn_SendEmail_Click",
                            enabled: true,
                            visible: true,
                            checked: false
                        },
                        {
                            id: "combo_SheetColor",
                            type: "combobox",
                            label: "배경 테마색",
                            onAction: "combo_SheetColor_Change",
                            enabled: true,
                            visible: true,
                            text: "초록색",
                            items: [
                                { id: "col_green", label: "초록색" },
                                { id: "col_blue", label: "파란색" }
                            ]
                        },
                        {
                            id: "drop_QuickReport",
                            type: "dropdown",
                            label: "보고서 브릿지",
                            onAction: "drop_QuickReport_Change",
                            enabled: true,
                            visible: true,
                            items: [
                                { id: "sales_rep", label: "매출 요약" },
                                { id: "hr_status", label: "근무 현황" }
                            ]
                        }
                    ]
                }
            ]
        }
    ]
};

const DEMO_HOME_INTRUSION = {
    tabs: [
        {
            id: "standardTabHome",
            label: "홈 (표준 탭 주입 데모)",
            visible: true,
            isStandardTab: true,
            idMso: "TabHome",
            groups: [
                {
                    id: "group_MyOfficeQuick",
                    label: "빠른 도구",
                    visible: true,
                    controls: [
                        {
                            id: "btn_CopilotWeb",
                            type: "button",
                            label: "AI 챗봇 호출",
                            size: "large",
                            imageMso: "💡", // Emoji Demo
                            onAction: "btn_CopilotWeb_Click",
                            enabled: true,
                            visible: true
                        },
                        {
                            id: "split_BackupMenu",
                            type: "splitbutton",
                            label: "보안 작업",
                            size: "large",
                            imageMso: "icons/2024-microsoft-365-content-icons/Microsoft Blue/48x48 Light Blue Icon/Lock Shield.svg", // Local SVG File Demo
                            onAction: "split_BackupMenu_Click",
                            enabled: true,
                            visible: true,
                            controls: [
                                { id: "btn_InnerLock", type: "button", label: "시트 비밀번호 잠금", imageMso: "Lock", onAction: "btn_InnerLock_Click" },
                                { id: "btn_InnerUnlock", type: "button", label: "보호 전체 해제", imageMso: "Unlock", onAction: "btn_InnerUnlock_Click" }
                            ]
                        },
                        {
                            id: "lbl_Version",
                            type: "labelcontrol",
                            label: "보안 v2.4 (인공지능)"
                        }
                    ]
                }
            ]
        }
    ]
};

// Initialize App
document.addEventListener("DOMContentLoaded", () => {
    // Load Workplace preset demo
    loadState(JSON.parse(JSON.stringify(DEMO_WORKPLACE)));
    
    // Bind click / input events
    bindEvents();
    
    // Initialize VBA Preset Library UI
    initVbaLibraryUI();
    
// Setup first dynamic folder keys
    if (typeof ICONS_MAP !== 'undefined') {
        const level1Keys = Object.keys(ICONS_MAP);
        if (level1Keys.length > 0) {
            m365SelectedLevel1 = level1Keys[0];
            if (Array.isArray(ICONS_MAP[m365SelectedLevel1])) {
                m365SelectedLevel2 = "";
            } else {
                const level2Keys = Object.keys(ICONS_MAP[m365SelectedLevel1]);
                if (level2Keys.length > 0) {
                    m365SelectedLevel2 = level2Keys[0];
                } else {
                    m365SelectedLevel2 = "";
                }
            }
        }
    }
    
    // Build Icon Grid inside picker
    buildIconPickerGrid();
});

// --- Predefined VBA Preset Library UI Initialization ---
function initVbaLibraryUI() {
    const container = document.getElementById("vba-preset-buttons-container");
    const viewer = document.getElementById("library-vba-viewer");
    
    if (!container || !viewer) return;
    
    container.innerHTML = "";
    VBA_PRESETS.forEach((preset) => {
        const btn = document.createElement("button");
        btn.className = "btn-secondary vba-preset-btn";
        btn.style.textAlign = "left";
        btn.style.padding = "6px 10px";
        btn.style.fontSize = "11px";
        btn.style.whiteSpace = "nowrap";
        btn.style.overflow = "hidden";
        btn.style.textOverflow = "ellipsis";
        btn.style.width = "100%";
        btn.style.justifyContent = "flex-start";
        btn.innerHTML = preset.label;
        btn.title = preset.description;
        
        btn.addEventListener("click", () => {
            document.querySelectorAll(".vba-preset-btn").forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
            
            // Format generic function name for representation (e.g. ProtectSheet)
            const cleanActionName = preset.id.replace(/_([a-z])/g, (g) => g[1].toUpperCase()) + "_Click";
            const codeToShow = preset.code.replace(/%ACTION_NAME%/g, cleanActionName);
            viewer.value = codeToShow;
        });
        
        container.appendChild(btn);
    });
    
    // Select first preset by default
    if (VBA_PRESETS.length > 0 && container.children.length > 0) {
        container.children[0].click();
    }
}

// --- 3. Event Binders ---
function bindEvents() {
    // Demo presets loader
    document.getElementById("btn-demo-workplace").addEventListener("click", () => {
        loadState(JSON.parse(JSON.stringify(DEMO_WORKPLACE)));
    });
    document.getElementById("btn-demo-intrusion").addEventListener("click", () => {
        loadState(JSON.parse(JSON.stringify(DEMO_HOME_INTRUSION)));
    });
    document.getElementById("btn-clear-all").addEventListener("click", () => {
        loadState({ tabs: [] });
    });

    // Palette triggers (10 controls)
    document.getElementById("add-tab").addEventListener("click", addNewTab);
    document.getElementById("add-group").addEventListener("click", addNewGroup);
    document.getElementById("add-button").addEventListener("click", () => addNewControl('button'));
    document.getElementById("add-togglebutton").addEventListener("click", () => addNewControl('togglebutton'));
    document.getElementById("add-checkbox").addEventListener("click", () => addNewControl('checkbox'));
    document.getElementById("add-editbox").addEventListener("click", () => addNewControl('editbox'));
    document.getElementById("add-combobox").addEventListener("click", () => addNewControl('combobox'));
    document.getElementById("add-dropdown").addEventListener("click", () => addNewControl('dropdown'));
    document.getElementById("add-menu").addEventListener("click", () => addNewControl('menu'));
    document.getElementById("add-labelcontrol").addEventListener("click", () => addNewControl('labelcontrol'));
    document.getElementById("add-separator").addEventListener("click", () => addNewControl('separator'));
    document.getElementById("add-box").addEventListener("click", () => addNewControl('box'));

    // Palette Drag & Drop binders
    const draggablePaletteItems = document.querySelectorAll(".palette-item[draggable='true']");
    draggablePaletteItems.forEach(item => {
        item.addEventListener("dragstart", (e) => {
            const type = item.getAttribute("data-type");
            e.dataTransfer.setData("text/plain", type);
            e.dataTransfer.effectAllowed = "copy";
            item.style.opacity = "0.5";
        });
        item.addEventListener("dragend", () => {
            item.style.opacity = "";
        });
    });

    // Output code panel switches
    const tabButtons = document.querySelectorAll(".tab-btn");
    tabButtons.forEach(btn => {
        btn.addEventListener("click", (e) => {
            tabButtons.forEach(b => b.classList.remove("active"));
            e.target.classList.add("active");
            
            const targetPanelId = e.target.getAttribute("data-tab");
            document.querySelectorAll(".code-panel").forEach(p => p.classList.remove("active"));
            document.getElementById(targetPanelId).classList.add("active");
        });
    });

    // Copy XML / VBA / Preset Library VBA
    document.getElementById("copy-xml").addEventListener("click", () => copyCode("code-xml-box", "copy-xml"));
    document.getElementById("copy-vba").addEventListener("click", () => copyCode("code-vba-box", "copy-vba"));
    
    const copyLibraryBtn = document.getElementById("copy-library-vba");
    if (copyLibraryBtn) {
        copyLibraryBtn.addEventListener("click", () => copyCode("library-vba-viewer", "copy-library-vba"));
    }

    // File packaging export triggers (Safely bind in case button is missing)
    const xlsmBtn = document.getElementById("btn-export-xlsm");
    if (xlsmBtn) xlsmBtn.addEventListener("click", () => exportExcelFile("xlsm"));
    const xlamBtn = document.getElementById("btn-export-xlam");
    if (xlamBtn) xlamBtn.addEventListener("click", () => exportExcelFile("xlam"));

    // Collapsible Notice card
    const warningCard = document.getElementById("warning-callout-card");
    const warningContent = document.getElementById("warning-callout-content");
    const warningArrow = document.getElementById("warning-toggle-arrow");
    if (warningCard && warningContent && warningArrow) {
        warningCard.addEventListener("click", () => {
            const isHidden = warningContent.style.display === "none";
            if (isHidden) {
                warningContent.style.display = "block";
                warningArrow.style.transform = "rotate(180deg)";
            } else {
                warningContent.style.display = "none";
                warningArrow.style.transform = "rotate(0deg)";
            }
        });
    }

    // Selected Property Panel inputs live tracking
    bindPropertyFields();

    // Modal Close
    document.querySelector(".close-modal").addEventListener("click", closeIconPicker);
    document.getElementById("modal-icon-search").addEventListener("input", filterIconsList);

    // Dynamic Icon Source Tabs switch (Excel M365 vs Emoji)
    const pickerSourceBtns = document.querySelectorAll(".picker-source-btn");
    pickerSourceBtns.forEach(btn => {
        btn.addEventListener("click", (e) => {
            pickerSourceBtns.forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
            currentIconSourceTab = btn.getAttribute("data-source");
            
            buildIconPickerGrid();
        });
    });
}

function initializeControlVbaProperties(ctrl) {
    if (ctrl.type === 'separator' || ctrl.type === 'labelcontrol' || ctrl.type === 'box' || ctrl.type === 'buttongroup') return;
    
    if (!ctrl.onAction) {
        if (ctrl.type === 'checkbox') ctrl.onAction = ctrl.id + "_Click";
        else if (ctrl.type === 'editbox') ctrl.onAction = ctrl.id + "_Change";
        else if (['combobox', 'dropdown', 'gallery'].includes(ctrl.type)) ctrl.onAction = ctrl.id + "_Change";
        else ctrl.onAction = ctrl.id + "_Click";
    }

    if (!ctrl.vbaCustomCode) {
        let defaultCode = "";
        if (ctrl.type === 'checkbox') {
            defaultCode = `Sub ${ctrl.onAction}(control As IRibbonControl, pressed As Boolean)\n    MsgBox "체크 상태 변경: " & pressed, vbInformation, "엑셀 리본 빌더"\nEnd Sub`;
        } else if (ctrl.type === 'editbox') {
            defaultCode = `Sub ${ctrl.onAction}(control As IRibbonControl, text As String)\n    MsgBox "입력된 텍스트: " & text, vbInformation, "엑셀 리본 빌더"\nEnd Sub`;
        } else if (['combobox', 'dropdown'].includes(ctrl.type)) {
            defaultCode = `Sub ${ctrl.onAction}(control As IRibbonControl, textOrId As String, Optional index As Integer)\n    MsgBox "선택된 값: " & textOrId, vbInformation, "엑셀 리본 빌더"\nEnd Sub`;
        } else if (ctrl.type === 'gallery') {
            defaultCode = `Sub ${ctrl.onAction}(control As IRibbonControl, id As String, index As Integer)\n    MsgBox "갤러리 선택 항목: " & id & " (인덱스: " & index & ")", vbInformation, "엑셀 리본 빌더"\nEnd Sub`;
        } else if (ctrl.type === 'togglebutton') {
            defaultCode = `Sub ${ctrl.onAction}(control As IRibbonControl, pressed As Boolean)\n    MsgBox "토글 상태: " & pressed, vbInformation, "엑셀 리본 빌더"\nEnd Sub`;
        } else {
            defaultCode = `Sub ${ctrl.onAction}(control As IRibbonControl)\n    MsgBox "${ctrl.label || '기능'}을 실행합니다.", vbInformation, "엑셀 리본 빌더"\nEnd Sub`;
        }
        ctrl.vbaCustomCode = defaultCode;
        ctrl.vbaPresetId = ""; // default as Custom
    }
}

function loadState(state) {
    ribbonState = state;
    selectedElement = null;
    
    // Auto initialize VBA properties for loaded demo presets
    ribbonState.tabs.forEach(tab => {
        if (tab.groups) {
            tab.groups.forEach(group => {
                if (group.controls) {
                    group.controls.forEach(ctrl => {
                        initializeControlVbaProperties(ctrl);
                    });
                }
            });
        }
    });

    if (ribbonState.tabs.length > 0) {
        selectedElement = { type: 'tab', path: [0] };
    }
    updateRibbonUI();
}

function updateRibbonUI() {
    renderRibbonSimulator();
    renderPropertyEditor();
    compileCodes();
}

// --- 4.5 Unified Icon Html Renderer (SVG vs Emoji vs Image File) ---
function getIconHtml(imageMsoValue, size = 18) {
    if (!imageMsoValue) {
        return `<circle cx="12" cy="12" r="10" fill="none" stroke="currentColor" stroke-width="2"></circle>`;
    }

    // 1. Local SVG File representation
    if (imageMsoValue.startsWith("icons/")) {
        return `<img src="${imageMsoValue}" class="mso-icon-img" style="width:${size}px; height:${size}px;">`;
    }

    // 2. Unicode Emojis
    const isEmoji = imageMsoValue.length <= 4 && /\p{Emoji}/u.test(imageMsoValue);
    if (isEmoji) {
        return `<span class="emoji-render-view" style="font-size: ${size}px; line-height:1; display:inline-block; font-family:'Segoe UI Emoji','Apple Color Emoji',sans-serif;">${imageMsoValue}</span>`;
    }

    // 3. Normal SVG Embedded Fallbacks
    if (window.getIconSvg) {
        return window.getIconSvg(imageMsoValue, size);
    }

    return `<span style="font-size:10px;">★</span>`;
}

// --- 5. Interactive Simulated Ribbon Renderer ---
// --- 5. Interactive Simulated Ribbon Renderer ---
function renderSingleControl(ctrl, activeTabIdx, grpIdx, ctrlIdx, totalControlsCount) {
    if (ctrl.type === 'separator') {
        const sepEl = document.createElement("div");
        sepEl.className = "ribbon-separator-sim";
        
        // Drag and Drop reordering support
        sepEl.draggable = true;
        sepEl.addEventListener("dragstart", (e) => {
            e.stopPropagation();
            e.dataTransfer.setData("text/plain", `move:${activeTabIdx},${grpIdx},${ctrlIdx}`);
            e.dataTransfer.effectAllowed = "move";
            sepEl.style.opacity = "0.5";
        });
        sepEl.addEventListener("dragend", () => {
            sepEl.style.opacity = "";
        });

        sepEl.addEventListener("dragover", (e) => {
            e.preventDefault();
            e.stopPropagation();
            e.dataTransfer.dropEffect = "move";
            sepEl.classList.add("drag-hover-control");
        });
        sepEl.addEventListener("dragenter", (e) => {
            e.preventDefault();
            e.stopPropagation();
            sepEl.classList.add("drag-hover-control");
        });
        sepEl.addEventListener("dragleave", () => {
            sepEl.classList.remove("drag-hover-control");
        });
        sepEl.addEventListener("drop", (e) => {
            e.preventDefault();
            e.stopPropagation();
            sepEl.classList.remove("drag-hover-control");
            const data = e.dataTransfer.getData("text/plain");
            if (data) {
                handleDropOnControl(data, activeTabIdx, grpIdx, ctrlIdx);
            }
        });

        const isSepSelected = selectedElement && selectedElement.type === 'control'
            && selectedElement.path[0] === activeTabIdx 
            && selectedElement.path[1] === grpIdx
            && selectedElement.path[2] === ctrlIdx;

        if (isSepSelected) {
            sepEl.style.backgroundColor = "var(--excel-green-accent)";
            sepEl.style.width = "2px";
            
            const overlay = createOverlay(
                () => deleteElement('control', [activeTabIdx, grpIdx, ctrlIdx]),
                ctrlIdx > 0 ? () => moveElement('control', [activeTabIdx, grpIdx, ctrlIdx], -1) : null,
                ctrlIdx < totalControlsCount - 1 ? () => moveElement('control', [activeTabIdx, grpIdx, ctrlIdx], 1) : null,
                null,
                null
            );
            sepEl.appendChild(overlay);
        }

        sepEl.addEventListener("click", (e) => {
            e.stopPropagation();
            selectedElement = { type: 'control', path: [activeTabIdx, grpIdx, ctrlIdx] };
            updateRibbonUI();
        });

        return sepEl;
    }

    const ctrlEl = document.createElement("div");
    const isCtrlSelected = selectedElement && selectedElement.type === 'control'
        && selectedElement.path[0] === activeTabIdx 
        && selectedElement.path[1] === grpIdx
        && selectedElement.path[2] === ctrlIdx;

    // Drag and Drop reordering support for regular controls
    ctrlEl.draggable = true;
    ctrlEl.addEventListener("dragstart", (e) => {
        e.stopPropagation();
        e.dataTransfer.setData("text/plain", `move:${activeTabIdx},${grpIdx},${ctrlIdx}`);
        e.dataTransfer.effectAllowed = "move";
        ctrlEl.style.opacity = "0.5";
    });
    ctrlEl.addEventListener("dragend", () => {
        ctrlEl.style.opacity = "";
    });

    ctrlEl.addEventListener("dragover", (e) => {
        e.preventDefault();
        e.stopPropagation();
        e.dataTransfer.dropEffect = "move";
        ctrlEl.classList.add("drag-hover-control");
    });
    ctrlEl.addEventListener("dragenter", (e) => {
        e.preventDefault();
        e.stopPropagation();
        ctrlEl.classList.add("drag-hover-control");
    });
    ctrlEl.addEventListener("dragleave", () => {
        ctrlEl.classList.remove("drag-hover-control");
    });
    ctrlEl.addEventListener("drop", (e) => {
        e.preventDefault();
        e.stopPropagation();
        ctrlEl.classList.remove("drag-hover-control");
        const data = e.dataTransfer.getData("text/plain");
        if (data) {
            handleDropOnControl(data, activeTabIdx, grpIdx, ctrlIdx);
        }
    });

    // Unified rendering properties
    if (ctrl.type === 'checkbox') {
        ctrlEl.className = "ribbon-control-sim type-checkbox";
        if (isCtrlSelected) ctrlEl.classList.add("selected");
        ctrlEl.innerHTML = `<div class="checkbox-sim-box">${ctrl.checked ? '✓' : ''}</div><span class="control-label">${ctrl.label}</span>`;
        
    } else if (ctrl.type === 'editbox') {
        ctrlEl.className = "ribbon-control-sim type-editbox";
        if (isCtrlSelected) ctrlEl.classList.add("selected");
        ctrlEl.innerHTML = `<span class="control-label">${ctrl.label}</span><div class="editbox-sim-input">${ctrl.text || ''}</div>`;
        
    } else if (ctrl.type === 'combobox') {
        ctrlEl.className = "ribbon-control-sim type-combobox";
        if (isCtrlSelected) ctrlEl.classList.add("selected");
        ctrlEl.innerHTML = `<span class="control-label">${ctrl.label}</span><div class="dropdown-sim-select-light"><span>${ctrl.text || ''}</span><span class="dropdown-arrow-sim-light">▼</span></div>`;
        
    } else if (ctrl.type === 'dropdown') {
        ctrlEl.className = "ribbon-control-sim type-dropdown";
        if (isCtrlSelected) ctrlEl.classList.add("selected");
        const firstLabel = ctrl.items && ctrl.items.length > 0 ? ctrl.items[0].label : '선택 없음';
        ctrlEl.innerHTML = `<span class="control-label">${ctrl.label}</span><div class="dropdown-sim-select-light"><span>${firstLabel}</span><span class="dropdown-arrow-sim-light">▼</span></div>`;
        
    } else if (ctrl.type === 'box') {
        ctrlEl.className = "ribbon-control-sim type-box";
        if (isCtrlSelected) ctrlEl.classList.add("selected");
        ctrlEl.innerHTML = `<span class="control-label" style="font-style:italic; color:#888; font-size:10px;">📦 상자 [${ctrl.boxStyle || 'vertical'}, ${ctrl.stackLimit || 3}개]</span><span class="control-label">${ctrl.label}</span>`;
        
    } else if (ctrl.type === 'labelcontrol') {
        ctrlEl.className = "ribbon-control-sim type-labelcontrol";
        if (isCtrlSelected) ctrlEl.classList.add("selected");
        ctrlEl.textContent = ctrl.label || '';
        
    } else if (ctrl.type === 'menu') {
        ctrlEl.className = `ribbon-control-sim size-${ctrl.size || 'large'} type-menu`;
        if (isCtrlSelected) ctrlEl.classList.add("selected");
        const icon = getIconHtml(ctrl.imageMso, ctrl.size === 'large' ? 24 : 14);
        ctrlEl.innerHTML = `<div class="control-icon">${icon}</div><span class="control-label">${ctrl.label} ▾</span>`;
        
    } else if (ctrl.type === 'togglebutton') {
        // Toggle Button acts like radio buttons
        ctrlEl.className = `ribbon-control-sim size-${ctrl.size || 'large'} type-togglebutton`;
        if (isCtrlSelected) ctrlEl.classList.add("selected");
        if (ctrl.checked) ctrlEl.classList.add("pressed"); // Visual pressed state
        const radioCircle = ctrl.checked ? '●' : '○';
        
        // No icon/image for togglebutton!
        if (ctrl.size === 'large') {
            ctrlEl.innerHTML = `<span class="control-label"><span class="radio-sim-circle">${radioCircle}</span> ${ctrl.label}</span>`;
        } else {
            ctrlEl.innerHTML = `<span class="radio-sim-circle">${radioCircle}</span><span class="control-label">${ctrl.label}</span>`;
        }
        
    } else {
        // Standard button
        ctrlEl.className = `ribbon-control-sim size-${ctrl.size || 'large'}`;
        if (isCtrlSelected) ctrlEl.classList.add("selected");
        const icon = getIconHtml(ctrl.imageMso, ctrl.size === 'large' ? 24 : 14);
        ctrlEl.innerHTML = `<div class="control-icon">${icon}</div><span class="control-label">${ctrl.label}</span>`;
    }

    // Toggle interaction on click directly in simulator
    ctrlEl.addEventListener("click", (e) => {
        e.stopPropagation();
        selectedElement = { type: 'control', path: [activeTabIdx, grpIdx, ctrlIdx] };
        
        if (ctrl.type === 'checkbox') {
            ctrl.checked = !ctrl.checked;
        } else if (ctrl.type === 'togglebutton') {
            ctrl.checked = !ctrl.checked;
            if (ctrl.checked) {
                // Enforce single-selection radio grouping
                const group = ribbonState.tabs[activeTabIdx].groups[grpIdx];
                const targetGroup = ctrl.toggleGroup || "";
                group.controls.forEach(c => {
                    if (c !== ctrl && c.type === 'togglebutton' && (c.toggleGroup || "") === targetGroup) {
                        c.checked = false;
                    }
                });
            }
        }
        updateRibbonUI();
    });

    // Control Overlays for actions (Upgraded 4-way D-Pad)
    if (isCtrlSelected) {
        const isLarge = ctrl.type === 'separator' || ctrl.size === 'large';
        let overlay;
        if (isLarge) {
            overlay = createOverlay(
                () => deleteElement('control', [activeTabIdx, grpIdx, ctrlIdx]),
                ctrlIdx > 0 ? () => moveElement('control', [activeTabIdx, grpIdx, ctrlIdx], -1) : null,
                ctrlIdx < totalControlsCount - 1 ? () => moveElement('control', [activeTabIdx, grpIdx, ctrlIdx], 1) : null,
                null,
                null
            );
        } else {
            overlay = createOverlay(
                () => deleteElement('control', [activeTabIdx, grpIdx, ctrlIdx]),
                ctrlIdx > 0 ? () => moveElement('control', [activeTabIdx, grpIdx, ctrlIdx], -3) : null,
                ctrlIdx < totalControlsCount - 1 ? () => moveElement('control', [activeTabIdx, grpIdx, ctrlIdx], 3) : null,
                ctrlIdx > 0 ? () => moveElement('control', [activeTabIdx, grpIdx, ctrlIdx], -1) : null,
                ctrlIdx < totalControlsCount - 1 ? () => moveElement('control', [activeTabIdx, grpIdx, ctrlIdx], 1) : null
            );
        }
        ctrlEl.appendChild(overlay);
    }

    return ctrlEl;
}

function renderRibbonSimulator() {
    const tabsBar = document.getElementById("sim-tabs-bar");
    const ribbonBody = document.getElementById("sim-ribbon-body");
    
    tabsBar.innerHTML = "";
    ribbonBody.innerHTML = "";

    // 1. Render Tabs (11px size set by CSS)
    ribbonState.tabs.forEach((tab, tabIdx) => {
        const tabEl = document.createElement("div");
        tabEl.className = "excel-tab-header custom-tab-indicator";
        if (tab.isStandardTab) {
            tabEl.className = "excel-tab-header file-tab";
        }
        
        const isSelected = selectedElement && selectedElement.type === 'tab' && selectedElement.path[0] === tabIdx;
        if (isSelected) {
            tabEl.classList.add("active");
        }
        
        tabEl.textContent = tab.label || "새 탭";
        tabEl.addEventListener("click", (e) => {
            e.stopPropagation();
            selectedElement = { type: 'tab', path: [tabIdx] };
            updateRibbonUI();
        });

        // Add control overlays for CUSTOM tabs (Tab reordering left/right support!)
        if (isSelected && !tab.isStandardTab) {
            const overlay = createOverlay(
                () => deleteElement('tab', [tabIdx]),
                tabIdx > 0 ? () => moveElement('tab', [tabIdx], -1) : null,
                tabIdx < ribbonState.tabs.length - 1 ? () => moveElement('tab', [tabIdx], 1) : null
            );
            tabEl.appendChild(overlay);
        }

        tabsBar.appendChild(tabEl);
    });

    // 2. Render Groups & Controls in the Active Tab
    const activeTabIdx = selectedElement ? selectedElement.path[0] : 0;
    const activeTab = ribbonState.tabs[activeTabIdx];
    
    if (activeTab && activeTab.groups) {
        activeTab.groups.forEach((group, grpIdx) => {
            const groupEl = document.createElement("div");
            groupEl.className = "ribbon-group-sim";
            
            const isSelected = selectedElement && selectedElement.type === 'group' 
                && selectedElement.path[0] === activeTabIdx 
                && selectedElement.path[1] === grpIdx;
                
            if (isSelected) {
                groupEl.classList.add("selected");
            }

            groupEl.addEventListener("click", (e) => {
                e.stopPropagation();
                selectedElement = { type: 'group', path: [activeTabIdx, grpIdx] };
                updateRibbonUI();
            });

            // Drag over / Drag enter / Drag leave / Drop for groups
            groupEl.addEventListener("dragover", (e) => {
                e.preventDefault();
                e.stopPropagation();
                e.dataTransfer.dropEffect = "copy";
                groupEl.classList.add("drag-over");
            });
            groupEl.addEventListener("dragenter", (e) => {
                e.preventDefault();
                e.stopPropagation();
                groupEl.classList.add("drag-over");
            });
            groupEl.addEventListener("dragleave", () => {
                groupEl.classList.remove("drag-over");
            });

            groupEl.addEventListener("drop", (e) => {
                e.preventDefault();
                groupEl.classList.remove("drag-over");
                
                const data = e.dataTransfer.getData("text/plain");
                if (data) {
                    if (data.startsWith("move:")) {
                        const parts = data.substring(5).split(",");
                        if (parts.length === 3) {
                            const srcTabIdx = parseInt(parts[0], 10);
                            const srcGrpIdx = parseInt(parts[1], 10);
                            const srcCtrlIdx = parseInt(parts[2], 10);
                            
                            const destControls = ribbonState.tabs[activeTabIdx].groups[grpIdx].controls;
                            moveControlToPosition(srcTabIdx, srcGrpIdx, srcCtrlIdx, activeTabIdx, grpIdx, destControls.length);
                        }
                    } else {
                        addNewControlToSpecificGroup(data, activeTabIdx, grpIdx);
                    }
                }
            });

            // Group Label (11px uniform)
            const labelEl = document.createElement("div");
            labelEl.className = "ribbon-group-label-sim";
            labelEl.textContent = group.label || "새 그룹";
            groupEl.appendChild(labelEl);

            // Group controls holder
            const controlsWrapper = document.createElement("div");
            controlsWrapper.className = "ribbon-group-controls";

            if (group.controls && group.controls.length > 0) {
                let i = 0;
                while (i < group.controls.length) {
                    const ctrl = group.controls[i];
                    
                    if (ctrl.type === 'box') {
                        const stackLimit = ctrl.stackLimit || 3;
                        const children = [];
                        let j = i + 1;
                        while (j < group.controls.length && children.length < stackLimit) {
                            const candidate = group.controls[j];
                            if (candidate.type === 'separator' || candidate.type === 'box' || candidate.type === 'buttongroup') {
                                break;
                            }
                            children.push(candidate);
                            j++;
                        }
                        
                        const boxEl = renderSingleControl(ctrl, activeTabIdx, grpIdx, i, group.controls.length);
                        
                        const boxChildrenContainer = document.createElement("div");
                        boxChildrenContainer.className = `ribbon-box-children-sim ${ctrl.boxStyle || 'vertical'}`;
                        boxChildrenContainer.style.display = "flex";
                        boxChildrenContainer.style.flexDirection = (ctrl.boxStyle === 'horizontal') ? 'row' : 'column';
                        boxChildrenContainer.style.gap = "4px";
                        boxChildrenContainer.style.marginTop = "4px";
                        boxChildrenContainer.style.padding = "4px";
                        boxChildrenContainer.style.border = "1px dashed #ccc";
                        boxChildrenContainer.style.borderRadius = "4px";
                        boxChildrenContainer.style.flexGrow = "1";
                        boxChildrenContainer.style.justifyContent = "center";
                        
                        children.forEach(child => {
                            const childIdx = group.controls.indexOf(child);
                            const childEl = renderSingleControl(child, activeTabIdx, grpIdx, childIdx, group.controls.length);
                            boxChildrenContainer.appendChild(childEl);
                        });
                        
                        boxEl.appendChild(boxChildrenContainer);
                        controlsWrapper.appendChild(boxEl);
                        
                        i = j;
                    } else {
                        const ctrlEl = renderSingleControl(ctrl, activeTabIdx, grpIdx, i, group.controls.length);
                        controlsWrapper.appendChild(ctrlEl);
                        i++;
                    }
                }
            } else {
                const emptyInd = document.createElement("div");
                emptyInd.className = "empty-group-indicator";
                emptyInd.textContent = "컨트롤 비어있음";
                controlsWrapper.appendChild(emptyInd);
            }

            groupEl.appendChild(controlsWrapper);

            // Group Overlays for actions
            if (isSelected) {
                const overlay = createOverlay(
                    () => deleteElement('group', [activeTabIdx, grpIdx]),
                    grpIdx > 0 ? () => moveElement('group', [activeTabIdx, grpIdx], -1) : null,
                    grpIdx < activeTab.groups.length - 1 ? () => moveElement('group', [activeTabIdx, grpIdx], 1) : null
                );
                groupEl.appendChild(overlay);
            }

            ribbonBody.appendChild(groupEl);
        });
    }
}

// --- 6. Element Creation, Deletion & Movement Helpers ---
function isControlIdExists(id) {
    if (!ribbonState || !ribbonState.tabs) return false;
    for (const tab of ribbonState.tabs) {
        if (tab.id === id) return true;
        if (tab.groups) {
            for (const group of tab.groups) {
                if (group.id === id) return true;
                if (group.controls) {
                    for (const ctrl of group.controls) {
                        if (ctrl.id === id) return true;
                        if (ctrl.controls) {
                            for (const subCtrl of ctrl.controls) {
                                if (subCtrl.id === id) return true;
                            }
                        }
                    }
                }
            }
        }
    }
    return false;
}

function generateUniqueControlId(type) {
    let prefix = "btn";
    if (type === 'separator') prefix = "sep";
    else if (type === 'checkbox') prefix = "chk";
    else if (type === 'editbox') prefix = "edit";
    else if (type === 'combobox') prefix = "combo";
    else if (type === 'dropdown') prefix = "drop";
    else if (type === 'gallery') prefix = "gallery";
    else if (type === 'box') prefix = "box";
    else if (type === 'buttongroup') prefix = "btnGroup";
    else if (type === 'labelcontrol') prefix = "lbl";
    else if (type === 'menu') prefix = "menu";
    else if (type === 'splitbutton') prefix = "split";
    else if (type === 'togglebutton') prefix = "toggle";
    
    let counter = 1;
    while (true) {
        const candidateId = `${prefix}_Custom_${counter}`;
        if (!isControlIdExists(candidateId)) {
            return candidateId;
        }
        counter++;
    }
}

function addNewTab() {
    const newIdx = ribbonState.tabs.length + 1;
    const tab = {
        id: `customTab_${newIdx}`,
        label: `맞춤형 탭 ${newIdx}`,
        visible: true,
        isStandardTab: false,
        idMso: "",
        groups: []
    };
    ribbonState.tabs.push(tab);
    selectedElement = { type: 'tab', path: [ribbonState.tabs.length - 1] };
    updateRibbonUI();
}

function addNewGroup() {
    if (ribbonState.tabs.length === 0) {
        alert("그룹을 추가할 탭이 없습니다. 먼저 '새 탭'을 생성해 주세요.");
        return;
    }
    
    let activeTabIdx = selectedElement ? selectedElement.path[0] : 0;
    if (activeTabIdx >= ribbonState.tabs.length) activeTabIdx = 0;
    const activeTab = ribbonState.tabs[activeTabIdx];
    
    const newIdx = activeTab.groups.length + 1;
    const group = {
        id: `group_CustomGrp_${activeTabIdx + 1}_${newIdx}`,
        label: `기능 그룹 ${newIdx}`,
        visible: true,
        controls: []
    };
    
    activeTab.groups.push(group);
    selectedElement = { type: 'group', path: [activeTabIdx, activeTab.groups.length - 1] };
    updateRibbonUI();
}

function addNewControlToSpecificGroup(type, tabIdx, grpIdx) {
    if (ribbonState.tabs.length === 0) return;
    const tab = ribbonState.tabs[tabIdx];
    if (!tab || !tab.groups || !tab.groups[grpIdx]) return;
    const group = tab.groups[grpIdx];

    const newIdx = group.controls.length + 1;
    const newId = generateUniqueControlId(type);
    let control = { id: newId, type: type };

    if (type === 'separator') {
        control = { id: newId, type: "separator" };
    } else if (type === 'checkbox') {
        control = { id: newId, type: "checkbox", label: `선택 체크박스 ${newIdx}`, checked: false, onAction: `${newId}_Click`, enabled: true, visible: true };
    } else if (type === 'editbox') {
        control = { id: newId, type: "editbox", label: `텍스트 입력 ${newIdx}`, text: "", onAction: `${newId}_Change`, enabled: true, visible: true };
    } else if (type === 'combobox') {
        control = { id: newId, type: "combobox", label: `콤보 박스 ${newIdx}`, text: "옵션 1", onAction: `${newId}_Change`, enabled: true, visible: true, items: [{ id: `${newId}_item1`, label: "옵션 1" }, { id: `${newId}_item2`, label: "옵션 2" }] };
    } else if (type === 'dropdown') {
        control = { id: newId, type: "dropdown", label: `선택 목록 ${newIdx}`, onAction: `${newId}_Change`, enabled: true, visible: true, items: [{ id: `${newId}_item1`, label: "옵션 1" }, { id: `${newId}_item2`, label: "옵션 2" }] };
    } else if (type === 'box') {
        control = { id: newId, type: "box", label: `상자 ${newIdx}`, boxStyle: "vertical", stackLimit: 3 };
    } else if (type === 'buttongroup') {
        control = { id: newId, type: "buttongroup", label: `버튼 그룹 ${newIdx}` };
    } else if (type === 'labelcontrol') {
        control = { id: newId, type: "labelcontrol", label: `레이블 정보 ${newIdx}` };
    } else if (type === 'menu') {
        control = { 
            id: newId, 
            type: "menu", 
            label: `하위 메뉴 ${newIdx}`, 
            size: "large", 
            imageMso: "icons/2024-microsoft-365-content-icons/Microsoft Blue/48x48 Light Blue Icon/Settings.svg", 
            enabled: true, 
            visible: true,
            items: [
                { id: `${newId}_sub1`, label: "서브 명령 1", onAction: `${newId}_sub1_Click` },
                { id: `${newId}_sub2`, label: "서브 명령 2", onAction: `${newId}_sub2_Click` }
            ]
        };
    } else {
        // button or togglebutton
        control = { id: newId, type: type, label: `${type === 'togglebutton' ? '토글 버튼' : '새 버튼'} ${newIdx}`, size: "large", imageMso: "icons/2024-microsoft-365-content-icons/Microsoft Blue/48x48 Light Blue Icon/Top Speed.svg", onAction: `${newId}_Click`, enabled: true, visible: true, checked: false };
    }

    initializeControlVbaProperties(control);
    group.controls.push(control);
    selectedElement = { type: 'control', path: [tabIdx, grpIdx, group.controls.length - 1] };
    updateRibbonUI();
}

function addNewControl(type) {
    if (ribbonState.tabs.length === 0) {
        alert("컨트롤을 배치할 그룹이 없습니다. 먼저 '새 탭'과 '새 그룹'을 추가해 주세요.");
        return;
    }
    
    let activeTabIdx = selectedElement ? selectedElement.path[0] : 0;
    if (activeTabIdx >= ribbonState.tabs.length) activeTabIdx = 0;
    const activeTab = ribbonState.tabs[activeTabIdx];

    if (activeTab.groups.length === 0) {
        alert("컨트롤을 배치할 그룹이 없습니다. 먼저 '새 그룹'을 추가해 주세요.");
        return;
    }

    let activeGrpIdx = selectedElement && selectedElement.type !== 'tab' ? selectedElement.path[1] : 0;
    if (activeGrpIdx >= activeTab.groups.length) activeGrpIdx = 0;

    addNewControlToSpecificGroup(type, activeTabIdx, activeGrpIdx);
}

function deleteElement(type, path) {
    if (type === 'tab') {
        ribbonState.tabs.splice(path[0], 1);
        selectedElement = ribbonState.tabs.length > 0 ? { type: 'tab', path: [0] } : null;
    } else if (type === 'group') {
        ribbonState.tabs[path[0]].groups.splice(path[1], 1);
        selectedElement = { type: 'tab', path: [path[0]] };
    } else if (type === 'control') {
        ribbonState.tabs[path[0]].groups[path[1]].controls.splice(path[2], 1);
        selectedElement = { type: 'group', path: [path[0], path[1]] };
    }
    updateRibbonUI();
}

function moveElement(type, path, direction) {
    if (type === 'tab') {
        const index = path[0];
        const targetIndex = index + direction;
        if (targetIndex >= 0 && targetIndex < ribbonState.tabs.length) {
            const [movedTab] = ribbonState.tabs.splice(index, 1);
            ribbonState.tabs.splice(targetIndex, 0, movedTab);
            selectedElement = { type: 'tab', path: [targetIndex] };
        }
    } else if (type === 'group') {
        const tabIdx = path[0];
        const index = path[1];
        const targetIndex = index + direction;
        const groups = ribbonState.tabs[tabIdx].groups;
        if (targetIndex >= 0 && targetIndex < groups.length) {
            const [movedGroup] = groups.splice(index, 1);
            groups.splice(targetIndex, 0, movedGroup);
            selectedElement = { type: 'group', path: [tabIdx, targetIndex] };
        }
    } else if (type === 'control') {
        const tabIdx = path[0];
        const grpIdx = path[1];
        const index = path[2];
        const targetIndex = index + direction;
        const controls = ribbonState.tabs[tabIdx].groups[grpIdx].controls;
        if (targetIndex >= 0 && targetIndex < controls.length) {
            const [movedControl] = controls.splice(index, 1);
            controls.splice(targetIndex, 0, movedControl);
            selectedElement = { type: 'control', path: [tabIdx, grpIdx, targetIndex] };
        }
    }
    updateRibbonUI();
}

function moveControlToPosition(srcTabIdx, srcGrpIdx, srcCtrlIdx, destTabIdx, destGrpIdx, destCtrlIdx) {
    const srcTab = ribbonState.tabs[srcTabIdx];
    if (!srcTab || !srcTab.groups || !srcTab.groups[srcGrpIdx]) return;
    const srcGroup = srcTab.groups[srcGrpIdx];
    const controlToMove = srcGroup.controls[srcCtrlIdx];
    if (!controlToMove) return;
    
    const destTab = ribbonState.tabs[destTabIdx];
    if (!destTab || !destTab.groups || !destTab.groups[destGrpIdx]) return;
    const destGroup = destTab.groups[destGrpIdx];
    
    // Remove from source
    srcGroup.controls.splice(srcCtrlIdx, 1);
    
    // Calculate new destination index
    let targetIdx = destCtrlIdx;
    // If moving within the same group and the source was before the destination, the index shifts down by 1!
    if (srcTabIdx === destTabIdx && srcGrpIdx === destGrpIdx) {
        if (srcCtrlIdx < destCtrlIdx) {
            targetIdx--;
        }
    }
    
    // Insert at destination
    destGroup.controls.splice(targetIdx, 0, controlToMove);
    
    // Update selected element to follow the moved control
    selectedElement = { type: 'control', path: [destTabIdx, destGrpIdx, targetIdx] };
    updateRibbonUI();
}

function insertNewControlAtPosition(type, tabIdx, grpIdx, destCtrlIdx) {
    if (ribbonState.tabs.length === 0) return;
    const tab = ribbonState.tabs[tabIdx];
    if (!tab || !tab.groups || !tab.groups[grpIdx]) return;
    const group = tab.groups[grpIdx];

    const newIdx = group.controls.length + 1;
    const newId = generateUniqueControlId(type);
    let control = { id: newId, type: type };

    if (type === 'separator') {
        control = { id: newId, type: "separator" };
    } else if (type === 'checkbox') {
        control = { id: newId, type: "checkbox", label: `선택 체크박스 ${newIdx}`, checked: false, onAction: `${newId}_Click`, enabled: true, visible: true };
    } else if (type === 'editbox') {
        control = { id: newId, type: "editbox", label: `텍스트 입력 ${newIdx}`, text: "", onAction: `${newId}_Change`, enabled: true, visible: true };
    } else if (type === 'combobox') {
        control = { id: newId, type: "combobox", label: `콤보 박스 ${newIdx}`, text: "옵션 1", onAction: `${newId}_Change`, enabled: true, visible: true, items: [{ id: `${newId}_item1`, label: "옵션 1" }, { id: `${newId}_item2`, label: "옵션 2" }] };
    } else if (type === 'dropdown') {
        control = { id: newId, type: "dropdown", label: `선택 목록 ${newIdx}`, onAction: `${newId}_Change`, enabled: true, visible: true, items: [{ id: `${newId}_item1`, label: "옵션 1" }, { id: `${newId}_item2`, label: "옵션 2" }] };
    } else if (type === 'box') {
        control = { id: newId, type: "box", label: `상자 ${newIdx}`, boxStyle: "vertical", stackLimit: 3 };
    } else if (type === 'buttongroup') {
        control = { id: newId, type: "buttongroup", label: `버튼 그룹 ${newIdx}` };
    } else if (type === 'labelcontrol') {
        control = { id: newId, type: "labelcontrol", label: `레이블 정보 ${newIdx}` };
    } else if (type === 'menu') {
        control = { 
            id: newId, 
            type: "menu", 
            label: `하위 메뉴 ${newIdx}`, 
            size: "large", 
            imageMso: "icons/2024-microsoft-365-content-icons/Microsoft Blue/48x48 Light Blue Icon/Settings.svg", 
            enabled: true, 
            visible: true,
            items: [
                { id: `${newId}_sub1`, label: "서브 명령 1", onAction: `${newId}_sub1_Click` },
                { id: `${newId}_sub2`, label: "서브 명령 2", onAction: `${newId}_sub2_Click` }
            ]
        };
    } else {
        // button or togglebutton
        control = { id: newId, type: type, label: `${type === 'togglebutton' ? '토글 버튼' : '새 버튼'} ${newIdx}`, size: "large", imageMso: "icons/2024-microsoft-365-content-icons/Microsoft Blue/48x48 Light Blue Icon/Top Speed.svg", onAction: `${newId}_Click`, enabled: true, visible: true, checked: false };
    }

    initializeControlVbaProperties(control);
    group.controls.splice(destCtrlIdx, 0, control);
    selectedElement = { type: 'control', path: [tabIdx, grpIdx, destCtrlIdx] };
    updateRibbonUI();
}

function handleDropOnControl(data, destTabIdx, destGrpIdx, destCtrlIdx) {
    if (data.startsWith("move:")) {
        const parts = data.substring(5).split(",");
        if (parts.length === 3) {
            const srcTabIdx = parseInt(parts[0], 10);
            const srcGrpIdx = parseInt(parts[1], 10);
            const srcCtrlIdx = parseInt(parts[2], 10);
            moveControlToPosition(srcTabIdx, srcGrpIdx, srcCtrlIdx, destTabIdx, destGrpIdx, destCtrlIdx);
        }
    } else {
        insertNewControlAtPosition(data, destTabIdx, destGrpIdx, destCtrlIdx);
    }
}

// Helper to create control overlays for reordering & deleting elements in simulator
function createOverlay(onDelete, onMoveLeft, onMoveRight, onMoveUp, onMoveDown) {
    const overlay = document.createElement("div");
    overlay.className = "sim-controls-overlay";
    
    // 1. D-Pad Left
    if (onMoveLeft) {
        const btnLeft = document.createElement("button");
        btnLeft.type = "button";
        btnLeft.className = "overlay-btn move-left";
        btnLeft.innerHTML = "◀";
        btnLeft.title = "왼쪽(이전 컬럼) 이동";
        btnLeft.addEventListener("click", (e) => {
            e.stopPropagation();
            onMoveLeft();
        });
        overlay.appendChild(btnLeft);
    }
    
    // 2. D-Pad Up
    if (onMoveUp) {
        const btnUp = document.createElement("button");
        btnUp.type = "button";
        btnUp.className = "overlay-btn move-up";
        btnUp.innerHTML = "▲";
        btnUp.title = "위로 이동";
        btnUp.addEventListener("click", (e) => {
            e.stopPropagation();
            onMoveUp();
        });
        overlay.appendChild(btnUp);
    }

    // 3. D-Pad Down
    if (onMoveDown) {
        const btnDown = document.createElement("button");
        btnDown.type = "button";
        btnDown.className = "overlay-btn move-down";
        btnDown.innerHTML = "▼";
        btnDown.title = "아래로 이동";
        btnDown.addEventListener("click", (e) => {
            e.stopPropagation();
            onMoveDown();
        });
        overlay.appendChild(btnDown);
    }

    // 4. D-Pad Right
    if (onMoveRight) {
        const btnRight = document.createElement("button");
        btnRight.type = "button";
        btnRight.className = "overlay-btn move-right";
        btnRight.innerHTML = "▶";
        btnRight.title = "오른쪽(다음 컬럼) 이동";
        btnRight.addEventListener("click", (e) => {
            e.stopPropagation();
            onMoveRight();
        });
        overlay.appendChild(btnRight);
    }

    // 5. Delete Action
    if (onDelete) {
        const btnDel = document.createElement("button");
        btnDel.type = "button";
        btnDel.className = "overlay-btn delete";
        btnDel.innerHTML = "✖";
        btnDel.title = "삭제";
        btnDel.addEventListener("click", (e) => {
            e.stopPropagation();
            onDelete();
        });
        overlay.appendChild(btnDel);
    }

    return overlay;
}

// --- 7. Selected Element Property Panel Generator ---
function renderPropertyEditor() {
    const editorHeader = document.getElementById("editor-header-text");
    const editorBody = document.getElementById("property-editor-fields");
    
    editorBody.innerHTML = "";

    if (!selectedElement) {
        editorHeader.innerHTML = `<span>속성 에디터</span>`;
        editorBody.innerHTML = `<div style="color:var(--text-secondary); text-align:center; padding:12px; font-size:12px; font-style:italic;">시뮬레이터에서 수정할 구성 요소를 클릭하세요.</div>`;
        return;
    }

    const { type, path } = selectedElement;

    if (type === 'tab') {
        const tab = ribbonState.tabs[path[0]];
        editorHeader.innerHTML = `<span>📂 탭 속성 편집</span>`;
        
        editorBody.innerHTML = `
            <div class="property-grid">
                <div class="form-group">
                    <label>컴포넌트 ID</label>
                    <input type="text" id="prop-tab-id" value="${tab.id || ''}">
                </div>
                <div class="form-group">
                    <label>탭 노출명 (레이블)</label>
                    <input type="text" id="prop-tab-label" value="${tab.label || ''}">
                </div>
                <div class="form-group">
                    <label>표준 탭 확장 주입 여부</label>
                    <select id="prop-tab-isStandard">
                        <option value="false" ${!tab.isStandardTab ? 'selected' : ''}>완전 독립형 새 탭 생성</option>
                        <option value="true" ${tab.isStandardTab ? 'selected' : ''}>기존 표준 탭 내부에 삽입</option>
                    </select>
                </div>
                <div class="form-group" id="prop-tab-idMso-wrapper" style="display: ${tab.isStandardTab ? 'block' : 'none'};">
                    <label>타겟 표준 탭 ID (idMso)</label>
                    <select id="prop-tab-idMso">
                        <option value="TabHome" ${tab.idMso === 'TabHome' ? 'selected' : ''}>홈 탭 (TabHome)</option>
                        <option value="TabInsert" ${tab.idMso === 'TabInsert' ? 'selected' : ''}>삽입 탭 (TabInsert)</option>
                        <option value="TabPageLayoutExcel" ${tab.idMso === 'TabPageLayoutExcel' ? 'selected' : ''}>페이지 레이아웃 (TabPageLayoutExcel)</option>
                        <option value="TabFormulas" ${tab.idMso === 'TabFormulas' ? 'selected' : ''}>수식 탭 (TabFormulas)</option>
                        <option value="TabData" ${tab.idMso === 'TabData' ? 'selected' : ''}>데이터 탭 (TabData)</option>
                        <option value="TabReview" ${tab.idMso === 'TabReview' ? 'selected' : ''}>검토 탭 (TabReview)</option>
                        <option value="TabView" ${tab.idMso === 'TabView' ? 'selected' : ''}>보기 탭 (TabView)</option>
                        <option value="TabDeveloper" ${tab.idMso === 'TabDeveloper' ? 'selected' : ''}>개발 도구 (TabDeveloper)</option>
                    </select>
                </div>
                <div class="form-group" id="prop-tab-positiontype-wrapper" style="display: ${!tab.isStandardTab ? 'block' : 'none'};">
                    <label>탭 시작 위치</label>
                    <select id="prop-tab-positiontype">
                        <option value="default" ${tab.positionType === 'default' || !tab.positionType ? 'selected' : ''}>기본 (끝에 배치)</option>
                        <option value="before" ${tab.positionType === 'before' ? 'selected' : ''}>특정 표준 탭 앞에 배치</option>
                        <option value="after" ${tab.positionType === 'after' ? 'selected' : ''}>특정 표준 탭 뒤에 배치</option>
                    </select>
                </div>
                <div class="form-group" id="prop-tab-positiontarget-wrapper" style="display: ${!tab.isStandardTab && (tab.positionType === 'before' || tab.positionType === 'after') ? 'block' : 'none'};">
                    <label>대상 표준 탭</label>
                    <select id="prop-tab-positiontarget">
                        <option value="TabHome" ${tab.positionTarget === 'TabHome' || !tab.positionTarget ? 'selected' : ''}>홈 탭 (TabHome)</option>
                        <option value="TabInsert" ${tab.positionTarget === 'TabInsert' ? 'selected' : ''}>삽입 탭 (TabInsert)</option>
                        <option value="TabPageLayoutExcel" ${tab.positionTarget === 'TabPageLayoutExcel' ? 'selected' : ''}>페이지 레이아웃 (TabPageLayoutExcel)</option>
                        <option value="TabFormulas" ${tab.positionTarget === 'TabFormulas' ? 'selected' : ''}>수식 탭 (TabFormulas)</option>
                        <option value="TabData" ${tab.positionTarget === 'TabData' ? 'selected' : ''}>데이터 탭 (TabData)</option>
                        <option value="TabReview" ${tab.positionTarget === 'TabReview' ? 'selected' : ''}>검토 탭 (TabReview)</option>
                        <option value="TabView" ${tab.positionTarget === 'TabView' ? 'selected' : ''}>보기 탭 (TabView)</option>
                        <option value="TabDeveloper" ${tab.positionTarget === 'TabDeveloper' ? 'selected' : ''}>개발 도구 (TabDeveloper)</option>
                    </select>
                </div>
            </div>
        `;
        
        document.getElementById("prop-tab-isStandard").addEventListener("change", (e) => {
            const val = e.target.value === 'true';
            document.getElementById("prop-tab-idMso-wrapper").style.display = val ? 'block' : 'none';
            document.getElementById("prop-tab-positiontype-wrapper").style.display = val ? 'none' : 'block';
            if (val) {
                document.getElementById("prop-tab-positiontarget-wrapper").style.display = 'none';
            } else {
                const pType = document.getElementById("prop-tab-positiontype").value;
                document.getElementById("prop-tab-positiontarget-wrapper").style.display = (pType === 'before' || pType === 'after') ? 'block' : 'none';
            }
        });

        document.getElementById("prop-tab-positiontype").addEventListener("change", (e) => {
            const val = e.target.value;
            tab.positionType = val;
            document.getElementById("prop-tab-positiontarget-wrapper").style.display = (val === 'before' || val === 'after') ? 'block' : 'none';
            updateRibbonUI();
        });

        document.getElementById("prop-tab-positiontarget").addEventListener("change", (e) => {
            tab.positionTarget = e.target.value;
            updateRibbonUI();
        });

    } else if (type === 'group') {
        const group = ribbonState.tabs[path[0]].groups[path[1]];
        editorHeader.innerHTML = `<span>📦 그룹 속성 편집</span>`;
        
        editorBody.innerHTML = `
            <div class="property-grid">
                <div class="form-group">
                    <label>그룹 ID</label>
                    <input type="text" id="prop-group-id" value="${group.id || ''}">
                </div>
                <div class="form-group">
                    <label>그룹 표시 레이블</label>
                    <input type="text" id="prop-group-label" value="${group.label || ''}">
                </div>
            </div>
        `;
    } else if (type === 'control') {
        const ctrl = ribbonState.tabs[path[0]].groups[path[1]].controls[path[2]];
        
        if (ctrl.type === 'separator') {
            editorHeader.innerHTML = `<span>➖ 구분선</span>`;
            editorBody.innerHTML = `<div class="property-grid"><div class="form-group"><label>컴포넌트 ID</label><input type="text" id="prop-ctrl-id" value="${ctrl.id || ''}"></div></div>`;
            return;
        }

        editorHeader.innerHTML = `<span>🎯 컨트롤 속성 편집 (${ctrl.type.toUpperCase()})</span>`;
        
        let sizeMarkup = '';
        let iconMarkup = '';
        let callbackMarkup = '';
        let extraFields = '';
        let layoutMarkup = '';

        if (!['labelcontrol', 'separator'].includes(ctrl.type)) {
            const labelText = ctrl.type === 'editbox' || ctrl.type === 'combobox' ? 'VBA 변경 함수 (onChange)' : 'VBA 매크로 함수 (onAction)';
            callbackMarkup = `
                <div class="form-group">
                    <label>${labelText}</label>
                    <input type="text" id="prop-ctrl-onaction" value="${ctrl.onAction || ''}">
                </div>
            `;
        }

        const hasSize = ['button', 'togglebutton', 'menu'].includes(ctrl.type);
        if (hasSize) {
            sizeMarkup = `
                <div class="form-group">
                    <label>버튼 크기</label>
                    <div class="radio-toggle">
                        <input type="radio" id="size-large" name="btn-size" value="large" ${ctrl.size === 'large' ? 'checked' : ''}>
                        <label for="size-large">대형 (Large)</label>
                        <input type="radio" id="size-normal" name="btn-size" value="normal" ${ctrl.size === 'normal' ? 'checked' : ''}>
                        <label for="size-normal">일반 (Normal)</label>
                    </div>
                </div>
            `;
        }

        const hasIcon = ['button', 'menu'].includes(ctrl.type);
        if (hasIcon) {
            iconMarkup = `
                <div class="form-group" style="grid-column: span 1;">
                    <label>아이콘 설정</label>
                    <div class="icon-input-wrapper">
                        <input type="text" id="prop-ctrl-icon" value="${ctrl.imageMso || ''}" readonly style="font-size:11px;">
                        <button type="button" class="btn-secondary" id="btn-open-icon-picker">🔍 검색/변경</button>
                        <div class="icon-preview-box" id="prop-icon-preview">
                            ${getIconHtml(ctrl.imageMso, 18)}
                        </div>
                    </div>
                </div>
            `;
        }

        const isSmallCtrl = ctrl.type !== 'separator' && ctrl.type !== 'box' && ctrl.size !== 'large';
        if (isSmallCtrl) {
            layoutMarkup = `
                <div class="form-group" style="grid-column: span 2; display: flex; align-items: center; gap: 8px; margin-top: 8px;">
                    <input type="checkbox" id="prop-ctrl-startnewcolumn" style="width: auto; margin: 0;" ${ctrl.startNewColumn ? 'checked' : ''}>
                    <label for="prop-ctrl-startnewcolumn" style="margin: 0; font-weight: normal; cursor: pointer; user-select: none;">새 열로 시작하기 (Start New Column)</label>
                </div>
            `;
        }

        if (ctrl.type === 'box') {
            extraFields = `
                <div class="form-group">
                    <label>상자 정렬 방식 (boxStyle)</label>
                    <select id="prop-ctrl-boxstyle">
                        <option value="vertical" ${ctrl.boxStyle === 'vertical' ? 'selected' : ''}>세로 정렬 (vertical)</option>
                        <option value="horizontal" ${ctrl.boxStyle === 'horizontal' ? 'selected' : ''}>가로 정렬 (horizontal)</option>
                    </select>
                </div>
                <div class="form-group">
                    <label>상자 내부 하위개체 개수</label>
                    <select id="prop-ctrl-stacklimit">
                        <option value="1" ${ctrl.stackLimit === 1 ? 'selected' : ''}>1개</option>
                        <option value="2" ${ctrl.stackLimit === 2 ? 'selected' : ''}>2개</option>
                        <option value="3" ${ctrl.stackLimit === 3 || !ctrl.stackLimit ? 'selected' : ''}>3개 (기본값)</option>
                    </select>
                </div>
            `;
        } else if (ctrl.type === 'checkbox') {
            extraFields = `
                <div class="form-group">
                    <label>기본 선택 여부</label>
                    <select id="prop-ctrl-checked">
                        <option value="false" ${!ctrl.checked ? 'selected' : ''}>체크 해제 (false)</option>
                        <option value="true" ${ctrl.checked ? 'selected' : ''}>선택 체크 (true)</option>
                    </select>
                </div>
            `;
        } else if (ctrl.type === 'togglebutton') {
            extraFields = `
                <div class="form-group">
                    <label>기본 선택 여부 (토글)</label>
                    <select id="prop-ctrl-checked">
                        <option value="false" ${!ctrl.checked ? 'selected' : ''}>미선택 (false)</option>
                        <option value="true" ${ctrl.checked ? 'selected' : ''}>선택됨 (true)</option>
                    </select>
                </div>
                <div class="form-group">
                    <label>라디오 그룹명 (그룹 내 택1)</label>
                    <input type="text" id="prop-ctrl-togglegroup" value="${ctrl.toggleGroup || ''}" placeholder="예: group1">
                </div>
            `;
        } else if (ctrl.type === 'editbox') {
            extraFields = `
                <div class="form-group">
                    <label>기본 플레이스홀더 / 텍스트</label>
                    <input type="text" id="prop-ctrl-text" value="${ctrl.text || ''}">
                </div>
            `;
        } else if (['combobox', 'dropdown'].includes(ctrl.type)) {
            let itemsMarkup = '';
            if (ctrl.items) {
                ctrl.items.forEach((item, itemIdx) => {
                    itemsMarkup += `
                        <div class="list-item-row" data-idx="${itemIdx}">
                            <input type="text" class="prop-drop-item-id" placeholder="ID" value="${item.id}">
                            <input type="text" class="prop-drop-item-label" placeholder="텍스트" value="${item.label}">
                            <button type="button" class="overlay-btn delete drop-item-del-btn" style="display:flex; width:22px; height:22px;">✖</button>
                        </div>
                    `;
                });
            }

            extraFields = `
                <div class="form-group">
                    <label>기본 선택 값</label>
                    <input type="text" id="prop-ctrl-text" value="${ctrl.text || ''}">
                </div>
                <div class="form-group" style="grid-column: span 2;">
                    <label>동적 아이템 항목 관리</label>
                    <div class="list-editor">
                        <div class="list-editor-items" id="prop-dropdown-items-container">
                            ${itemsMarkup}
                        </div>
                        <button type="button" class="btn-secondary" id="btn-add-drop-item" style="padding:3px 8px; font-size:10px; align-self:flex-start; margin-top:4px;">+ 항목 추가</button>
                    </div>
                </div>
            `;
        } else if (ctrl.type === 'menu') {
            let itemsMarkup = '';
            if (ctrl.items) {
                ctrl.items.forEach((item, itemIdx) => {
                    itemsMarkup += `
                        <div class="list-item-row menu-item-row" data-idx="${itemIdx}" style="display: grid; grid-template-columns: 1fr 1fr 1.2fr 24px; gap: 4px; align-items: center; margin-bottom: 4px;">
                            <input type="text" class="prop-menu-item-id" placeholder="ID" value="${item.id}" style="font-size:10.5px; padding:2px;">
                            <input type="text" class="prop-menu-item-label" placeholder="텍스트" value="${item.label}" style="font-size:10.5px; padding:2px;">
                            <input type="text" class="prop-menu-item-action" placeholder="onAction" value="${item.onAction || ''}" style="font-size:10.5px; padding:2px;">
                            <button type="button" class="overlay-btn delete menu-item-del-btn" style="display:flex; width:22px; height:22px; align-items:center; justify-content:center; padding:0;">✖</button>
                        </div>
                    `;
                });
            }

            extraFields = `
                <div class="form-group" style="grid-column: span 2;">
                    <label>서브 메뉴 항목 관리 (ID / 레이블명 / 실행 매크로명)</label>
                    <div class="list-editor">
                        <div class="list-editor-items" id="prop-menu-items-container">
                            ${itemsMarkup}
                        </div>
                        <button type="button" class="btn-secondary" id="btn-add-menu-item" style="padding:3px 8px; font-size:10px; align-self:flex-start; margin-top:4px;">+ 서브 메뉴 추가</button>
                    </div>
                </div>
            `;
        }

        editorBody.innerHTML = `
            <div class="property-grid">
                <div class="form-group">
                    <label>컴포넌트 ID</label>
                    <input type="text" id="prop-ctrl-id" value="${ctrl.id || ''}">
                </div>
                <div class="form-group">
                    <label>표시 레이블 명</label>
                    <input type="text" id="prop-ctrl-label" value="${ctrl.label || ''}">
                </div>
                ${callbackMarkup}
                ${sizeMarkup}
                ${iconMarkup}
                ${layoutMarkup}
                ${extraFields}
                
                <!-- VBA Code Connector & Textarea Editor -->
                <div class="vba-connect-section" style="${ctrl.type === 'labelcontrol' ? 'display: none;' : ''}">
                    <div class="vba-connect-left">
                        <label>VBA 코드 프리셋 연결</label>
                        <select id="prop-ctrl-vba-preset" style="width: 100%;">
                            <option value="" ${!ctrl.vbaPresetId ? 'selected' : ''}>사용자 정의 (직접 작성)</option>
                            ${VBA_PRESETS.map(preset => `<option value="${preset.id}" ${ctrl.vbaPresetId === preset.id ? 'selected' : ''}>${preset.label}</option>`).join('')}
                        </select>
                        <div style="font-size: 10px; color: var(--text-secondary); margin-top: 4px; line-height: 1.35;">
                            프리셋을 선택하면 동작 매크로 코드가 오른쪽에 자동 연결됩니다.
                        </div>
                    </div>
                    <div class="vba-connect-right">
                        <label>연결된 VBA 실행 코드 편집</label>
                        <textarea id="prop-ctrl-vba-code" placeholder="Sub %ACTION_NAME%(control As IRibbonControl)..." spellcheck="false" style="width: 100%;">${ctrl.vbaCustomCode || ''}</textarea>
                    </div>
                </div>
            </div>
        `;

        if (['combobox', 'dropdown'].includes(ctrl.type)) {
            document.getElementById("btn-add-drop-item").addEventListener("click", () => {
                if (!ctrl.items) ctrl.items = [];
                const nextId = ctrl.items.length + 1;
                ctrl.items.push({ id: `${ctrl.id}_item${nextId}`, label: `옵션 ${nextId}` });
                updateRibbonUI();
            });

            document.querySelectorAll(".drop-item-del-btn").forEach(btn => {
                btn.addEventListener("click", (e) => {
                    const row = btn.closest(".list-item-row");
                    const idx = parseInt(row.getAttribute("data-idx"));
                    ctrl.items.splice(idx, 1);
                    updateRibbonUI();
                });
            });
        }

        if (ctrl.type === 'menu') {
            document.getElementById("btn-add-menu-item").addEventListener("click", () => {
                if (!ctrl.items) ctrl.items = [];
                const nextId = ctrl.items.length + 1;
                ctrl.items.push({ 
                    id: `${ctrl.id}_sub${nextId}`, 
                    label: `서브 명령 ${nextId}`, 
                    onAction: `${ctrl.id}_sub${nextId}_Click` 
                });
                updateRibbonUI();
            });

            document.querySelectorAll(".menu-item-del-btn").forEach(btn => {
                btn.addEventListener("click", (e) => {
                    const row = btn.closest(".menu-item-row");
                    const idx = parseInt(row.getAttribute("data-idx"));
                    ctrl.items.splice(idx, 1);
                    updateRibbonUI();
                });
            });
        }

        if (isSmallCtrl) {
            document.getElementById("prop-ctrl-startnewcolumn").addEventListener("change", (e) => {
                ctrl.startNewColumn = e.target.checked;
                updateRibbonUI();
            });
        }

        if (hasIcon) {
            document.getElementById("btn-open-icon-picker").addEventListener("click", () => {
                currentIconTargetControl = ctrl;
                openIconPicker();
            });
        }

        if (ctrl.type === 'box') {
            const boxStyleSelect = document.getElementById("prop-ctrl-boxstyle");
            if (boxStyleSelect) {
                boxStyleSelect.addEventListener("change", (e) => {
                    ctrl.boxStyle = e.target.value;
                    updateRibbonUI();
                });
            }
            const stackLimitSelect = document.getElementById("prop-ctrl-stacklimit");
            if (stackLimitSelect) {
                stackLimitSelect.addEventListener("change", (e) => {
                    ctrl.stackLimit = parseInt(e.target.value, 10);
                    updateRibbonUI();
                });
            }
        }
    }
}

// Live sync property fields
function bindPropertyFields() {
    const editorBody = document.getElementById("property-editor-fields");
    
    editorBody.addEventListener("input", (e) => {
        if (!selectedElement) return;
        const { type, path } = selectedElement;
        const targetId = e.target.id;

        if (type === 'tab') {
            const tab = ribbonState.tabs[path[0]];
            if (targetId === "prop-tab-id") tab.id = e.target.value;
            if (targetId === "prop-tab-label") tab.label = e.target.value;
            if (targetId === "prop-tab-isStandard") {
                tab.isStandardTab = e.target.value === 'true';
                tab.idMso = tab.isStandardTab ? "TabHome" : "";
                updateRibbonUI();
                return;
            }
            if (targetId === "prop-tab-idMso") tab.idMso = e.target.value;
            
        } else if (type === 'group') {
            const group = ribbonState.tabs[path[0]].groups[path[1]];
            if (targetId === "prop-group-id") group.id = e.target.value;
            if (targetId === "prop-group-label") group.label = e.target.value;

        } else if (type === 'control') {
            const ctrl = ribbonState.tabs[path[0]].groups[path[1]].controls[path[2]];
            if (targetId === "prop-ctrl-id") ctrl.id = e.target.value;
            if (targetId === "prop-ctrl-label") ctrl.label = e.target.value;
            if (targetId === "prop-ctrl-boxstyle") ctrl.boxStyle = e.target.value;
            if (targetId === "prop-ctrl-stacklimit") ctrl.stackLimit = parseInt(e.target.value, 10);
            if (targetId === "prop-ctrl-onaction") {
                const oldAction = ctrl.onAction || "";
                const newAction = e.target.value;
                ctrl.onAction = newAction;
                if (ctrl.vbaCustomCode && oldAction) {
                    try {
                        const regex = new RegExp('Sub\\s+' + oldAction + '\\b', 'g');
                        ctrl.vbaCustomCode = ctrl.vbaCustomCode.replace(regex, 'Sub ' + newAction);
                    } catch (err) {
                        console.error("VBA rename sync failed", err);
                    }
                }
                const codeArea = document.getElementById("prop-ctrl-vba-code");
                if (codeArea) {
                    codeArea.value = ctrl.vbaCustomCode || "";
                }
            }
            if (targetId === "prop-ctrl-vba-preset") {
                const presetId = e.target.value;
                ctrl.vbaPresetId = presetId;
                if (presetId) {
                    const preset = VBA_PRESETS.find(p => p.id === presetId);
                    if (preset) {
                        const actionName = ctrl.onAction || (ctrl.id + "_Click");
                        ctrl.vbaCustomCode = preset.code.replace(/%ACTION_NAME%/g, actionName);
                    }
                }
                const codeArea = document.getElementById("prop-ctrl-vba-code");
                if (codeArea) {
                    codeArea.value = ctrl.vbaCustomCode || "";
                }
                renderRibbonSimulator();
                compileCodes();
                return;
            }
            if (targetId === "prop-ctrl-vba-code") {
                ctrl.vbaCustomCode = e.target.value;
                ctrl.vbaPresetId = ""; // Custom
                const presetSelect = document.getElementById("prop-ctrl-vba-preset");
                if (presetSelect) {
                    presetSelect.value = "";
                }
                compileCodes();
                return;
            }
            if (targetId === "prop-ctrl-text") ctrl.text = e.target.value;
            
            if (targetId === "prop-ctrl-checked") {
                ctrl.checked = e.target.value === 'true';
                if (ctrl.type === 'togglebutton' && ctrl.checked) {
                    const activeTabIdx = path[0];
                    const activeGrpIdx = path[1];
                    const group = ribbonState.tabs[activeTabIdx].groups[activeGrpIdx];
                    const targetGroup = ctrl.toggleGroup || "";
                    group.controls.forEach(c => {
                        if (c !== ctrl && c.type === 'togglebutton' && (c.toggleGroup || "") === targetGroup) {
                            c.checked = false;
                        }
                    });
                }
            }

            if (targetId === "prop-ctrl-togglegroup") {
                ctrl.toggleGroup = e.target.value;
                if (ctrl.type === 'togglebutton' && ctrl.checked) {
                    const activeTabIdx = path[0];
                    const activeGrpIdx = path[1];
                    const group = ribbonState.tabs[activeTabIdx].groups[activeGrpIdx];
                    const targetGroup = ctrl.toggleGroup || "";
                    group.controls.forEach(c => {
                        if (c !== ctrl && c.type === 'togglebutton' && (c.toggleGroup || "") === targetGroup) {
                            c.checked = false;
                        }
                    });
                }
            }

            if (e.target.name === "btn-size") {
                ctrl.size = e.target.value;
                updateRibbonUI();
                return;
            }

            if (e.target.classList.contains("prop-drop-item-id") || e.target.classList.contains("prop-drop-item-label")) {
                const row = e.target.closest(".list-item-row");
                const itemIdx = parseInt(row.getAttribute("data-idx"));
                
                if (e.target.classList.contains("prop-drop-item-id")) {
                    ctrl.items[itemIdx].id = e.target.value;
                } else {
                    ctrl.items[itemIdx].label = e.target.value;
                }
            }

            if (e.target.classList.contains("prop-menu-item-id") || e.target.classList.contains("prop-menu-item-label") || e.target.classList.contains("prop-menu-item-action")) {
                const row = e.target.closest(".menu-item-row");
                const itemIdx = parseInt(row.getAttribute("data-idx"));
                
                if (e.target.classList.contains("prop-menu-item-id")) {
                    ctrl.items[itemIdx].id = e.target.value;
                } else if (e.target.classList.contains("prop-menu-item-label")) {
                    ctrl.items[itemIdx].label = e.target.value;
                } else if (e.target.classList.contains("prop-menu-item-action")) {
                    ctrl.items[itemIdx].onAction = e.target.value;
                }
            }
        }

        renderRibbonSimulator();
        compileCodes();
    });
}

// --- 8. Real-time Double Code Compiler ---
function compileCodes() {
    const xml = compileCustomUiXml();
    document.getElementById("code-xml-box").value = xml;

    const vba = compileVbaCallbacks();
    document.getElementById("code-vba-box").value = vba;
}

function compileSingleControlXml(ctrl, indent) {
    let xml = "";
    const sizeAttr = ctrl.size === 'large' ? ' size="large"' : '';
    
    // Treat local paths and emojis as a standard fallback imageMso in the final XML
    let imageAttr = '';
    if (ctrl.imageMso) {
        const isEmoji = ctrl.imageMso.length <= 4 && /\p{Emoji}/u.test(ctrl.imageMso);
        const isLocalFile = ctrl.imageMso.startsWith("icons/");
        
        if (isEmoji || isLocalFile) {
            // Extract file name or use standard HappyFace fallback
            if (isLocalFile) {
                const filename = ctrl.imageMso.split('/').pop().replace('.svg', '').replace(/ /g, '');
                imageAttr = ` imageMso="${filename}"`; // Excel matches standard names if loaded locally
            } else {
                imageAttr = ` imageMso="HappyFace"`;
            }
        } else {
            imageAttr = ` imageMso="${ctrl.imageMso}"`;
        }
    }

    if (ctrl.type === 'separator') {
        xml += `${indent}<separator id="${ctrl.id}" />\n`;
    } else if (ctrl.type === 'checkbox') {
        xml += `${indent}<checkBox id="${ctrl.id}" label="${escapeXml(ctrl.label)}" onAction="${ctrl.onAction}" />\n`;
    } else if (ctrl.type === 'editbox') {
        xml += `${indent}<editBox id="${ctrl.id}" label="${escapeXml(ctrl.label)}" onChange="${ctrl.onAction}" />\n`;
    } else if (ctrl.type === 'combobox') {
        xml += `${indent}<comboBox id="${ctrl.id}" label="${escapeXml(ctrl.label)}" onChange="${ctrl.onAction}">\n`;
        if (ctrl.items) {
            ctrl.items.forEach(item => {
                xml += `${indent}  <item id="${item.id}" label="${escapeXml(item.label)}" />\n`;
            });
        }
        xml += `${indent}</comboBox>\n`;
    } else if (ctrl.type === 'dropdown') {
        xml += `${indent}<dropDown id="${ctrl.id}" label="${escapeXml(ctrl.label)}" onAction="${ctrl.onAction}">\n`;
        if (ctrl.items) {
            ctrl.items.forEach(item => {
                xml += `${indent}  <item id="${item.id}" label="${escapeXml(item.label)}" />\n`;
            });
        }
        xml += `${indent}</dropDown>\n`;
    } else if (ctrl.type === 'labelcontrol') {
        xml += `${indent}<labelControl id="${ctrl.id}" label="${escapeXml(ctrl.label)}" />\n`;
    } else if (ctrl.type === 'menu') {
        xml += `${indent}<menu id="${ctrl.id}" label="${escapeXml(ctrl.label)}"${sizeAttr}${imageAttr}>\n`;
        if (ctrl.items) {
            ctrl.items.forEach(item => {
                xml += `${indent}  <button id="${item.id}" label="${escapeXml(item.label)}" onAction="${item.onAction || ''}" />\n`;
            });
        }
        xml += `${indent}</menu>\n`;
    } else if (ctrl.type === 'togglebutton') {
        xml += `${indent}<toggleButton id="${ctrl.id}" label="${escapeXml(ctrl.label)}"${sizeAttr} onAction="${ctrl.onAction}" />\n`;
    } else {
        // standard button
        xml += `${indent}<button id="${ctrl.id}" label="${escapeXml(ctrl.label)}"${sizeAttr}${imageAttr} onAction="${ctrl.onAction}" />\n`;
    }
    return xml;
}

function compileCustomUiXml() {
    let xml = `<!-- Office 2010+ 호환 CustomUI XML 스키마 -->\n`;
    xml += `<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">\n`;
    xml += `  <ribbon>\n`;
    xml += `    <tabs>\n`;

    ribbonState.tabs.forEach(tab => {
        if (tab.isStandardTab) {
            xml += `      <!-- 표준 탭 확장 주입 -->\n`;
            xml += `      <tab idMso="${tab.idMso}">\n`;
        } else {
            let posAttr = "";
            if (tab.positionType === 'before' && tab.positionTarget) {
                posAttr = ` insertBeforeMso="${tab.positionTarget}"`;
            } else if (tab.positionType === 'after' && tab.positionTarget) {
                posAttr = ` insertAfterMso="${tab.positionTarget}"`;
            }
            xml += `      <tab id="${tab.id}" label="${escapeXml(tab.label)}"${posAttr}>\n`;
        }

        if (tab.groups) {
            tab.groups.forEach(group => {
                xml += `        <group id="${group.id}" label="${escapeXml(group.label)}">\n`;
                
                if (group.controls && group.controls.length > 0) {
                    let i = 0;
                    while (i < group.controls.length) {
                        const ctrl = group.controls[i];
                        
                        if (ctrl.type === 'box') {
                            const stackLimit = ctrl.stackLimit || 3;
                            const children = [];
                            let j = i + 1;
                            while (j < group.controls.length && children.length < stackLimit) {
                                const candidate = group.controls[j];
                                if (candidate.type === 'separator' || candidate.type === 'box' || candidate.type === 'buttongroup') {
                                    break;
                                }
                                children.push(candidate);
                                j++;
                            }
                            
                            xml += `          <box id="${ctrl.id}" boxStyle="${ctrl.boxStyle || 'vertical'}">\n`;
                            children.forEach(child => {
                                xml += compileSingleControlXml(child, "            ");
                            });
                            xml += `          </box>\n`;
                            
                            i = j;
                        } else {
                            xml += compileSingleControlXml(ctrl, "          ");
                            i++;
                        }
                    }
                }
                
                xml += `        </group>\n`;
            });
        }
        xml += `      </tab>\n`;
    });

    xml += `    </tabs>\n`;
    xml += `  </ribbon>\n`;
    xml += `</customUI>`;
    return xml;
}


function compileVbaCallbacks() {
    let vba = `' =======================================================\n`;
    vba += `'  EXCEL RIBBON CALLBACK VBA CODE (자동 생성된 소스코드)\n`;
    vba += `'  [가이드] 다운로드한 엑셀에 Alt+F11로 모듈 삽입 후 아래 전체 복사 붙여넣기\n`;
    vba += `' =======================================================\n\n`;

    let generatedActions = new Set();

    ribbonState.tabs.forEach(tab => {
        if (tab.groups) {
            tab.groups.forEach(group => {
                if (group.controls) {
                    group.controls.forEach(ctrl => {
                        if (ctrl.type === 'menu') {
                            if (ctrl.items) {
                                ctrl.items.forEach(sub => {
                                    if (!sub.onAction || generatedActions.has(sub.onAction)) return;
                                    generatedActions.add(sub.onAction);
                                    vba += `' ${ctrl.label} > ${sub.label} 버튼 클릭 콜백\n`;
                                    vba += `Sub ${sub.onAction}(control As IRibbonControl)\n`;
                                    vba += `    MsgBox "${sub.label} 기능을 실행합니다.", vbInformation, "엑셀 리본 빌더"\n`;
                                    vba += `End Sub\n\n`;
                                });
                            }
                            return;
                        }

                        if (ctrl.type === 'separator' || ctrl.type === 'labelcontrol' || ctrl.type === 'box' || ctrl.type === 'buttongroup' || !ctrl.onAction) return;
                        if (generatedActions.has(ctrl.onAction)) return;
                        generatedActions.add(ctrl.onAction);

                        if (ctrl.vbaCustomCode) {
                            vba += ctrl.vbaCustomCode + "\n\n";
                        } else {
                            if (ctrl.type === 'checkbox') {
                                vba += `' ${ctrl.label} (체크박스 변경 시 실행)\n`;
                                vba += `Sub ${ctrl.onAction}(control As IRibbonControl, pressed As Boolean)\n`;
                                vba += `    MsgBox "체크 상태 변경: " & pressed, vbInformation, "엑셀 리본 빌더"\n`;
                                vba += `End Sub\n\n`;
                            } else if (ctrl.type === 'editbox') {
                                vba += `' ${ctrl.label} (텍스트 값 변경 시 실행)\n`;
                                vba += `Sub ${ctrl.onAction}(control As IRibbonControl, text As String)\n`;
                                vba += `    MsgBox "입력된 텍스트: " & text, vbInformation, "엑셀 리본 빌더"\n`;
                                vba += `End Sub\n\n`;
                            } else if (['combobox', 'dropdown'].includes(ctrl.type)) {
                                vba += `' ${ctrl.label} (선택값 변경 시 실행)\n`;
                                vba += `Sub ${ctrl.onAction}(control As IRibbonControl, textOrId As String, Optional index As Integer)\n`;
                                vba += `    MsgBox "선택된 값: " & textOrId, vbInformation, "엑셀 리본 빌더"\n`;
                                vba += `End Sub\n\n`;
                            } else if (ctrl.type === 'togglebutton') {
                                vba += `' ${ctrl.label} (토글 상태 변경 시 실행)\n`;
                                vba += `Sub ${ctrl.onAction}(control As IRibbonControl, pressed As Boolean)\n`;
                                vba += `    MsgBox "토글 상태: " & pressed, vbInformation, "엑셀 리본 빌더"\n`;
                                vba += `End Sub\n\n`;
                            } else {
                                vba += `' ${ctrl.label} 버튼 클릭 콜백\n`;
                                vba += `Sub ${ctrl.onAction}(control As IRibbonControl)\n`;
                                vba += `    MsgBox "${ctrl.label} 기능을 실행합니다.", vbInformation, "엑셀 리본 빌더"\n`;
                                vba += `End Sub\n\n`;
                            }
                        }
                    });
                }
            });
        }
    });

    return vba;
}

function escapeXml(unsafe) {
    if (!unsafe) return "";
    return unsafe.replace(/[<>&'"]/g, function (c) {
        switch (c) {
            case '<': return '&lt;';
            case '>': return '&gt;';
            case '&': return '&amp;';
            case '\'': return '&apos;';
            case '"': return '&quot;';
        }
    });
}

// --- 9. Modal Searchable Icon Picker with Dynamic Folder & Emoji sub-tabs ---
function buildIconPickerGrid() {
    const grid = document.getElementById("modal-icons-grid");
    const subTabsBar = document.getElementById("modal-sub-tabs");
    
    grid.innerHTML = "";
    subTabsBar.innerHTML = "";

    const query = document.getElementById("modal-icon-search").value.toLowerCase();

    // 1. Excel Dynamic local Folder scan tab
    if (currentIconSourceTab === 'mso') {
        if (typeof ICONS_MAP === 'undefined') {
            grid.innerHTML = `<div style="padding:20px; font-style:italic; color:#888;">icons_map.js가 유실되었습니다. 파워셸을 실행해 주세요.</div>`;
            return;
        }

        const level1Keys = Object.keys(ICONS_MAP);
        if (level1Keys.length === 0) return;

        // Monolithic category folder check
        if (m365SelectedLevel1 && !level1Keys.includes(m365SelectedLevel1)) {
            m365SelectedLevel1 = level1Keys[0];
        }

        // Render Level 1 dynamic sub-tabs
        level1Keys.forEach(l1Key => {
            const btn = document.createElement("button");
            btn.className = "sub-tab-btn";
            if (l1Key === m365SelectedLevel1) btn.classList.add("active");
            btn.textContent = l1Key;
            
            btn.addEventListener("click", () => {
                m365SelectedLevel1 = l1Key;
                if (Array.isArray(ICONS_MAP[m365SelectedLevel1])) {
                    m365SelectedLevel2 = "";
                } else {
                    const subKeys = Object.keys(ICONS_MAP[m365SelectedLevel1]);
                    m365SelectedLevel2 = subKeys.length > 0 ? subKeys[0] : "";
                }
                buildIconPickerGrid();
            });
            subTabsBar.appendChild(btn);
        });

        // Render Level 2 dynamic sub-tabs as a second row!
        if (!Array.isArray(ICONS_MAP[m365SelectedLevel1])) {
            const l2Container = document.createElement("div");
            l2Container.className = "sub-tabs-level2";
            
            const l2Keys = Object.keys(ICONS_MAP[m365SelectedLevel1] || {});
            l2Keys.forEach(l2Key => {
                const btn = document.createElement("button");
                btn.className = "sub-tab-btn-level2";
                if (l2Key === m365SelectedLevel2) btn.classList.add("active");
                btn.textContent = l2Key;
                
                btn.addEventListener("click", () => {
                    m365SelectedLevel2 = l2Key;
                    buildIconPickerGrid();
                });
                l2Container.appendChild(btn);
            });
            subTabsBar.appendChild(l2Container);
        }

        // Display SVG icons list inside grid matching active folder
        let filesList = [];
        if (Array.isArray(ICONS_MAP[m365SelectedLevel1])) {
            filesList = ICONS_MAP[m365SelectedLevel1] || [];
        } else {
            filesList = ICONS_MAP[m365SelectedLevel1]?.[m365SelectedLevel2] || [];
        }
        
        // Search filter (searches globally across all folders if search query exists!)
        if (query) {
            // Global directory search
            level1Keys.forEach(l1 => {
                if (Array.isArray(ICONS_MAP[l1])) {
                    const files = ICONS_MAP[l1] || [];
                    files.forEach(fileName => {
                        const displayName = fileName.replace(".svg", "");
                        if (displayName.toLowerCase().includes(query) || fileName.toLowerCase().includes(query)) {
                            const fullRelPath = `icons/2024-microsoft-365-content-icons/${l1}/${fileName}`;
                            const card = createIconCard(fileName, displayName, fullRelPath);
                            grid.appendChild(card);
                        }
                    });
                } else {
                    const subMap = ICONS_MAP[l1] || {};
                    Object.keys(subMap).forEach(l2 => {
                        const files = subMap[l2] || [];
                        files.forEach(fileName => {
                            const displayName = fileName.replace(".svg", "");
                            if (displayName.toLowerCase().includes(query) || fileName.toLowerCase().includes(query)) {
                                const fullRelPath = `icons/2024-microsoft-365-content-icons/${l1}/${l2}/${fileName}`;
                                const card = createIconCard(fileName, displayName, fullRelPath);
                                grid.appendChild(card);
                            }
                        });
                    });
                }
            });
        } else {
            // Normal browsing of selected folder
            filesList.forEach(fileName => {
                const displayName = fileName.replace(".svg", "");
                let fullRelPath = "";
                if (Array.isArray(ICONS_MAP[m365SelectedLevel1])) {
                    fullRelPath = `icons/2024-microsoft-365-content-icons/${m365SelectedLevel1}/${fileName}`;
                } else {
                    fullRelPath = `icons/2024-microsoft-365-content-icons/${m365SelectedLevel1}/${m365SelectedLevel2}/${fileName}`;
                }
                const card = createIconCard(fileName, displayName, fullRelPath);
                grid.appendChild(card);
            });
        }
        
    } else if (currentIconSourceTab === 'emoji') {
        // 2. Classified Emoji Tab
        if (typeof EMOJI_GROUPS === 'undefined') return;

        // Render Emoji Group sub-tabs
        Object.keys(EMOJI_GROUPS).forEach(grpKey => {
            const btn = document.createElement("button");
            btn.className = "sub-tab-btn";
            if (grpKey === emojiSelectedGroup) btn.classList.add("active");
            btn.textContent = EMOJI_GROUPS[grpKey].name;
            
            btn.addEventListener("click", () => {
                emojiSelectedGroup = grpKey;
                buildIconPickerGrid();
            });
            subTabsBar.appendChild(btn);
        });

        // Search filter or standard list
        if (query) {
            // Global Emoji search across all groups
            Object.keys(EMOJI_GROUPS).forEach(grpKey => {
                EMOJI_GROUPS[grpKey].items.forEach(emoji => {
                    if (emoji.includes(query) || grpKey.includes(query)) {
                        const card = createIconCard(emoji, emoji, emoji);
                        grid.appendChild(card);
                    }
                });
            });
        } else {
            // Display active group emojis
            const items = EMOJI_GROUPS[emojiSelectedGroup]?.items || [];
            items.forEach(emoji => {
                const card = createIconCard(emoji, emoji, emoji);
                grid.appendChild(card);
            });
        }
    }
}

function createIconCard(id, label, iconKey) {
    const card = document.createElement("div");
    card.className = "icon-select-card";

    const visual = document.createElement("div");
    visual.className = "icon-visual";
    visual.innerHTML = getIconHtml(iconKey, 24);
    card.appendChild(visual);

    const span = document.createElement("span");
    span.textContent = label;
    card.appendChild(span);

    // Selected click action
    card.addEventListener("click", () => {
        if (currentIconTargetControl) {
            currentIconTargetControl.imageMso = iconKey;
            updateRibbonUI();
            closeIconPicker();
        }
    });

    return card;
}

function openIconPicker() {
    document.getElementById("icon-picker-modal").classList.add("active");
    document.getElementById("modal-icon-search").value = "";
    document.getElementById("modal-icon-search").focus();
    
    buildIconPickerGrid();
}

function closeIconPicker() {
    document.getElementById("icon-picker-modal").classList.remove("active");
    currentIconTargetControl = null;
}

function filterIconsList() {
    buildIconPickerGrid();
}

// --- 10. Copy and File Exporters ---
function copyCode(boxId, btnId) {
    const code = document.getElementById(boxId).value;
    navigator.clipboard.writeText(code).then(() => {
        const btn = document.getElementById(btnId);
        const originalText = btn.innerHTML;
        btn.innerHTML = `✓ 복사 완료!`;
        btn.style.color = "var(--excel-green-primary)";
        btn.style.borderColor = "var(--excel-green-primary)";

        setTimeout(() => {
            btn.innerHTML = originalText;
            btn.style.color = "";
            btn.style.borderColor = "";
        }, 1500);
    }).catch(err => {
        console.error("복사 실패:", err);
        alert("복사 실패했습니다. 텍스트 영역을 직접 복사해 주세요.");
    });
}

async function exportExcelFile(type = 'xlsm') {
    const xml = compileCustomUiXml();
    
    try {
        const btn = document.getElementById(type === 'xlsm' ? "btn-export-xlsm" : "btn-export-xlam");
        const origText = btn ? btn.innerHTML : "";
        if (btn) {
            btn.innerHTML = `⚙ 빌드 패키징 중...`;
            btn.disabled = true;
        }

        const blob = await ExcelZipGenerator.generate(xml, type);
        
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = type === 'xlsm' ? "CustomRibbon_Workbook.xlsm" : "CustomRibbon_AddIn.xlam";
        document.body.appendChild(a);
        a.click();
        
        setTimeout(() => {
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            if (btn) {
                btn.innerHTML = `✓ 내보내기 완료!`;
                btn.disabled = false;
            }
            
            setTimeout(() => {
                if (btn) btn.innerHTML = origText;
            }, 1800);
        }, 100);

    } catch (err) {
        console.error(err);
        alert("바이너리 빌드 중 오류가 발생했습니다: " + err.message);
    }
}
