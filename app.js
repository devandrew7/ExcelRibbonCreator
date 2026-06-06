/**
 * Excel Custom Ribbon Creator - Extended Core Application Logic Script (v4)
 * Implements M365 local directory dynamic folder mapping, 5-group emoji sub-tabs,
 * unified image/emoji renderer, tab reordering, and light theme state binding.
 */

// --- 1. Global State Management ---
let ribbonState = {
    tabs: [],
    backstage: null,
    commands: []
};

let isBackstageOpen = false;
let activeBackstageTabId = null;

let selectedElement = null; // { type: 'tab'|'group'|'control'|'separator', path: [tabIdx, grpIdx, ctrlIdx] }

// --- 1.2 Office RibbonX Editor States & Configuration ---
let activeSchemaVersion = 'customUI14'; // 'customUI14' or 'customUI'
let activeExplorerNode = 'customui_file'; // 'customui_file', 'rels', 'workbook', etc.

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
                            imageMso: "FileOpen",
                            onAction: "btn_MergeFiles_Click",
                            enabled: true,
                            visible: true
                        },
                        {
                            id: "btn_CleanDuplicates",
                            type: "button",
                            label: "중복값 제거",
                            size: "normal",
                            imageMso: "Delete",
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
                            imageMso: "Mail",
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
                            imageMso: "Help", // Default Mso icon
                            onAction: "btn_CopilotWeb_Click",
                            enabled: true,
                            visible: true
                        },
                        {
                            id: "split_BackupMenu",
                            type: "splitbutton",
                            label: "보안 작업",
                            size: "large",
                            imageMso: "SheetProtect", // Local SVG File Demo
                            onAction: "split_BackupMenu_Click",
                            enabled: true,
                            visible: true,
                            controls: [
                                { id: "btn_InnerLock", type: "button", label: "시트 비밀번호 잠금", imageMso: "SheetProtect", onAction: "btn_InnerLock_Click" },
                                { id: "btn_InnerUnlock", type: "button", label: "보호 전체 해제", imageMso: "SheetUnprotect", onAction: "btn_InnerUnlock_Click" }
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
    if (typeof OFFICE_MSO_ICONS !== 'undefined') {
        const level1Keys = Object.keys(OFFICE_MSO_ICONS);
        if (level1Keys.length > 0) {
            m365SelectedLevel1 = level1Keys[0];
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
    // Schema version selector change listener
    const schemaSelector = document.getElementById("schema-version-selector");
    if (schemaSelector) {
        schemaSelector.addEventListener("change", (e) => {
            activeSchemaVersion = e.target.value;
            compileCodes();
            renderDocumentExplorer();
            if (activeExplorerNode === 'customui_file') {
                handleExplorerNodeSelection('customui_file');
            }
        });
    }

    // RibbonX Template selector change listener
    const templateSelector = document.getElementById("template-preset-selector");
    if (templateSelector) {
        templateSelector.addEventListener("change", (e) => {
            const templateVal = e.target.value;
            if (!templateVal) return;

            if (templateVal === 'demo_workplace') {
                loadState(JSON.parse(JSON.stringify(DEMO_WORKPLACE)));
                activeSchemaVersion = 'customUI14';
                if (schemaSelector) schemaSelector.value = 'customUI14';
                showValidationBanner(true, "데모 1: 자동화 종합 세트가 로드되었습니다.");
            } else if (templateVal === 'demo_intrusion') {
                loadState(JSON.parse(JSON.stringify(DEMO_HOME_INTRUSION)));
                activeSchemaVersion = 'customUI14';
                if (schemaSelector) schemaSelector.value = 'customUI14';
                showValidationBanner(true, "데모 2: 홈 탭 주입 확장이 로드되었습니다.");
            } else if (RIBBONX_TEMPLATES[templateVal]) {
                const template = RIBBONX_TEMPLATES[templateVal];
                
                if (templateVal === 'custom_outspace' && activeSchemaVersion === 'customUI') {
                    alert("Backstage(파일 메뉴) 구성 요소는 Office 2010+ (customUI14)에서만 지원됩니다. 스키마 버전을 자동으로 변경합니다.");
                    activeSchemaVersion = 'customUI14';
                    if (schemaSelector) schemaSelector.value = 'customUI14';
                }

                let xmlText = template.xml;
                if (activeSchemaVersion === 'customUI') {
                    xmlText = xmlText.replace("http://schemas.microsoft.com/office/2009/07/customui", "http://schemas.microsoft.com/office/2006/01/customui");
                }
                
                document.getElementById("code-xml-box").value = xmlText;
                
                const result = xmlToState(xmlText);
                if (!result.error) {
                    loadState(result);
                    
                    let msg = `성공: '${e.target.selectedOptions[0].textContent}' 템플릿이 로드되었습니다.`;
                    if (result.hasNonSimulatable) {
                        msg += " (시뮬레이터 미지원 XML 태그는 텍스트에 유지됩니다)";
                    }
                    showValidationBanner(true, msg);
                } else {
                    showValidationBanner(false, "템플릿 파싱 오류: " + result.error);
                }
            }
            templateSelector.value = "";
        });
    }

    document.getElementById("btn-clear-all").addEventListener("click", () => {
        loadState({ tabs: [] });
        const banner = document.getElementById("xml-validation-banner");
        if (banner) banner.style.display = "none";
    });

    // Palette triggers (10 controls)
    document.getElementById("add-tab")?.addEventListener("click", addNewTab);
    document.getElementById("add-group")?.addEventListener("click", addNewGroup);
    document.getElementById("add-button")?.addEventListener("click", () => addNewControl('button'));
    document.getElementById("add-togglebutton")?.addEventListener("click", () => addNewControl('togglebutton'));
    document.getElementById("add-checkbox")?.addEventListener("click", () => addNewControl('checkbox'));
    document.getElementById("add-editbox")?.addEventListener("click", () => addNewControl('editbox'));
    document.getElementById("add-combobox")?.addEventListener("click", () => addNewControl('combobox'));
    document.getElementById("add-dropdown")?.addEventListener("click", () => addNewControl('dropdown'));
    document.getElementById("add-menu")?.addEventListener("click", () => addNewControl('menu'));
    document.getElementById("add-labelcontrol")?.addEventListener("click", () => addNewControl('labelcontrol'));
    document.getElementById("add-separator")?.addEventListener("click", () => addNewControl('separator'));
    document.getElementById("add-box")?.addEventListener("click", () => addNewControl('box'));
    document.getElementById("add-gallery")?.addEventListener("click", () => addNewControl('gallery'));
    document.getElementById("add-buttongroup")?.addEventListener("click", () => addNewControl('buttongroup'));
    document.getElementById("add-splitbutton")?.addEventListener("click", () => addNewControl('splitbutton'));
    document.getElementById("add-dynamicmenu")?.addEventListener("click", () => addNewControl('dynamicmenu'));

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

    // Direct manual icon apply in picker modal
    const applyManualBtn = document.getElementById("btn-apply-manual-icon");
    if (applyManualBtn) {
        applyManualBtn.addEventListener("click", () => {
            const manualInput = document.getElementById("modal-manual-icon-input");
            const iconKey = manualInput.value.trim();
            if (iconKey && currentIconTargetControl) {
                currentIconTargetControl.imageMso = iconKey;
                updateRibbonUI();
                closeIconPicker();
            }
        });
    }

    const manualInputEl = document.getElementById("modal-manual-icon-input");
    if (manualInputEl) {
        manualInputEl.addEventListener("keydown", (e) => {
            if (e.key === "Enter" && applyManualBtn) {
                applyManualBtn.click();
            }
        });
    }

    // Sidebar tab buttons switching
    const tabBtnPalette = document.getElementById("tab-btn-palette");
    const tabBtnExplorer = document.getElementById("tab-btn-explorer");
    const contentPalette = document.getElementById("sidebar-palette-content");
    const contentExplorer = document.getElementById("sidebar-explorer-content");

    if (tabBtnPalette && tabBtnExplorer && contentPalette && contentExplorer) {
        tabBtnPalette.addEventListener("click", () => {
            tabBtnPalette.classList.add("active");
            tabBtnPalette.style.background = "#fdfdfd";
            tabBtnPalette.style.color = "var(--excel-green-primary)";
            tabBtnPalette.style.fontWeight = "bold";
            tabBtnPalette.style.borderColor = "var(--border-color)";

            tabBtnExplorer.classList.remove("active");
            tabBtnExplorer.style.background = "none";
            tabBtnExplorer.style.color = "var(--text-secondary)";
            tabBtnExplorer.style.fontWeight = "normal";
            tabBtnExplorer.style.borderColor = "transparent";

            contentPalette.style.display = "flex";
            contentExplorer.style.display = "none";
        });

        tabBtnExplorer.addEventListener("click", () => {
            tabBtnExplorer.classList.add("active");
            tabBtnExplorer.style.background = "#fdfdfd";
            tabBtnExplorer.style.color = "var(--excel-green-primary)";
            tabBtnExplorer.style.fontWeight = "bold";
            tabBtnExplorer.style.borderColor = "var(--border-color)";

            tabBtnPalette.classList.remove("active");
            tabBtnPalette.style.background = "none";
            tabBtnPalette.style.color = "var(--text-secondary)";
            tabBtnPalette.style.fontWeight = "normal";
            tabBtnPalette.style.borderColor = "transparent";

            contentPalette.style.display = "none";
            contentExplorer.style.display = "flex";
            
            renderDocumentExplorer();
        });
    }

    // XML Editor action bindings (Validation, Sync Apply, Close banner)
    document.getElementById("btn-close-validation-banner")?.addEventListener("click", () => {
        const banner = document.getElementById("xml-validation-banner");
        if (banner) banner.style.display = "none";
    });

    document.getElementById("btn-validate-xml")?.addEventListener("click", () => {
        const xmlText = document.getElementById("code-xml-box").value;
        const result = validateXmlText(xmlText);
        showValidationBanner(result.valid, result.message);
    });

    document.getElementById("btn-apply-xml")?.addEventListener("click", () => {
        const xmlText = document.getElementById("code-xml-box").value;
        const result = xmlToState(xmlText);
        if (result.error) {
            showValidationBanner(false, result.error);
        } else {
            activeSchemaVersion = result.schema || 'customUI14';
            const sSelector = document.getElementById("schema-version-selector");
            if (sSelector) sSelector.value = activeSchemaVersion;
            
            loadState(result);
            
            let msg = "시뮬레이터에 XML 데이터가 성공적으로 반영되었습니다!";
            if (result.hasNonSimulatable) {
                msg += "\n⚠️ 주의: 시뮬레이터에 표시되지 않는 요소가 포함되어 있으며, XML에는 정상 유지됩니다.";
            }
            showValidationBanner(true, msg);
        }
    });

    document.addEventListener("click", () => {
        document.querySelectorAll(".gallery-sim-dropdown, .splitbutton-sim-dropdown, .menu-dropdown-sim").forEach(d => d.style.display = 'none');
    });
}

function initializeControlVbaProperties(ctrl) {
    if (ctrl.type === 'separator' || ctrl.type === 'labelcontrol' || ctrl.type === 'box' || ctrl.type === 'buttongroup') return;
    
    if (ctrl.type === 'dynamicmenu') {
        if (!ctrl.getContent) ctrl.getContent = ctrl.id + "_GetContent";
    } else if (!ctrl.onAction) {
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
        } else if (ctrl.type === 'dynamicmenu') {
            defaultCode = `Sub ${ctrl.getContent}(control As IRibbonControl, ByRef content)\n` +
                          `    ' 동적으로 표시할 메뉴 XML을 문자열로 정의합니다.\n` +
                          `    Dim xml As String\n` +
                          `    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & _\n` +
                          `          "  <button id=""dynBtn1"" label=""동적 버튼 1"" imageMso=""HappyFace"" onAction=""DynButton_Click"" />" & _\n` +
                          `          "  <button id=""dynBtn2"" label=""동적 버튼 2"" imageMso=""Calendar"" onAction=""DynButton_Click"" />" & _\n` +
                          `          "</menu>"\n` +
                          `    content = xml\n` +
                          `End Sub\n\n` +
                          `Sub DynButton_Click(control As IRibbonControl)\n` +
                          `    MsgBox "동적 버튼 클릭됨: " & control.Id, vbInformation, "엑셀 리본 빌더"\n` +
                          `End Sub`;
        } else {
            defaultCode = `Sub ${ctrl.onAction}(control As IRibbonControl)\n    MsgBox "${ctrl.label || '기능'}을 실행합니다.", vbInformation, "엑셀 리본 빌더"\nEnd Sub`;
        }
        ctrl.vbaCustomCode = defaultCode;
        ctrl.vbaPresetId = ""; // default as Custom
    }
}

function loadState(state) {
    ribbonState = {
        tabs: state.tabs || [],
        backstage: state.backstage || null,
        commands: state.commands || []
    };
    selectedElement = null;
    isBackstageOpen = false;
    activeBackstageTabId = null;
    
    // Auto initialize VBA properties for loaded demo presets
    if (ribbonState.tabs) {
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
    }

    if (ribbonState.tabs && ribbonState.tabs.length > 0) {
        selectedElement = { type: 'tab', path: [0] };
    }
    updateRibbonUI();
}

function updateRibbonUI() {
    if (selectedElement) {
        const { type, path } = selectedElement;
        if (type === 'tab') {
            if (!ribbonState.tabs || path[0] >= ribbonState.tabs.length) {
                selectedElement = null;
            }
        } else if (type === 'group') {
            const tab = ribbonState.tabs[path[0]];
            if (!tab || !tab.groups || path[1] >= tab.groups.length) {
                selectedElement = null;
            }
        } else if (type === 'control') {
            const tab = ribbonState.tabs[path[0]];
            const group = tab && tab.groups ? tab.groups[path[1]] : null;
            if (!group || !group.controls || path[2] >= group.controls.length) {
                selectedElement = null;
            }
        }
    }
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
        return `<img src="${imageMsoValue}" class="mso-icon-img" style="width:${size}px; height:${size}px;" title="${imageMsoValue}">`;
    }

    // 2. Normal SVG Embedded Fallbacks (or custom user-typed values)
    if (window.getIconSvg) {
        const svg = window.getIconSvg(imageMsoValue, size);
        // If it's one of our curated SVGs, render it directly
        if (svg && svg.indexOf('Mso') === -1) {
            return svg.replace('<svg ', `<svg title="${imageMsoValue}" `);
        }
    }

    // Otherwise, we load it from the extracted folder, with a fallback to the generic circle SVG
    const fallback = window.getIconSvg ? window.getIconSvg(imageMsoValue, size) : `<span>★</span>`;
    const escapedFallback = fallback.replace(/"/g, '&quot;').replace(/'/g, '&#39;');

    return `<img src="imageMSO/${imageMsoValue}.png" onerror="this.outerHTML='${escapedFallback}'" class="mso-icon-img" style="width:${size}px; height:${size}px;" title="${imageMsoValue}">`;
}

// --- 5. Interactive Simulated Ribbon Renderer ---
// --- 5. Interactive Simulated Ribbon Renderer ---
function renderSingleControl(ctrl, activeTabIdx, grpIdx, ctrlIdx, totalControlsCount) {
    if (ctrl.type === 'separator') {
        const sepEl = document.createElement("div");
        sepEl.className = "ribbon-separator-sim";
        
        if (ctrl.isMso) {
            sepEl.addEventListener("click", (e) => {
                e.stopPropagation();
                showValidationBanner(true, `내장 구분선 - 이 컨트롤은 읽기 전용입니다.`);
            });
            return sepEl;
        }

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
    const isCtrlSelected = !ctrl.isMso && selectedElement && selectedElement.type === 'control'
        && selectedElement.path[0] === activeTabIdx 
        && selectedElement.path[1] === grpIdx
        && selectedElement.path[2] === ctrlIdx;

    if (!ctrl.isMso) {
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
    } else {
        ctrlEl.draggable = false;
    }

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
        const firstLabel = ctrl.items && ctrl.items.length > 0 ? ctrl.items[0].label : (ctrl.label || '선택 없음');
        ctrlEl.innerHTML = `<span class="control-label">${ctrl.isMso ? '목록' : ctrl.label}</span><div class="dropdown-sim-select-light"><span>${ctrl.isMso ? ctrl.label : firstLabel}</span><span class="dropdown-arrow-sim-light">▼</span></div>`;
        
    } else if (ctrl.type === 'box') {
        ctrlEl.className = "ribbon-control-sim type-box";
        if (isCtrlSelected) ctrlEl.classList.add("selected");
        ctrlEl.innerHTML = `<span class="control-label" style="font-style:italic; color:#888; font-size:10px;">📦 상자 [${ctrl.boxStyle || 'vertical'}, ${ctrl.stackLimit || 3}개]</span><span class="control-label">${ctrl.label}</span>`;
        
    } else if (ctrl.type === 'buttongroup') {
        ctrlEl.className = "ribbon-control-sim type-buttongroup";
        if (isCtrlSelected) ctrlEl.classList.add("selected");
        ctrlEl.innerHTML = `<span class="control-label" style="font-style:italic; color:#888; font-size:9px;">🔗 버튼 그룹</span><span class="control-label">${ctrl.label || ''}</span>`;
        
    } else if (ctrl.type === 'gallery') {
        ctrlEl.className = `ribbon-control-sim size-${ctrl.size || 'large'} type-gallery`;
        if (isCtrlSelected) ctrlEl.classList.add("selected");
        const icon = getIconHtml(ctrl.imageMso, ctrl.size === 'large' ? 24 : 14);
        
        let dropdownItemsHtml = '';
        if (ctrl.items && ctrl.items.length > 0) {
            dropdownItemsHtml = `
                <div class="gallery-sim-dropdown" style="display:none; position:absolute; top:100%; left:0; background:white; border:1px solid #ccc; border-radius:4px; padding:6px; box-shadow:0 2px 8px rgba(0,0,0,0.15); z-index:1000; width:${(ctrl.columns || 5) * 28 + 12}px; grid-template-columns:repeat(${ctrl.columns || 5}, 1fr); gap:4px;">
            `;
            ctrl.items.forEach(item => {
                let colorStyle = "";
                const hexRegex = /^#([0-9A-F]{3}){1,2}$/i;
                if (hexRegex.test(item.imageMso)) {
                    colorStyle = `background-color:${item.imageMso};`;
                    if (item.imageMso.toLowerCase() === '#ffffff' || item.imageMso.toLowerCase() === '#fff') {
                        colorStyle += " border: 1px solid #ddd;";
                    }
                } else if (hexRegex.test(item.label)) {
                    colorStyle = `background-color:${item.label};`;
                    if (item.label.toLowerCase() === '#ffffff' || item.label.toLowerCase() === '#fff') {
                        colorStyle += " border: 1px solid #ddd;";
                    }
                } else if (item.imageMso === "AppointmentColor1" || item.label.includes("빨강")) colorStyle = "background-color:#FF5A5F;";
                else if (item.imageMso === "AppointmentColor2" || item.label.includes("주황")) colorStyle = "background-color:#FFA726;";
                else if (item.imageMso === "AppointmentColor3" || item.label.includes("노랑")) colorStyle = "background-color:#FFEE58;";
                else if (item.imageMso === "AppointmentColor4" || item.label.includes("초록")) colorStyle = "background-color:#41B883;";
                else if (item.imageMso === "AppointmentColor5" || item.label.includes("파랑")) colorStyle = "background-color:#5E97F6;";
                else if (item.imageMso === "AppointmentColor6" || item.label.includes("남색")) colorStyle = "background-color:#1E3A8A;";
                else if (item.imageMso === "AppointmentColor7" || item.label.includes("보라")) colorStyle = "background-color:#AB47BC;";
                else if (item.imageMso === "AppointmentColor8" || item.label.includes("검정")) colorStyle = "background-color:#000000;";
                else if (item.imageMso === "AppointmentColor9" || item.label.includes("회색")) colorStyle = "background-color:#90A4AE;";
                else if (item.imageMso === "AppointmentColor0" || item.label.includes("흰색")) colorStyle = "background-color:#FFFFFF; border: 1px solid #ddd;";
                else colorStyle = "background-color:#ddd;";
                
                dropdownItemsHtml += `
                    <div class="gallery-tile-sim" title="${item.label} (${item.imageMso || ''})" style="width:${ctrl.itemWidth || 24}px; height:${ctrl.itemHeight || 24}px; border-radius:2px; cursor:pointer; ${colorStyle}" onclick="event.stopPropagation(); alert('갤러리 항목 선택: ${item.label} (ID: ${item.id})')"></div>
                `;
            });
            dropdownItemsHtml += `</div>`;
        }
        
        ctrlEl.innerHTML = `
            <div class="control-icon">${icon}</div>
            <span class="control-label">${ctrl.label} ▾</span>
            ${dropdownItemsHtml}
        `;
        
        ctrlEl.style.position = 'relative';
        ctrlEl.addEventListener("click", (e) => {
            e.stopPropagation();
            const dp = ctrlEl.querySelector(".gallery-sim-dropdown");
            if (dp) {
                const wasVisible = dp.style.display === 'grid';
                document.querySelectorAll(".gallery-sim-dropdown, .splitbutton-sim-dropdown, .menu-dropdown-sim").forEach(d => d.style.display = 'none');
                dp.style.display = wasVisible ? 'none' : 'grid';
            }
        });

    } else if (ctrl.type === 'splitbutton') {
        ctrlEl.className = `ribbon-control-sim size-${ctrl.size || 'large'} type-splitbutton`;
        if (isCtrlSelected) ctrlEl.classList.add("selected");
        ctrlEl.style.display = "flex";
        ctrlEl.style.alignItems = "stretch";
        ctrlEl.style.padding = "0";
        ctrlEl.style.position = 'relative';
        
        const icon = getIconHtml(ctrl.imageMso, ctrl.size === 'large' ? 24 : 14);
        
        let dropdownMenuHtml = '';
        if (ctrl.items && ctrl.items.length > 0) {
            dropdownMenuHtml = `
                <div class="splitbutton-sim-dropdown" style="display:none; position:absolute; top:100%; left:0; background:white; border:1px solid #ccc; border-radius:4px; box-shadow:0 2px 8px rgba(0,0,0,0.15); z-index:1000; min-width:120px; padding:4px 0;">
            `;
            ctrl.items.forEach(item => {
                dropdownMenuHtml += `
                    <div class="splitbutton-dropdown-item" style="padding:6px 12px; font-size:11px; cursor:pointer; color:#323130;" onmouseover="this.style.background='#f3f3f2'" onmouseout="this.style.background='transparent'" onclick="event.stopPropagation(); alert('분할 버튼 옵션 클릭: ${item.label} (ID: ${item.id})')">${item.label}</div>
                `;
            });
            dropdownMenuHtml += `</div>`;
        }
        
        if (ctrl.size === 'large') {
            ctrlEl.innerHTML = `
                <div class="splitbutton-left-btn" style="flex:1; display:flex; flex-direction:column; align-items:center; justify-content:center; padding:4px;" onclick="event.stopPropagation(); alert('분할 버튼 메인 클릭: ${ctrl.label} (onAction: ${ctrl.onAction})')">
                    <div class="control-icon" style="margin-bottom:2px;">${icon}</div>
                    <span class="control-label">${ctrl.label}</span>
                </div>
                <div class="splitbutton-right-arrow" style="width:16px; border-left:1px solid #ddd; display:flex; align-items:center; justify-content:center; cursor:pointer;" onclick="event.stopPropagation();">
                    <span style="font-size:9px; color:#666;">▼</span>
                </div>
                ${dropdownMenuHtml}
            `;
        } else {
            ctrlEl.innerHTML = `
                <div class="splitbutton-left-btn" style="flex:1; display:flex; align-items:center; gap:4px; padding:2px 4px;" onclick="event.stopPropagation(); alert('분할 버튼 메인 클릭: ${ctrl.label} (onAction: ${ctrl.onAction})')">
                    <div class="control-icon">${icon}</div>
                    <span class="control-label">${ctrl.label}</span>
                </div>
                <div class="splitbutton-right-arrow" style="width:12px; border-left:1px solid #ddd; display:flex; align-items:center; justify-content:center; cursor:pointer;" onclick="event.stopPropagation();">
                    <span style="font-size:8px; color:#666;">▼</span>
                </div>
                ${dropdownMenuHtml}
            `;
        }
        
        const rightArrow = ctrlEl.querySelector(".splitbutton-right-arrow");
        if (rightArrow) {
            rightArrow.addEventListener("click", (e) => {
                e.stopPropagation();
                const dp = ctrlEl.querySelector(".splitbutton-sim-dropdown");
                if (dp) {
                    const wasVisible = dp.style.display === 'block';
                    document.querySelectorAll(".gallery-sim-dropdown, .splitbutton-sim-dropdown, .menu-dropdown-sim").forEach(d => d.style.display = 'none');
                    dp.style.display = wasVisible ? 'none' : 'block';
                }
            });
        }

    } else if (ctrl.type === 'labelcontrol') {
        ctrlEl.className = "ribbon-control-sim type-labelcontrol";
        if (isCtrlSelected) ctrlEl.classList.add("selected");
        ctrlEl.textContent = ctrl.label || '';
        
    } else if (ctrl.type === 'menu') {
        ctrlEl.className = `ribbon-control-sim size-${ctrl.size || 'large'} type-menu`;
        if (isCtrlSelected) ctrlEl.classList.add("selected");
        ctrlEl.style.position = 'relative';
        const icon = getIconHtml(ctrl.imageMso, ctrl.size === 'large' ? 24 : 14);
        
        let dropdownHtml = '';
        if (ctrl.items && ctrl.items.length > 0) {
            dropdownHtml = `
                <div class="menu-dropdown-sim" style="display:none; position:absolute; top:100%; left:0; background:white; border:1px solid #ccc; border-radius:4px; box-shadow:0 2px 8px rgba(0,0,0,0.15); z-index:1000; min-width:120px; padding:4px 0;">
            `;
            ctrl.items.forEach(item => {
                dropdownHtml += `
                    <div class="splitbutton-dropdown-item" style="padding:6px 12px; font-size:11px; cursor:pointer; color:#323130; text-align:left;" onmouseover="this.style.background='#f3f3f2'" onmouseout="this.style.background='transparent'" onclick="event.stopPropagation(); alert('메뉴 옵션 클릭: ${item.label} (ID: ${item.id})')">${item.label}</div>
                `;
            });
            dropdownHtml += `</div>`;
        }
        
        ctrlEl.innerHTML = `
            <div class="control-icon">${icon}</div>
            <span class="control-label">${ctrl.label} ▾</span>
            ${dropdownHtml}
        `;
        
        ctrlEl.addEventListener("click", (e) => {
            e.stopPropagation();
            const dp = ctrlEl.querySelector(".menu-dropdown-sim");
            if (dp) {
                const wasVisible = dp.style.display === 'block';
                document.querySelectorAll(".gallery-sim-dropdown, .splitbutton-sim-dropdown, .menu-dropdown-sim").forEach(d => d.style.display = 'none');
                dp.style.display = wasVisible ? 'none' : 'block';
            }
            selectedElement = { type: 'control', path: [activeTabIdx, grpIdx, ctrlIdx] };
            updateRibbonUI();
        });

    } else if (ctrl.type === 'dynamicmenu') {
        ctrlEl.className = `ribbon-control-sim size-${ctrl.size || 'large'} type-dynamicmenu`;
        if (isCtrlSelected) ctrlEl.classList.add("selected");
        ctrlEl.style.position = 'relative';
        const icon = getIconHtml(ctrl.imageMso, ctrl.size === 'large' ? 24 : 14);
        
        let dropdownHtml = `
            <div class="menu-dropdown-sim" style="display:none; position:absolute; top:100%; left:0; background:white; border:1px solid #ccc; border-radius:4px; box-shadow:0 2px 8px rgba(0,0,0,0.15); z-index:1000; min-width:180px; padding:8px 10px; font-size:11px; color:#555; text-align:left;">
                <div style="font-weight:bold; color:var(--excel-green-primary); border-bottom:1px solid #eee; padding-bottom:4px; margin-bottom:4px;">⚡ 동적 생성 메뉴 (Dynamic Menu)</div>
                <div style="font-style:italic; color:#888; margin-bottom:4px; word-break:break-all;">getContent: ${ctrl.getContent || ctrl.id + '_GetContent'}</div>
                <div style="color:#666; font-size:10px; margin-bottom:4px;">VBA 콜백 호출로 런타임에 아래 항목들이 로드됩니다:</div>
                <div style="padding-left:6px; border-left:2px solid #ddd; color:#333; font-size:10.5px;">
                    • 동적 버튼 1<br>• 동적 버튼 2
                </div>
            </div>
        `;
        
        ctrlEl.innerHTML = `
            <div class="control-icon">${icon}</div>
            <span class="control-label">${ctrl.label} ▾</span>
            ${dropdownHtml}
        `;
        
        ctrlEl.addEventListener("click", (e) => {
            e.stopPropagation();
            const dp = ctrlEl.querySelector(".menu-dropdown-sim");
            if (dp) {
                const wasVisible = dp.style.display === 'block';
                document.querySelectorAll(".gallery-sim-dropdown, .splitbutton-sim-dropdown, .menu-dropdown-sim").forEach(d => d.style.display = 'none');
                dp.style.display = wasVisible ? 'none' : 'block';
            }
            selectedElement = { type: 'control', path: [activeTabIdx, grpIdx, ctrlIdx] };
            updateRibbonUI();
        });
        
    } else if (ctrl.type === 'togglebutton') {
        ctrlEl.className = `ribbon-control-sim size-${ctrl.size || 'large'} type-togglebutton`;
        if (isCtrlSelected) ctrlEl.classList.add("selected");
        if (ctrl.checked) ctrlEl.classList.add("pressed");
        
        if (ctrl.isMso) {
            const icon = getIconHtml(ctrl.imageMso, ctrl.size === 'large' ? 24 : 14);
            ctrlEl.innerHTML = `<div class="control-icon">${icon}</div><span class="control-label" style="font-weight:bold;">${ctrl.label}</span>`;
        } else {
            const radioCircle = ctrl.checked ? '●' : '○';
            if (ctrl.size === 'large') {
                ctrlEl.innerHTML = `<span class="control-label"><span class="radio-sim-circle">${radioCircle}</span> ${ctrl.label}</span>`;
            } else {
                ctrlEl.innerHTML = `<span class="radio-sim-circle">${radioCircle}</span><span class="control-label">${ctrl.label}</span>`;
            }
        }
        
    } else {
        // Standard button
        ctrlEl.className = `ribbon-control-sim size-${ctrl.size || 'large'}`;
        if (isCtrlSelected) ctrlEl.classList.add("selected");
        const icon = getIconHtml(ctrl.imageMso, ctrl.size === 'large' ? 24 : 14);
        ctrlEl.innerHTML = `<div class="control-icon">${icon}</div><span class="control-label">${ctrl.label}</span>`;
    }

    if (ctrl.isMso) {
        ctrlEl.addEventListener("click", (e) => {
            e.stopPropagation();
            showValidationBanner(true, `내장 컨트롤 클릭: ${ctrl.label} (imageMso="${ctrl.imageMso || ''}") - 이 컨트롤은 읽기 전용입니다.`);
        });
    } else {
        // Toggle interaction on click directly in simulator
        ctrlEl.addEventListener("click", (e) => {
            e.stopPropagation();
            selectedElement = { type: 'control', path: [activeTabIdx, grpIdx, ctrlIdx] };
            
            if (ctrl.type === 'checkbox') {
                ctrl.checked = !ctrl.checked;
            } else if (ctrl.type === 'togglebutton') {
                ctrl.checked = !ctrl.checked;
                if (ctrl.checked) {
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
    }

    return ctrlEl;
}

function getMockControlsForGroup(idMso) {
    switch (idMso) {
        case 'GroupClipboard':
            return [
                { id: 'mso_paste', type: 'button', label: '붙여넣기', size: 'large', imageMso: 'Paste', isMso: true },
                { id: 'mso_cut', type: 'button', label: '잘라내기', size: 'normal', imageMso: 'Cut', isMso: true },
                { id: 'mso_copy', type: 'button', label: '복사', size: 'normal', imageMso: 'Copy', isMso: true },
                { id: 'mso_painter', type: 'button', label: '서식 복사', size: 'normal', imageMso: 'FormatPainter', isMso: true }
            ];
        case 'GroupFont':
            return [
                { id: 'mso_fontname', type: 'dropdown', label: '맑은 고딕', size: 'normal', isMso: true, items: [{ label: '맑은 고딕' }] },
                { id: 'mso_fontsize', type: 'dropdown', label: '11', size: 'normal', isMso: true, items: [{ label: '11' }] },
                { id: 'mso_bold', type: 'togglebutton', label: '가', size: 'normal', imageMso: 'Bold', isMso: true, checked: false },
                { id: 'mso_italic', type: 'togglebutton', label: '배', size: 'normal', imageMso: 'Italic', isMso: true, checked: false },
                { id: 'mso_underline', type: 'togglebutton', label: '밑', size: 'normal', imageMso: 'Underline', isMso: true, checked: false }
            ];
        case 'GroupEnterDataAlignment':
            return [
                { id: 'mso_align_left', type: 'button', label: '왼쪽 맞춤', size: 'normal', imageMso: 'AlignLeft', isMso: true },
                { id: 'mso_align_center', type: 'button', label: '가운데 맞춤', size: 'normal', imageMso: 'AlignCenter', isMso: true },
                { id: 'mso_align_right', type: 'button', label: '오른쪽 맞춤', size: 'normal', imageMso: 'AlignRight', isMso: true },
                { id: 'mso_wrap_text', type: 'togglebutton', label: '텍스트 줄 바꿈', size: 'normal', imageMso: 'WrapText', isMso: true }
            ];
        case 'GroupEnterDataNumber':
            return [
                { id: 'mso_number_format', type: 'dropdown', label: '일반', size: 'normal', isMso: true, items: [{ label: '일반' }] },
                { id: 'mso_currency', type: 'button', label: '통화 표시', size: 'normal', imageMso: 'AccountingNumberFormat', isMso: true },
                { id: 'mso_percent', type: 'button', label: '백분율', size: 'normal', imageMso: 'PercentStyle', isMso: true },
                { id: 'mso_comma', type: 'button', label: '쉼표 스타일', size: 'normal', imageMso: 'CommaStyle', isMso: true }
            ];
        case 'GroupQuickFormatting':
            return [
                { id: 'mso_cond_fmt', type: 'button', label: '조건부 서식', size: 'large', imageMso: 'ConditionalFormattingMenu', isMso: true },
                { id: 'mso_table_fmt', type: 'button', label: '표 서식', size: 'large', imageMso: 'FormatAsTable', isMso: true },
                { id: 'mso_cell_styles', type: 'button', label: '셀 스타일', size: 'large', imageMso: 'CellStylesGallery', isMso: true }
            ];
        case 'GroupIllustrations':
            return [
                { id: 'mso_insert_pic', type: 'button', label: '그림', size: 'large', imageMso: 'PictureInsert', isMso: true },
                { id: 'mso_insert_shape', type: 'button', label: '도형', size: 'large', imageMso: 'ShapesInsert', isMso: true }
            ];
        default:
            return [];
    }
}

function renderRepurposedCommandsNotice() {
    const ribbonBody = document.getElementById("sim-ribbon-body");
    if (!ribbonState.commands || ribbonState.commands.length === 0) return;

    const banner = document.createElement("div");
    banner.className = "repurposed-commands-banner";
    banner.style.width = "100%";
    banner.style.padding = "12px 16px";
    banner.style.backgroundColor = "rgba(43, 87, 154, 0.08)";
    banner.style.border = "1px solid rgba(43, 87, 154, 0.2)";
    banner.style.borderRadius = "4px";
    banner.style.marginBottom = "10px";
    banner.style.fontSize = "11.5px";
    banner.style.color = "#2b579a";
    banner.style.display = "flex";
    banner.style.flexDirection = "column";
    banner.style.gap = "6px";
    banner.style.gridColumn = "1 / -1";

    banner.innerHTML = `
        <div style="font-weight: 700; display: flex; align-items: center; gap: 6px;">
            <span>🔄 엑셀 기본 명령어 재정의 (Command Repurposing) 감지됨</span>
        </div>
        <div style="font-size: 11px; color: #555; line-height: 1.45;">
            이 통합 문서가 엑셀에서 열릴 때 아래 엑셀 기본 기능들의 핵심 동작이 재정의됩니다.
        </div>
        <div class="repurposed-list" style="display: flex; flex-direction: column; gap: 4px; margin-top: 4px;">
            ${ribbonState.commands.map(cmd => {
                const actionText = cmd.onAction ? `<span style="font-family: monospace; background: #e1dfdd; padding: 2px 4px; border-radius: 2px; color: #2b579a;">Sub ${cmd.onAction}</span> 매크로 함수 호출` : '';
                const stateText = cmd.enabled === false ? '<span style="color:#d13438; font-weight:600;">기능 사용 중지 (Disabled)</span>' : '';
                const connection = actionText && stateText ? ' 및 ' : '';
                return `
                    <div style="display: flex; align-items: center; gap: 8px; background: #ffffff; padding: 6px 10px; border-radius: 4px; border: 1px solid #edebe9;">
                        <span style="font-weight: 600; color: #323130; min-width: 120px;">idMso="${cmd.idMso}"</span>
                        <span style="color: #605e5c;">➔</span>
                        <span>${actionText}${connection}${stateText}</span>
                    </div>
                `;
            }).join('')}
        </div>
    `;
    
    if (ribbonBody.firstChild) {
        ribbonBody.insertBefore(banner, ribbonBody.firstChild);
    } else {
        ribbonBody.appendChild(banner);
    }
}

function renderBackstage() {
    const simWindow = document.querySelector(".simulated-window");
    if (!simWindow) return;

    let overlay = document.getElementById("sim-backstage-overlay");
    if (!overlay) {
        overlay = document.createElement("div");
        overlay.id = "sim-backstage-overlay";
        overlay.className = "backstage-view";
        simWindow.appendChild(overlay);
    }

    overlay.style.display = isBackstageOpen ? "flex" : "none";
    if (!isBackstageOpen) return;

    overlay.innerHTML = "";

    // 1. Sidebar Panel
    const sidebar = document.createElement("div");
    sidebar.className = "backstage-sidebar";

    // Back arrow
    const backBtn = document.createElement("button");
    backBtn.className = "backstage-back-btn";
    backBtn.innerHTML = "←";
    backBtn.title = "시트로 돌아가기";
    backBtn.addEventListener("click", () => {
        isBackstageOpen = false;
        updateRibbonUI();
    });
    sidebar.appendChild(backBtn);

    // List of tabs
    const menuItems = [
        { id: 'mso_info', label: '정보', isMso: true },
        { id: 'mso_new', label: '새로 만들기', isMso: true },
        { id: 'mso_open', label: '열기', isMso: true },
        { id: 'mso_print', label: '인쇄', isMso: true }
    ];

    // Add custom backstage tabs
    if (ribbonState.backstage && ribbonState.backstage.tabs) {
        ribbonState.backstage.tabs.forEach(tab => {
            menuItems.push({ id: tab.id, label: tab.label, isCustom: true, rawTab: tab });
        });
    }

    if (!activeBackstageTabId) {
        activeBackstageTabId = menuItems[0].id;
    }

    menuItems.forEach(item => {
        const itemEl = document.createElement("div");
        itemEl.className = "backstage-menu-item";
        if (item.isCustom) itemEl.classList.add("custom-tab-item");
        if (item.id === activeBackstageTabId) itemEl.classList.add("active");
        itemEl.textContent = item.label;
        itemEl.addEventListener("click", () => {
            activeBackstageTabId = item.id;
            renderBackstage();
        });
        sidebar.appendChild(itemEl);
    });

    // Divider before fast buttons
    const sep = document.createElement("div");
    sep.className = "backstage-menu-separator";
    sidebar.appendChild(sep);

    // Fast buttons (Save, Save As, etc.)
    const fastButtons = [];
    if (ribbonState.backstage && ribbonState.backstage.buttons) {
        ribbonState.backstage.buttons.forEach(btn => {
            fastButtons.push({ id: btn.id, label: btn.label, onAction: btn.onAction });
        });
    } else {
        fastButtons.push({ id: 'mso_save', label: '저장', onAction: null });
        fastButtons.push({ id: 'mso_save_as', label: '다른 이름으로 저장', onAction: null });
    }
    
    fastButtons.push({ id: 'mso_close', label: '닫기', isClose: true });

    fastButtons.forEach(btn => {
        const btnEl = document.createElement("div");
        btnEl.className = "backstage-menu-item";
        btnEl.textContent = btn.label;
        btnEl.addEventListener("click", () => {
            if (btn.isClose) {
                isBackstageOpen = false;
                updateRibbonUI();
            } else if (btn.onAction) {
                showValidationBanner(true, `Backstage 콜백 호출: Sub ${btn.onAction}(control)`);
            } else {
                showValidationBanner(true, `백스테이지 '${btn.label}' 명령이 클릭되었습니다.`);
            }
        });
        sidebar.appendChild(btnEl);
    });

    overlay.appendChild(sidebar);

    // 2. Right Workspace Content Area
    const content = document.createElement("div");
    content.className = "backstage-content";

    const titleEl = document.createElement("div");
    titleEl.className = "backstage-content-title";
    
    const activeItem = menuItems.find(item => item.id === activeBackstageTabId);
    titleEl.textContent = activeItem ? activeItem.label : "Backstage";
    content.appendChild(titleEl);

    // Render contents based on active tab
    if (activeBackstageTabId === 'mso_info') {
        content.appendChild(renderBackstageMsoInfo());
    } else if (activeBackstageTabId === 'mso_new') {
        content.appendChild(renderBackstageMsoNew());
    } else if (activeBackstageTabId === 'mso_open') {
        content.appendChild(renderBackstageMsoOpen());
    } else if (activeBackstageTabId === 'mso_print') {
        content.appendChild(renderBackstageMsoPrint());
    } else if (activeItem && activeItem.isCustom && activeItem.rawTab) {
        const tab = activeItem.rawTab;
        
        const colsLayout = document.createElement("div");
        colsLayout.className = "backstage-columns-layout";
        
        // Column 1
        const col1 = document.createElement("div");
        col1.className = "backstage-column";
        if (tab.firstColumn && tab.firstColumn.groups) {
            tab.firstColumn.groups.forEach(tg => {
                col1.appendChild(renderBackstageTaskGroup(tg));
            });
        }
        colsLayout.appendChild(col1);

        // Column 2
        const col2 = document.createElement("div");
        col2.className = "backstage-column";
        if (tab.secondColumn && tab.secondColumn.groups) {
            tab.secondColumn.groups.forEach(tg => {
                col2.appendChild(renderBackstageTaskGroup(tg));
            });
        }
        colsLayout.appendChild(col2);

        content.appendChild(colsLayout);
    }

    overlay.appendChild(content);
}

function renderBackstageTaskGroup(tg) {
    const tgEl = document.createElement("div");
    tgEl.className = "backstage-task-group";
    
    const title = document.createElement("div");
    title.className = "backstage-task-group-title";
    title.textContent = tg.label || "작업 그룹";
    tgEl.appendChild(title);

    if (tg.categories) {
        tg.categories.forEach(cat => {
            const catEl = document.createElement("div");
            catEl.className = "backstage-category";
            
            if (cat.label) {
                const catTitle = document.createElement("div");
                catTitle.className = "backstage-category-title";
                catTitle.textContent = cat.label;
                catEl.appendChild(catTitle);
            }

            const tasksList = document.createElement("div");
            tasksList.className = "backstage-tasks-list";

            if (cat.tasks) {
                cat.tasks.forEach(task => {
                    const taskEl = document.createElement("div");
                    taskEl.className = "backstage-task-item";
                    
                    const icon = document.createElement("div");
                    icon.className = "backstage-task-item-icon";
                    icon.innerHTML = getIconHtml(task.imageMso, 20);
                    taskEl.appendChild(icon);

                    const details = document.createElement("div");
                    details.className = "backstage-task-item-details";
                    
                    const taskName = document.createElement("h4");
                    taskName.textContent = task.label || "작업";
                    details.appendChild(taskName);

                    const taskDesc = document.createElement("p");
                    taskDesc.textContent = task.onAction ? `연결된 매크로: Sub ${task.onAction}(control)` : "동작 매크로 없음";
                    details.appendChild(taskDesc);

                    taskEl.appendChild(details);

                    taskEl.addEventListener("click", () => {
                        if (task.onAction) {
                            showValidationBanner(true, `백스테이지 매크로 호출 완료: Sub ${task.onAction}(control)`);
                        } else {
                            showValidationBanner(true, `백스테이지 작업 클릭: ${task.label}`);
                        }
                    });

                    tasksList.appendChild(taskEl);
                });
            }
            catEl.appendChild(tasksList);
            tgEl.appendChild(catEl);
        });
    }
    return tgEl;
}

function renderBackstageMsoInfo() {
    const div = document.createElement("div");
    div.style.display = "flex";
    div.style.flexDirection = "column";
    div.style.gap = "16px";
    div.innerHTML = `
        <div style="display:flex; gap:16px;">
            <div style="flex:1; border:1px solid #eaeaea; background:#fff; padding:16px; border-radius:4px; display:flex; gap:12px; align-items:flex-start;">
                <span style="font-size:24px; color:#107c41;">🔒</span>
                <div>
                    <h3 style="margin:0 0 6px 0; font-size:14px; font-weight:600;">통합 문서 보호</h3>
                    <p style="margin:0; font-size:11px; color:#666; line-height:1.4;">이 통합 문서를 열거나 변경할 수 있는 사용자를 제어합니다.</p>
                </div>
            </div>
            <div style="flex:1; border:1px solid #eaeaea; background:#fff; padding:16px; border-radius:4px; display:flex; gap:12px; align-items:flex-start;">
                <span style="font-size:24px; color:#107c41;">🔍</span>
                <div>
                    <h3 style="margin:0 0 6px 0; font-size:14px; font-weight:600;">통합 문서 검사</h3>
                    <p style="margin:0; font-size:11px; color:#666; line-height:1.4;">파일을 배포하기 전에 개인 정보 또는 숨겨진 메타데이터가 있는지 확인합니다.</p>
                </div>
            </div>
        </div>
        <div style="border:1px solid #eaeaea; background:#fff; padding:16px; border-radius:4px; font-size:11px; color:#555; line-height:1.6;">
            <strong>파일 정보 요약:</strong><br>
            • 파일 이름: Book1.xlsm (VBA 매크로 포함 통합 문서)<br>
            • XML 스키마 버전: ${activeSchemaVersion === 'customUI' ? 'Office 2007 (customUI)' : 'Office 2010+ (customUI14)'}<br>
            • 리본 탭 개수: ${ribbonState.tabs.length}개<br>
            • 명령 재정의 개수: ${ribbonState.commands.length}개
        </div>
    `;
    return div;
}

function renderBackstageMsoNew() {
    const div = document.createElement("div");
    div.style.display = "grid";
    div.style.gridTemplateColumns = "repeat(4, 1fr)";
    div.style.gap = "16px";
    div.innerHTML = `
        <div style="border:1px solid #eaeaea; background:#fff; padding:12px; border-radius:4px; text-align:center; cursor:pointer; min-height:120px; display:flex; flex-direction:column; justify-content:center; align-items:center; gap:8px;">
            <div style="width:50px; height:65px; border:1px dashed #bbb; background:#fafafa; border-radius:2px;"></div>
            <div style="font-size:11.5px; font-weight:600;">새 통합 문서</div>
        </div>
        <div style="border:1px solid #eaeaea; background:#fff; padding:12px; border-radius:4px; text-align:center; cursor:pointer; min-height:120px; display:flex; flex-direction:column; justify-content:center; align-items:center; gap:8px; opacity:0.8;">
            <div style="width:50px; height:65px; border:1px solid #badbcc; background:#e2f0e9; border-radius:2px; display:flex; align-items:center; justify-content:center; color:#107c41; font-size:9px; font-weight:bold;">Excel</div>
            <div style="font-size:11.5px;">개인 예산 가계부</div>
        </div>
        <div style="border:1px solid #eaeaea; background:#fff; padding:12px; border-radius:4px; text-align:center; cursor:pointer; min-height:120px; display:flex; flex-direction:column; justify-content:center; align-items:center; gap:8px; opacity:0.8;">
            <div style="width:50px; height:65px; border:1px solid #badbcc; background:#e2f0e9; border-radius:2px; display:flex; align-items:center; justify-content:center; color:#107c41; font-size:9px; font-weight:bold;">Excel</div>
            <div style="font-size:11.5px;">매출 보고서 분석</div>
        </div>
        <div style="border:1px solid #eaeaea; background:#fff; padding:12px; border-radius:4px; text-align:center; cursor:pointer; min-height:120px; display:flex; flex-direction:column; justify-content:center; align-items:center; gap:8px; opacity:0.8;">
            <div style="width:50px; height:65px; border:1px solid #badbcc; background:#e2f0e9; border-radius:2px; display:flex; align-items:center; justify-content:center; color:#107c41; font-size:9px; font-weight:bold;">Excel</div>
            <div style="font-size:11.5px;">간트 차트 프로젝트</div>
        </div>
    `;
    return div;
}

function renderBackstageMsoOpen() {
    const div = document.createElement("div");
    div.style.background = "#fff";
    div.style.border = "1px solid #eaeaea";
    div.style.borderRadius = "4px";
    div.style.padding = "16px";
    div.innerHTML = `
        <h4 style="margin:0 0 12px 0; font-size:12.5px; font-weight:600; color:#107c41;">최근 사용한 문서</h4>
        <div style="display:flex; flex-direction:column; gap:8px;">
            <div style="display:flex; justify-content:between; padding:8px 12px; border-bottom:1px solid #f3f2f1; font-size:11.5px; cursor:pointer;">
                <span style="font-weight:600; color:#323130;">CustomRibbon_Workbook.xlsm</span>
                <span style="color:#888; font-size:11px; margin-left:auto;">오늘, 오후 11:15</span>
            </div>
            <div style="display:flex; justify-content:between; padding:8px 12px; border-bottom:1px solid #f3f2f1; font-size:11.5px; cursor:pointer;">
                <span style="font-weight:600; color:#323130;">자동화_매크로_템플릿.xlsm</span>
                <span style="color:#888; font-size:11px; margin-left:auto;">어제, 오전 10:30</span>
            </div>
            <div style="display:flex; justify-content:between; padding:8px 12px; border-bottom:1px solid #f3f2f1; font-size:11.5px; cursor:pointer;">
                <span style="font-weight:600; color:#323130;">부서원_근태_취합결과.xlsx</span>
                <span style="color:#888; font-size:11px; margin-left:auto;">2026-06-03</span>
            </div>
        </div>
    `;
    return div;
}

function renderBackstageMsoPrint() {
    const div = document.createElement("div");
    div.style.display = "flex";
    div.style.gap = "24px";
    div.innerHTML = `
        <div style="flex:1; display:flex; flex-direction:column; gap:12px;">
            <button class="btn-primary" style="align-self:flex-start; background:#107c41; color:#fff; border:none; padding:10px 24px; font-weight:bold; border-radius:4px; font-size:13px; cursor:pointer;">인쇄</button>
            <div style="font-size:11.5px; margin-top:8px;">
                <strong>프린터 설정:</strong><br>
                <select style="width:100%; padding:6px; margin-top:4px; border:1px solid #ccc; border-radius:4px;">
                    <option>Microsoft Print to PDF (기본)</option>
                    <option>PDF 빌더 가상 프린터</option>
                </select>
            </div>
            <div style="font-size:11.5px; margin-top:8px;">
                <strong>인쇄 범위:</strong><br>
                <select style="width:100%; padding:6px; margin-top:4px; border:1px solid #ccc; border-radius:4px;">
                    <option>활성 시트 인쇄</option>
                    <option>전체 통합 문서 인쇄</option>
                    <option>선택 영역 인쇄</option>
                </select>
            </div>
        </div>
        <div style="flex:1.5; border:1px solid #eaeaea; background:#fff; border-radius:4px; padding:20px; display:flex; justify-content:center; align-items:center; min-height:220px; box-shadow:var(--shadow-sm);">
            <div style="width:140px; height:180px; border:1px solid #bbb; box-shadow:0 2px 4px rgba(0,0,0,0.1); background:#fff; padding:10px; display:flex; flex-direction:column; gap:6px;">
                <div style="height:12px; background:#e2f0e9; width:60%;"></div>
                <div style="height:6px; background:#f3f2f1; width:100%;"></div>
                <div style="height:6px; background:#f3f2f1; width:90%;"></div>
                <div style="height:40px; border:1px solid #c8c6c4; background:#fafafa; border-radius:2px; display:flex; align-items:center; justify-content:center; font-size:8px; color:#888;">차트 영역 프리뷰</div>
                <div style="height:6px; background:#f3f2f1; width:100%;"></div>
                <div style="height:6px; background:#f3f2f1; width:80%;"></div>
            </div>
        </div>
    `;
    return div;
}

function renderRibbonSimulator() {
    const tabsBar = document.getElementById("sim-tabs-bar");
    const ribbonBody = document.getElementById("sim-ribbon-body");
    
    tabsBar.innerHTML = "";
    ribbonBody.innerHTML = "";

    // A. Render standard "파일" tab (Excel backstage trigger)
    const fileTabEl = document.createElement("div");
    fileTabEl.className = "excel-tab-header file-tab";
    if (isBackstageOpen) {
        fileTabEl.classList.add("active");
    }
    fileTabEl.textContent = "파일";
    fileTabEl.addEventListener("click", (e) => {
        e.stopPropagation();
        isBackstageOpen = !isBackstageOpen;
        if (isBackstageOpen) {
            // Find first backstage tab to show
            if (ribbonState.backstage && ribbonState.backstage.tabs && ribbonState.backstage.tabs.length > 0) {
                activeBackstageTabId = ribbonState.backstage.tabs[0].id;
            } else {
                activeBackstageTabId = "mso_info";
            }
        }
        updateRibbonUI();
    });
    tabsBar.appendChild(fileTabEl);

    // 1. Render normal tabs
    ribbonState.tabs.forEach((tab, tabIdx) => {
        const tabEl = document.createElement("div");
        tabEl.className = "excel-tab-header custom-tab-indicator";
        if (tab.isStandardTab) {
            tabEl.className = "excel-tab-header file-tab";
        }
        
        const isSelected = !isBackstageOpen && selectedElement && selectedElement.type === 'tab' && selectedElement.path[0] === tabIdx;
        if (isSelected) {
            tabEl.classList.add("active");
        }
        
        tabEl.textContent = tab.label || "새 탭";
        tabEl.addEventListener("click", (e) => {
            e.stopPropagation();
            isBackstageOpen = false; // Close backstage if switching to regular tabs
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
    
    if (!isBackstageOpen && activeTab && activeTab.groups) {
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

            // Drag over / Drag enter / Drag leave / Drop for groups (only if not Mso group)
            if (!group.idMso) {
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
            }

            // Group Label (11px uniform)
            const labelEl = document.createElement("div");
            labelEl.className = "ribbon-group-label-sim";
            labelEl.innerHTML = (group.label || group.idMso || "새 그룹") + (group.idMso ? ' <span class="mso-read-only-badge">내장</span>' : '');
            groupEl.appendChild(labelEl);

            // Group controls holder
            const controlsWrapper = document.createElement("div");
            controlsWrapper.className = "ribbon-group-controls";

            // If it is built-in group, retrieve mock controls
            let groupControls = group.controls || [];
            if (group.idMso && groupControls.length === 0) {
                groupControls = getMockControlsForGroup(group.idMso);
            }

            if (groupControls && groupControls.length > 0) {
                let i = 0;
                while (i < groupControls.length) {
                    const ctrl = groupControls[i];
                    
                    if (ctrl.type === 'box' || ctrl.type === 'buttongroup') {
                        const stackLimit = ctrl.type === 'buttongroup' ? 5 : (ctrl.stackLimit || 3);
                        const children = [];
                        let j = i + 1;
                        while (j < groupControls.length && children.length < stackLimit) {
                            const candidate = groupControls[j];
                            if (candidate.type === 'separator' || candidate.type === 'box' || candidate.type === 'buttongroup') {
                                break;
                            }
                            children.push(candidate);
                            j++;
                        }
                        
                        const boxEl = renderSingleControl(ctrl, activeTabIdx, grpIdx, i, groupControls.length);
                        
                        const boxChildrenContainer = document.createElement("div");
                        boxChildrenContainer.className = ctrl.type === 'buttongroup' ? 'ribbon-buttongroup-children-sim' : `ribbon-box-children-sim ${ctrl.boxStyle || 'vertical'}`;
                        boxChildrenContainer.style.display = "flex";
                        boxChildrenContainer.style.flexDirection = ctrl.type === 'buttongroup' ? 'row' : ((ctrl.boxStyle === 'horizontal') ? 'row' : 'column');
                        boxChildrenContainer.style.gap = ctrl.type === 'buttongroup' ? "0px" : "4px";
                        boxChildrenContainer.style.marginTop = "4px";
                        boxChildrenContainer.style.padding = ctrl.type === 'buttongroup' ? "2px" : "4px";
                        boxChildrenContainer.style.border = ctrl.type === 'buttongroup' ? "1px solid #ccc" : "1px dashed #ccc";
                        boxChildrenContainer.style.borderRadius = "4px";
                        boxChildrenContainer.style.flexGrow = "1";
                        boxChildrenContainer.style.justifyContent = "center";
                        boxChildrenContainer.style.background = ctrl.type === 'buttongroup' ? "#f8f8f8" : "transparent";
                        
                        children.forEach(child => {
                            const childIdx = groupControls.indexOf(child);
                            const childEl = renderSingleControl(child, activeTabIdx, grpIdx, childIdx, groupControls.length);
                            boxChildrenContainer.appendChild(childEl);
                        });
                        
                        boxEl.appendChild(boxChildrenContainer);
                        controlsWrapper.appendChild(boxEl);
                        
                        i = j;
                    } else {
                        const ctrlEl = renderSingleControl(ctrl, activeTabIdx, grpIdx, i, groupControls.length);
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

    // 3. Render repurposed commands notice
    if (!isBackstageOpen) {
        renderRepurposedCommandsNotice();
    }

    // 4. Render Backstage Overlay View
    renderBackstage();
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
    } else if (type === 'gallery') {
        control = {
            id: newId,
            type: "gallery",
            label: `색상 팔레트 ${newIdx}`,
            size: "large",
            imageMso: "ColorPaletteExcel",
            columns: 5,
            itemWidth: 24,
            itemHeight: 24,
            onAction: `${newId}_Change`,
            enabled: true,
            visible: true,
            items: [
                { id: `${newId}_colorRed`, label: "빨강", imageMso: "AppointmentColor1" },
                { id: `${newId}_colorOrange`, label: "주황", imageMso: "AppointmentColor2" },
                { id: `${newId}_colorYellow`, label: "노랑", imageMso: "AppointmentColor3" },
                { id: `${newId}_colorGreen`, label: "초록", imageMso: "AppointmentColor4" },
                { id: `${newId}_colorBlue`, label: "파랑", imageMso: "AppointmentColor5" },
                { id: `${newId}_colorDarkBlue`, label: "남색", imageMso: "AppointmentColor6" },
                { id: `${newId}_colorPurple`, label: "보라", imageMso: "AppointmentColor7" },
                { id: `${newId}_colorBlack`, label: "검정", imageMso: "AppointmentColor8" },
                { id: `${newId}_colorGray`, label: "회색", imageMso: "AppointmentColor9" },
                { id: `${newId}_colorWhite`, label: "흰색", imageMso: "AppointmentColor0" }
            ]
        };
    } else if (type === 'splitbutton') {
        control = {
            id: newId,
            type: "splitbutton",
            label: `분할 버튼 ${newIdx}`,
            size: "large",
            imageMso: "PageSetupDialog",
            onAction: `${newId}_Click`,
            enabled: true,
            visible: true,
            items: [
                { id: `${newId}_sub1`, label: "하위 실행 1", onAction: `${newId}_sub1_Click` },
                { id: `${newId}_sub2`, label: "하위 실행 2", onAction: `${newId}_sub2_Click` }
            ]
        };
    } else if (type === 'dynamicmenu') {
        control = {
            id: newId,
            type: "dynamicmenu",
            label: `동적 메뉴 ${newIdx}`,
            size: "large",
            imageMso: "DrawingCanvas",
            getContent: `${newId}_GetContent`,
            enabled: true,
            visible: true
        };
    } else if (type === 'labelcontrol') {
        control = { id: newId, type: "labelcontrol", label: `레이블 정보 ${newIdx}` };
    } else if (type === 'menu') {
        control = { 
            id: newId, 
            type: "menu", 
            label: `하위 메뉴 ${newIdx}`, 
            size: "large", 
            imageMso: "Settings", 
            enabled: true, 
            visible: true,
            items: [
                { id: `${newId}_sub1`, label: "서브 명령 1", onAction: `${newId}_sub1_Click` },
                { id: `${newId}_sub2`, label: "서브 명령 2", onAction: `${newId}_sub2_Click` }
            ]
        };
    } else {
        // button or togglebutton
        control = { id: newId, type: type, label: `${type === 'togglebutton' ? '토글 버튼' : '새 버튼'} ${newIdx}`, size: "large", imageMso: "DeveloperTools", onAction: `${newId}_Click`, enabled: true, visible: true, checked: false };
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
            imageMso: "Settings", 
            enabled: true, 
            visible: true,
            items: [
                { id: `${newId}_sub1`, label: "서브 명령 1", onAction: `${newId}_sub1_Click` },
                { id: `${newId}_sub2`, label: "서브 명령 2", onAction: `${newId}_sub2_Click` }
            ]
        };
    } else {
        // button or togglebutton
        control = { id: newId, type: type, label: `${type === 'togglebutton' ? '토글 버튼' : '새 버튼'} ${newIdx}`, size: "large", imageMso: "DeveloperTools", onAction: `${newId}_Click`, enabled: true, visible: true, checked: false };
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
        if (!tab) {
            selectedElement = null;
            renderPropertyEditor();
            return;
        }
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
        const tab = ribbonState.tabs[path[0]];
        const group = tab && tab.groups ? tab.groups[path[1]] : null;
        if (!group) {
            selectedElement = null;
            renderPropertyEditor();
            return;
        }
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
        const tab = ribbonState.tabs[path[0]];
        const group = tab && tab.groups ? tab.groups[path[1]] : null;
        const ctrl = group && group.controls ? group.controls[path[2]] : null;
        if (!ctrl) {
            selectedElement = null;
            renderPropertyEditor();
            return;
        }
        
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

        if (!['labelcontrol', 'separator', 'dynamicmenu'].includes(ctrl.type)) {
            const labelText = ctrl.type === 'editbox' || ctrl.type === 'combobox' ? 'VBA 변경 함수 (onChange)' : 'VBA 매크로 함수 (onAction)';
            callbackMarkup = `
                <div class="form-group">
                    <label>${labelText}</label>
                    <input type="text" id="prop-ctrl-onaction" value="${ctrl.onAction || ''}">
                </div>
            `;
        }

        const hasSize = ['button', 'togglebutton', 'menu', 'gallery', 'splitbutton', 'dynamicmenu'].includes(ctrl.type);
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

        const hasIcon = ['button', 'togglebutton', 'menu', 'gallery', 'splitbutton', 'dynamicmenu'].includes(ctrl.type);
        if (hasIcon) {
            iconMarkup = `
                <div class="form-group" style="grid-column: span 1;">
                    <label>아이콘 설정</label>
                    <div class="icon-input-wrapper">
                        <input type="text" id="prop-ctrl-icon" value="${ctrl.imageMso || ''}" placeholder="imageMso 이름 입력" style="font-size:11px;">
                        <button type="button" class="btn-secondary" id="btn-open-icon-picker">🔍 검색/변경</button>
                        <div class="icon-preview-box" id="prop-icon-preview">
                            ${getIconHtml(ctrl.imageMso, 18)}
                        </div>
                    </div>
                </div>
            `;
        }

        const isSmallCtrl = ctrl.type !== 'separator' && ctrl.type !== 'box' && ctrl.type !== 'buttongroup' && ctrl.size !== 'large';
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
        } else if (ctrl.type === 'gallery') {
            let itemsMarkup = '';
            if (ctrl.items) {
                ctrl.items.forEach((item, itemIdx) => {
                    itemsMarkup += `
                        <div class="list-item-row gallery-item-row" data-idx="${itemIdx}" style="display: grid; grid-template-columns: 1fr 1fr 1.2fr 24px; gap: 4px; align-items: center; margin-bottom: 4px;">
                            <input type="text" class="prop-gallery-item-id" placeholder="ID" value="${item.id}" style="font-size:10.5px; padding:2px;">
                            <input type="text" class="prop-gallery-item-label" placeholder="레이블 / #HEX색상" value="${item.label}" style="font-size:10.5px; padding:2px;">
                            <input type="text" class="prop-gallery-item-imagemso" placeholder="imageMso / #HEX색상" value="${item.imageMso || ''}" style="font-size:10.5px; padding:2px;">
                            <button type="button" class="overlay-btn delete gallery-item-del-btn" style="display:flex; width:22px; height:22px; align-items:center; justify-content:center; padding:0;">✖</button>
                        </div>
                    `;
                });
            }

            extraFields = `
                <div class="form-group">
                    <label>열 개수 (columns)</label>
                    <input type="number" id="prop-ctrl-columns" value="${ctrl.columns || 5}">
                </div>
                <div class="form-group">
                    <label>항목 가로 너비 (itemWidth)</label>
                    <input type="number" id="prop-ctrl-itemwidth" value="${ctrl.itemWidth || 24}">
                </div>
                <div class="form-group">
                    <label>항목 세로 높이 (itemHeight)</label>
                    <input type="number" id="prop-ctrl-itemheight" value="${ctrl.itemHeight || 24}">
                </div>
                <div class="form-group">
                    <label>색상 팔레트 프리셋 빠른 적용</label>
                    <select id="prop-gallery-color-preset" style="width:100%; padding:4px; font-size:11px;">
                        <option value="" selected>-- 테마 선택 --</option>
                        <option value="appointment">기본 10색 (일정 색상)</option>
                        <option value="standard">표준 10색 (무지개색)</option>
                        <option value="office">오피스 테마 20색</option>
                        <option value="rainbow_bright">화려한 40색 그리드</option>
                    </select>
                </div>
                <div class="form-group" style="grid-column: span 2;">
                    <label>갤러리 아이템 항목 관리 (ID / 표시이름 또는 HEX색상 / imageMso 또는 HEX색상)</label>
                    <div class="list-editor">
                        <div class="list-editor-items" id="prop-gallery-items-container">
                            ${itemsMarkup}
                        </div>
                        <button type="button" class="btn-secondary" id="btn-add-gallery-item" style="padding:3px 8px; font-size:10px; align-self:flex-start; margin-top:4px;">+ 항목 추가</button>
                    </div>
                </div>
            `;
        } else if (ctrl.type === 'splitbutton') {
            let itemsMarkup = '';
            if (ctrl.items) {
                ctrl.items.forEach((item, itemIdx) => {
                    itemsMarkup += `
                        <div class="list-item-row split-item-row" data-idx="${itemIdx}" style="display: grid; grid-template-columns: 1fr 1fr 1.2fr 24px; gap: 4px; align-items: center; margin-bottom: 4px;">
                            <input type="text" class="prop-split-item-id" placeholder="ID" value="${item.id}" style="font-size:10.5px; padding:2px;">
                            <input type="text" class="prop-split-item-label" placeholder="레이블" value="${item.label}" style="font-size:10.5px; padding:2px;">
                            <input type="text" class="prop-split-item-action" placeholder="onAction" value="${item.onAction || ''}" style="font-size:10.5px; padding:2px;">
                            <button type="button" class="overlay-btn delete split-item-del-btn" style="display:flex; width:22px; height:22px; align-items:center; justify-content:center; padding:0;">✖</button>
                        </div>
                    `;
                });
            }

            extraFields = `
                <div class="form-group" style="grid-column: span 2;">
                    <label>드롭다운 서브 메뉴 관리 (ID / 레이블명 / 실행 매크로명)</label>
                    <div class="list-editor">
                        <div class="list-editor-items" id="prop-split-items-container">
                            ${itemsMarkup}
                        </div>
                        <button type="button" class="btn-secondary" id="btn-add-split-item" style="padding:3px 8px; font-size:10px; align-self:flex-start; margin-top:4px;">+ 항목 추가</button>
                    </div>
                </div>
            `;
        } else if (ctrl.type === 'dynamicmenu') {
            extraFields = `
                <div class="form-group" style="grid-column: span 2;">
                    <label>동적 메뉴 콘텐츠 생성 함수 (getContent)</label>
                    <input type="text" id="prop-ctrl-getcontent" value="${ctrl.getContent || ctrl.id + '_GetContent'}">
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

        if (ctrl.type === 'gallery') {
            document.getElementById("btn-add-gallery-item").addEventListener("click", () => {
                if (!ctrl.items) ctrl.items = [];
                const nextId = ctrl.items.length + 1;
                ctrl.items.push({ 
                    id: `${ctrl.id}_item${nextId}`, 
                    label: `항목 ${nextId}`, 
                    imageMso: "AppointmentColor1" 
                });
                updateRibbonUI();
            });

            document.querySelectorAll(".gallery-item-del-btn").forEach(btn => {
                btn.addEventListener("click", (e) => {
                    const row = btn.closest(".gallery-item-row");
                    const idx = parseInt(row.getAttribute("data-idx"));
                    ctrl.items.splice(idx, 1);
                    updateRibbonUI();
                });
            });
        }

        if (ctrl.type === 'splitbutton') {
            document.getElementById("btn-add-split-item").addEventListener("click", () => {
                if (!ctrl.items) ctrl.items = [];
                const nextId = ctrl.items.length + 1;
                ctrl.items.push({ 
                    id: `${ctrl.id}_sub${nextId}`, 
                    label: `서브 명령 ${nextId}`, 
                    onAction: `${ctrl.id}_sub${nextId}_Click` 
                });
                updateRibbonUI();
            });

            document.querySelectorAll(".split-item-del-btn").forEach(btn => {
                btn.addEventListener("click", (e) => {
                    const row = btn.closest(".split-item-row");
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
            if (targetId === "prop-ctrl-icon") {
                ctrl.imageMso = e.target.value.trim();
                const previewBox = document.getElementById("prop-icon-preview");
                if (previewBox) {
                    previewBox.innerHTML = getIconHtml(ctrl.imageMso, 18);
                }
            }
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
            if (targetId === "prop-ctrl-columns") ctrl.columns = parseInt(e.target.value, 10) || 5;
            if (targetId === "prop-ctrl-itemwidth") ctrl.itemWidth = parseInt(e.target.value, 10) || 24;
            if (targetId === "prop-ctrl-getcontent") {
                const oldContent = ctrl.getContent || "";
                const newContent = e.target.value;
                ctrl.getContent = newContent;
                if (ctrl.vbaCustomCode && oldContent) {
                    try {
                        const regex = new RegExp('Sub\\s+' + oldContent + '\\b', 'g');
                        ctrl.vbaCustomCode = ctrl.vbaCustomCode.replace(regex, 'Sub ' + newContent);
                    } catch (err) {
                        console.error("VBA rename sync failed", err);
                    }
                }
                const codeArea = document.getElementById("prop-ctrl-vba-code");
                if (codeArea) {
                    codeArea.value = ctrl.vbaCustomCode || "";
                }
            }
            if (targetId === "prop-gallery-color-preset") {
                const presetVal = e.target.value;
                if (presetVal) {
                    applyColorPresetToGallery(ctrl, presetVal);
                    updateRibbonUI();
                }
                return;
            }
            if (targetId === "prop-ctrl-itemheight") ctrl.itemHeight = parseInt(e.target.value, 10) || 24;
            
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

            if (e.target.classList.contains("prop-gallery-item-id") || e.target.classList.contains("prop-gallery-item-label") || e.target.classList.contains("prop-gallery-item-imagemso")) {
                const row = e.target.closest(".gallery-item-row");
                const itemIdx = parseInt(row.getAttribute("data-idx"));
                
                if (e.target.classList.contains("prop-gallery-item-id")) {
                    ctrl.items[itemIdx].id = e.target.value;
                } else if (e.target.classList.contains("prop-gallery-item-label")) {
                    ctrl.items[itemIdx].label = e.target.value;
                } else if (e.target.classList.contains("prop-gallery-item-imagemso")) {
                    ctrl.items[itemIdx].imageMso = e.target.value;
                }
            }

            if (e.target.classList.contains("prop-split-item-id") || e.target.classList.contains("prop-split-item-label") || e.target.classList.contains("prop-split-item-action")) {
                const row = e.target.closest(".split-item-row");
                const itemIdx = parseInt(row.getAttribute("data-idx"));
                
                if (e.target.classList.contains("prop-split-item-id")) {
                    ctrl.items[itemIdx].id = e.target.value;
                } else if (e.target.classList.contains("prop-split-item-label")) {
                    ctrl.items[itemIdx].label = e.target.value;
                } else if (e.target.classList.contains("prop-split-item-action")) {
                    ctrl.items[itemIdx].onAction = e.target.value;
                }
            }
        }

        renderRibbonSimulator();
        compileCodes();
    });
}

function applyColorPresetToGallery(ctrl, presetType) {
    const newId = ctrl.id;
    if (presetType === 'appointment') {
        ctrl.items = [
            { id: `${newId}_colorRed`, label: "빨강", imageMso: "AppointmentColor1" },
            { id: `${newId}_colorOrange`, label: "주황", imageMso: "AppointmentColor2" },
            { id: `${newId}_colorYellow`, label: "노랑", imageMso: "AppointmentColor3" },
            { id: `${newId}_colorGreen`, label: "초록", imageMso: "AppointmentColor4" },
            { id: `${newId}_colorBlue`, label: "파랑", imageMso: "AppointmentColor5" },
            { id: `${newId}_colorDarkBlue`, label: "남색", imageMso: "AppointmentColor6" },
            { id: `${newId}_colorPurple`, label: "보라", imageMso: "AppointmentColor7" },
            { id: `${newId}_colorBlack`, label: "검정", imageMso: "AppointmentColor8" },
            { id: `${newId}_colorGray`, label: "회색", imageMso: "AppointmentColor9" },
            { id: `${newId}_colorWhite`, label: "흰색", imageMso: "AppointmentColor0" }
        ];
        ctrl.columns = 5;
    } else if (presetType === 'standard') {
        ctrl.items = [
            { id: `${newId}_red`, label: "#FF0000", imageMso: "#FF0000" },
            { id: `${newId}_orange`, label: "#FFC000", imageMso: "#FFC000" },
            { id: `${newId}_yellow`, label: "#FFFF00", imageMso: "#FFFF00" },
            { id: `${newId}_lightgreen`, label: "#92D050", imageMso: "#92D050" },
            { id: `${newId}_green`, label: "#00B050", imageMso: "#00B050" },
            { id: `${newId}_lightblue`, label: "#00B0F0", imageMso: "#00B0F0" },
            { id: `${newId}_blue`, label: "#0070C0", imageMso: "#0070C0" },
            { id: `${newId}_darkblue`, label: "#002060", imageMso: "#002060" },
            { id: `${newId}_purple`, label: "#7030A0", imageMso: "#7030A0" },
            { id: `${newId}_darkred`, label: "#C00000", imageMso: "#C00000" }
        ];
        ctrl.columns = 5;
    } else if (presetType === 'office') {
        ctrl.items = [
            { id: `${newId}_c1`, label: "#FFFFFF", imageMso: "#FFFFFF" },
            { id: `${newId}_c2`, label: "#000000", imageMso: "#000000" },
            { id: `${newId}_c3`, label: "#E7E6E6", imageMso: "#E7E6E6" },
            { id: `${newId}_c4`, label: "#44546A", imageMso: "#44546A" },
            { id: `${newId}_c5`, label: "#4472C4", imageMso: "#4472C4" },
            { id: `${newId}_c6`, label: "#ED7D31", imageMso: "#ED7D31" },
            { id: `${newId}_c7`, label: "#A5A5A5", imageMso: "#A5A5A5" },
            { id: `${newId}_c8`, label: "#FFC000", imageMso: "#FFC000" },
            { id: `${newId}_c9`, label: "#5B9BD5", imageMso: "#5B9BD5" },
            { id: `${newId}_c10`, label: "#70AD47", imageMso: "#70AD47" },
            { id: `${newId}_c11`, label: "#F2F2F2", imageMso: "#F2F2F2" },
            { id: `${newId}_c12`, label: "#7F7F7F", imageMso: "#7F7F7F" },
            { id: `${newId}_c13`, label: "#D2D2D2", imageMso: "#D2D2D2" },
            { id: `${newId}_c14`, label: "#D9E1F2", imageMso: "#D9E1F2" },
            { id: `${newId}_c15`, label: "#FCE4D6", imageMso: "#FCE4D6" },
            { id: `${newId}_c16`, label: "#FFF2CC", imageMso: "#FFF2CC" },
            { id: `${newId}_c17`, label: "#E2EFDA", imageMso: "#E2EFDA" },
            { id: `${newId}_c18`, label: "#C9DAF8", imageMso: "#C9DAF8" },
            { id: `${newId}_c19`, label: "#F4CCCC", imageMso: "#F4CCCC" },
            { id: `${newId}_c20`, label: "#D0E0E3", imageMso: "#D0E0E3" }
        ];
        ctrl.columns = 10;
    } else if (presetType === 'rainbow_bright') {
        const rainbowColors = [
            "#FF0000", "#FF3300", "#FF6600", "#FF9900", "#FFCC00", "#FFFF00", "#CCFF00", "#99FF00",
            "#66FF00", "#33FF00", "#00FF00", "#00FF33", "#00FF66", "#00FF99", "#00FFCC", "#00FFFF",
            "#00CCFF", "#0099FF", "#0066FF", "#0033FF", "#0000FF", "#3300FF", "#6600FF", "#9900FF",
            "#CC00FF", "#FF00FF", "#FF00CC", "#FF0099", "#FF0066", "#FF0033", "#000000", "#222222",
            "#444444", "#666666", "#888888", "#AAAAAA", "#CCCCCC", "#EEEEEE", "#F5F5F5", "#FFFFFF"
        ];
        ctrl.items = rainbowColors.map((hex, idx) => ({
            id: `${newId}_hex_${idx}`,
            label: hex,
            imageMso: hex
        }));
        ctrl.columns = 8;
    }
}

// --- 8. Real-time Double Code Compiler ---
function compileCodes() {
    const xml = compileCustomUiXml();
    const xmlBox = document.getElementById("code-xml-box");
    if (xmlBox) xmlBox.value = xml;

    const vba = compileVbaCallbacks();
    const vbaBox = document.getElementById("code-vba-box");
    if (vbaBox) vbaBox.value = vba;
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
    } else if (ctrl.type === 'gallery') {
        const colsAttr = ctrl.columns ? ` columns="${ctrl.columns}"` : '';
        const wAttr = ctrl.itemWidth ? ` itemWidth="${ctrl.itemWidth}"` : '';
        const hAttr = ctrl.itemHeight ? ` itemHeight="${ctrl.itemHeight}"` : '';
        xml += `${indent}<gallery id="${ctrl.id}" label="${escapeXml(ctrl.label)}"${sizeAttr}${imageAttr}${colsAttr}${wAttr}${hAttr} onAction="${ctrl.onAction || ''}">\n`;
        if (ctrl.items) {
            ctrl.items.forEach(item => {
                const itemImgAttr = item.imageMso ? ` imageMso="${item.imageMso}"` : '';
                xml += `${indent}  <item id="${item.id}" label="${escapeXml(item.label)}"${itemImgAttr} />\n`;
            });
        }
        xml += `${indent}</gallery>\n`;
    } else if (ctrl.type === 'splitbutton') {
        xml += `${indent}<splitButton id="${ctrl.id}"${sizeAttr}>\n`;
        xml += `${indent}  <button id="${ctrl.id}_btn" label="${escapeXml(ctrl.label)}"${imageAttr} onAction="${ctrl.onAction || ''}" />\n`;
        xml += `${indent}  <menu id="${ctrl.id}_menu" label="${escapeXml(ctrl.label)} 메뉴">\n`;
        if (ctrl.items) {
            ctrl.items.forEach(item => {
                xml += `${indent}    <button id="${item.id}" label="${escapeXml(item.label)}" onAction="${item.onAction || ''}" />\n`;
            });
        }
        xml += `${indent}  </menu>\n`;
        xml += `${indent}</splitButton>\n`;
    } else if (ctrl.type === 'dynamicmenu') {
        const getContentAttr = ctrl.getContent ? ` getContent="${ctrl.getContent}"` : '';
        xml += `${indent}<dynamicMenu id="${ctrl.id}" label="${escapeXml(ctrl.label)}"${sizeAttr}${imageAttr}${getContentAttr} />\n`;
    } else {
        // standard button
        xml += `${indent}<button id="${ctrl.id}" label="${escapeXml(ctrl.label)}"${sizeAttr}${imageAttr} onAction="${ctrl.onAction}" />\n`;
    }
    return xml;
}

function compileCustomUiXml() {
    const isOffice2007 = activeSchemaVersion === 'customUI';
    const schemaNamespace = isOffice2007 
        ? "http://schemas.microsoft.com/office/2006/01/customui"
        : "http://schemas.microsoft.com/office/2009/07/customui";
    
    let xml = `<!-- Office ${isOffice2007 ? '2007' : '2010+'} 호환 CustomUI XML 스키마 -->\n`;
    xml += `<customUI xmlns="${schemaNamespace}">\n`;
    
    // 1. Commands block
    if (ribbonState.commands && ribbonState.commands.length > 0) {
        xml += `  <commands>\n`;
        ribbonState.commands.forEach(cmd => {
            const enabledAttr = cmd.enabled === false ? ' enabled="false"' : '';
            const actionAttr = cmd.onAction ? ` onAction="${cmd.onAction}"` : '';
            xml += `    <command idMso="${cmd.idMso}"${enabledAttr}${actionAttr} />\n`;
        });
        xml += `  </commands>\n`;
    }

    // 2. Ribbon block
    if (ribbonState.tabs && ribbonState.tabs.length > 0) {
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
                    if (group.idMso) {
                        const visibleAttr = group.visible === false ? ' visible="false"' : '';
                        xml += `        <group idMso="${group.idMso}"${visibleAttr} />\n`;
                    } else {
                        xml += `        <group id="${group.id}" label="${escapeXml(group.label)}">\n`;
                        
                        if (group.controls && group.controls.length > 0) {
                            let i = 0;
                            while (i < group.controls.length) {
                                const ctrl = group.controls[i];
                                
                                if (ctrl.type === 'box' || ctrl.type === 'buttongroup') {
                                    const stackLimit = ctrl.type === 'buttongroup' ? 5 : (ctrl.stackLimit || 3);
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
                                    
                                    if (ctrl.type === 'box') {
                                        xml += `          <box id="${ctrl.id}" boxStyle="${ctrl.boxStyle || 'vertical'}">\n`;
                                        children.forEach(child => {
                                            xml += compileSingleControlXml(child, "            ");
                                        });
                                        xml += `          </box>\n`;
                                    } else {
                                        xml += `          <buttonGroup id="${ctrl.id}">\n`;
                                        children.forEach(child => {
                                            xml += compileSingleControlXml(child, "            ");
                                        });
                                        xml += `          </buttonGroup>\n`;
                                    }
                                    
                                    i = j;
                                } else {
                                    xml += compileSingleControlXml(ctrl, "          ");
                                    i++;
                                }
                            }
                        }
                        
                        xml += `        </group>\n`;
                    }
                });
            }
            xml += `      </tab>\n`;
        });
        xml += `    </tabs>\n`;
        xml += `  </ribbon>\n`;
    }

    // 3. Backstage block
    if (ribbonState.backstage) {
        xml += `  <backstage>\n`;
        if (ribbonState.backstage.tabs) {
            ribbonState.backstage.tabs.forEach(tab => {
                xml += `    <tab id="${tab.id}" label="${escapeXml(tab.label)}">\n`;
                if (tab.firstColumn) {
                    xml += `      <firstColumn>\n`;
                    tab.firstColumn.groups.forEach(tg => {
                        xml += `        <taskGroup id="${tg.id}" label="${escapeXml(tg.label)}">\n`;
                        tg.categories.forEach(cat => {
                            xml += `          <category id="${cat.id}" label="${escapeXml(cat.label)}">\n`;
                            cat.tasks.forEach(task => {
                                const actionAttr = task.onAction ? ` onAction="${task.onAction}"` : '';
                                const imageAttr = task.imageMso ? ` imageMso="${task.imageMso}"` : '';
                                xml += `            <task id="${task.id}" label="${escapeXml(task.label)}"${imageAttr}${actionAttr}/>\n`;
                            });
                            xml += `          </category>\n`;
                        });
                        xml += `        </taskGroup>\n`;
                    });
                    xml += `      </firstColumn>\n`;
                }
                if (tab.secondColumn) {
                    xml += `      <secondColumn>\n`;
                    tab.secondColumn.groups.forEach(tg => {
                        xml += `        <taskGroup id="${tg.id}" label="${escapeXml(tg.label)}">\n`;
                        tg.categories.forEach(cat => {
                            xml += `          <category id="${cat.id}" label="${escapeXml(cat.label)}">\n`;
                            cat.tasks.forEach(task => {
                                const actionAttr = task.onAction ? ` onAction="${task.onAction}"` : '';
                                const imageAttr = task.imageMso ? ` imageMso="${task.imageMso}"` : '';
                                xml += `            <task id="${task.id}" label="${escapeXml(task.label)}"${imageAttr}${actionAttr}/>\n`;
                            });
                            xml += `          </category>\n`;
                        });
                        xml += `        </taskGroup>\n`;
                    });
                    xml += `      </secondColumn>\n`;
                }
                xml += `    </tab>\n`;
            });
        }
        if (ribbonState.backstage.buttons) {
            ribbonState.backstage.buttons.forEach(btn => {
                const actionAttr = btn.onAction ? ` onAction="${btn.onAction}"` : '';
                const imageAttr = btn.imageMso ? ` imageMso="${btn.imageMso}"` : '';
                xml += `    <button id="${btn.id}" label="${escapeXml(btn.label)}"${imageAttr}${actionAttr} />\n`;
            });
        }
        xml += `  </backstage>\n`;
    }

    xml += `</customUI>`;
    return xml;
}

function compileVbaCallbacks() {
    let vba = `' =======================================================\n`;
    vba += `'  EXCEL RIBBON CALLBACK VBA CODE (자동 생성된 소스코드)\n`;
    vba += `'  [가이드] 다운로드한 엑셀에 Alt+F11로 모듈 삽입 후 아래 전체 복사 붙여넣기\n`;
    vba += `' =======================================================\n\n`;

    let generatedActions = new Set();

    // 1. Core Ribbon tabs
    if (ribbonState.tabs) {
        ribbonState.tabs.forEach(tab => {
            if (tab.groups) {
                tab.groups.forEach(group => {
                    if (group.controls) {
                        group.controls.forEach(ctrl => {
                            if (ctrl.type === 'menu' || ctrl.type === 'splitbutton') {
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
                                if (ctrl.type === 'menu') return;
                            }

                            if (ctrl.type === 'separator' || ctrl.type === 'labelcontrol' || ctrl.type === 'box' || ctrl.type === 'buttongroup') return;
                            
                            if (ctrl.type === 'dynamicmenu') {
                                const cbName = ctrl.getContent || (ctrl.id + "_GetContent");
                                if (generatedActions.has(cbName)) return;
                                generatedActions.add(cbName);
                                if (ctrl.vbaCustomCode) {
                                    vba += ctrl.vbaCustomCode + "\n\n";
                                } else {
                                    vba += `' ${ctrl.label} (동적 메뉴 콘텐츠 생성 콜백)\n`;
                                    vba += `Sub ${cbName}(control As IRibbonControl, ByRef content)\n`;
                                    vba += `    Dim xml As String\n`;
                                    vba += `    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & _\n`;
                                    vba += `          "  <button id=""dynBtn1"" label=""동적 버튼 1"" imageMso=""HappyFace"" onAction=""DynButton_Click"" />" & _\n`;
                                    vba += `          "  <button id=""dynBtn2"" label=""동적 버튼 2"" imageMso=""Calendar"" onAction=""DynButton_Click"" />" & _\n`;
                                    vba += `          "</menu>"\n`;
                                    vba += `    content = xml\n`;
                                    vba += `End Sub\n\n`;
                                }
                                return;
                            }

                            if (!ctrl.onAction) return;
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
                                } else if (ctrl.type === 'gallery') {
                                    vba += `' ${ctrl.label} (갤러리 항목 선택 시 실행)\n`;
                                    vba += `Sub ${ctrl.onAction}(control As IRibbonControl, id As String, index As Integer)\n`;
                                    vba += `    MsgBox "갤러리 선택 항목: " & id & " (인덱스: " & index & ")", vbInformation, "엑셀 리본 빌더"\n`;
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
    }

    // 2. Repurposed commands callbacks
    if (ribbonState.commands) {
        ribbonState.commands.forEach(cmd => {
            if (cmd.onAction && !generatedActions.has(cmd.onAction)) {
                generatedActions.add(cmd.onAction);
                vba += `' 명령 재정의 (idMso="${cmd.idMso}") 콜백\n`;
                vba += `Sub ${cmd.onAction}(control As IRibbonControl, ByRef cancelDefault As Boolean)\n`;
                vba += `    MsgBox "${cmd.idMso} 명령어가 재정의되어 실행되었습니다.", vbInformation, "엑셀 리본 빌더"\n`;
                vba += `    cancelDefault = True ' 기본 동작 취소 여부\n`;
                vba += `End Sub\n\n`;
            }
        });
    }

    // 3. Backstage callbacks
    if (ribbonState.backstage) {
        if (ribbonState.backstage.tabs) {
            ribbonState.backstage.tabs.forEach(tab => {
                const cols = [tab.firstColumn, tab.secondColumn];
                cols.forEach(col => {
                    if (col && col.groups) {
                        col.groups.forEach(tg => {
                            if (tg.categories) {
                                tg.categories.forEach(cat => {
                                    if (cat.tasks) {
                                        cat.tasks.forEach(task => {
                                            if (task.onAction && !generatedActions.has(task.onAction)) {
                                                generatedActions.add(task.onAction);
                                                vba += `' Backstage 작업 > ${task.label} 콜백\n`;
                                                vba += `Sub ${task.onAction}(control As IRibbonControl)\n`;
                                                vba += `    MsgBox "백스테이지 작업 '${task.label}' 실행됨", vbInformation, "엑셀 리본 빌더"\n`;
                                                vba += `End Sub\n\n`;
                                            }
                                        });
                                    }
                                });
                            }
                        });
                    }
                });
            });
        }
        if (ribbonState.backstage.buttons) {
            ribbonState.backstage.buttons.forEach(btn => {
                if (btn.onAction && !generatedActions.has(btn.onAction)) {
                    generatedActions.add(btn.onAction);
                    vba += `' Backstage 빠른 실행 버튼 > ${btn.label} 콜백\n`;
                    vba += `Sub ${btn.onAction}(control As IRibbonControl)\n`;
                    vba += `    MsgBox "백스테이지 버튼 '${btn.label}' 클릭됨", vbInformation, "엑셀 리본 빌더"\n`;
                    vba += `End Sub\n\n`;
                }
            });
        }
    }

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

    if (typeof OFFICE_MSO_ICONS === 'undefined') {
        grid.innerHTML = `<div style="padding:20px; font-style:italic; color:#888;">OFFICE_MSO_ICONS 데이터가 로드되지 않았습니다.</div>`;
        return;
    }

    const categories = Object.keys(OFFICE_MSO_ICONS);
    const displayCategories = [...categories, "imagemso0", "imagemso1", "imagemso2", "imagemso3", "imagemso4"];

    // Initialize active category state
    if (!m365SelectedLevel1 || !displayCategories.includes(m365SelectedLevel1)) {
        m365SelectedLevel1 = categories[0];
    }

    // Render dynamic category sub-tabs (Level 1)
    displayCategories.forEach(cat => {
        const btn = document.createElement("button");
        btn.className = "sub-tab-btn";
        if (cat === m365SelectedLevel1) btn.classList.add("active");
        btn.textContent = cat;
        
        btn.addEventListener("click", () => {
            m365SelectedLevel1 = cat;
            buildIconPickerGrid();
        });
        subTabsBar.appendChild(btn);
    });

    // Search or normal grid browsing
    if (query) {
        let resultsCount = 0;
        const renderedKeys = new Set();

        // 1. Search in curated icons first for custom SVG paths
        categories.forEach(cat => {
            const iconsMap = OFFICE_MSO_ICONS[cat];
            Object.keys(iconsMap).forEach(key => {
                const info = iconsMap[key];
                const label = info.label;
                if (key.toLowerCase().includes(query) || label.toLowerCase().includes(query) || (info.desc && info.desc.toLowerCase().includes(query))) {
                    if (resultsCount < 200) {
                        const card = createIconCard(key, `${label} (${key})`, key);
                        grid.appendChild(card);
                        renderedKeys.add(key);
                        resultsCount++;
                    }
                }
            });
        });

        // 2. Search in all 3,244 standard imageMso names
        if (typeof ALL_IMAGE_MSO_LIST !== 'undefined') {
            for (let i = 0; i < ALL_IMAGE_MSO_LIST.length; i++) {
                const key = ALL_IMAGE_MSO_LIST[i];
                if (renderedKeys.has(key)) continue;

                if (key.toLowerCase().includes(query)) {
                    if (resultsCount < 200) {
                        const card = createIconCard(key, key, key);
                        grid.appendChild(card);
                        renderedKeys.add(key);
                        resultsCount++;
                    } else {
                        break;
                    }
                }
            }
        }

        if (resultsCount === 0) {
            grid.innerHTML = `<div style="padding:20px; font-style:italic; color:#888;">검색 결과가 없습니다. 우측 직접 입력을 시도해 보세요.</div>`;
        }
    } else {
        // Normal browsing mode
        if (m365SelectedLevel1.startsWith("imagemso")) {
            if (typeof ALL_IMAGE_MSO_LIST !== 'undefined') {
                const pageNum = parseInt(m365SelectedLevel1.replace("imagemso", ""), 10);
                const startIdx = pageNum * 650;
                const endIdx = Math.min(startIdx + 650, ALL_IMAGE_MSO_LIST.length);
                
                for (let i = startIdx; i < endIdx; i++) {
                    const key = ALL_IMAGE_MSO_LIST[i];
                    const card = createIconCard(key, key, key);
                    grid.appendChild(card);
                }

                // Add a notice card at the end
                const notice = document.createElement("div");
                notice.style.gridColumn = "1 / -1";
                notice.style.textAlign = "center";
                notice.style.padding = "20px";
                notice.style.color = "#666";
                notice.style.fontSize = "13px";
                notice.style.fontStyle = "italic";
                notice.innerHTML = `현재 탭은 ${startIdx + 1}번째부터 ${endIdx}번째까지의 내장 아이콘 ${endIdx - startIdx}개를 표시하고 있습니다. (검색창을 통해 3,200개 이상 전체 검색도 가능)`;
                grid.appendChild(notice);
            }
        } else {
            // Render icons for currently selected category
            const iconsMap = OFFICE_MSO_ICONS[m365SelectedLevel1] || {};
            Object.keys(iconsMap).forEach(key => {
                const info = iconsMap[key];
                const card = createIconCard(key, `${info.label} (${key})`, key);
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
    const manualInput = document.getElementById("modal-manual-icon-input");
    if (manualInput) manualInput.value = "";
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

        const blob = await ExcelZipGenerator.generate(xml, type, activeSchemaVersion);
        
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

// ==========================================
// --- RibbonX Editor Features Extension ---
// ==========================================

const RIBBONX_TEMPLATES = {
    custom_tab: {
        xml: `<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="customTab" label="Custom Tab">
        <group id="customGroup" label="Custom Group">
          <button id="customButton1" label="Button 1" size="large" onAction="onAction1" imageMso="HappyFace" />
          <toggleButton id="customToggleButton1" label="Toggle Button" size="normal" onAction="onToggleAction" />
          <checkBox id="customCheckBox1" label="Checkbox" onAction="onCheckAction" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>`
    },
    excel_custom_tab: {
        xml: `<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="customTab" label="Contoso" insertAfterMso="TabHome">
        <group idMso="GroupClipboard" />
        <group idMso="GroupFont" />
        <group id="customGroup" label="Contoso Tools">
          <button id="customButton1" label="ConBold" size="large" onAction="conBoldSub" imageMso="Bold" />
          <button id="customButton2" label="ConItalic" size="large" onAction="conItalicSub" imageMso="Italic" />
          <button id="customButton3" label="ConUnderline" size="large" onAction="conUnderlineSub" imageMso="Underline" />
        </group>
        <group idMso="GroupEnterDataAlignment" />
        <group idMso="GroupEnterDataNumber" />
        <group idMso="GroupQuickFormatting" />
      </tab>
    </tabs>
  </ribbon>
</customUI>`
    },
    custom_outspace: {
        xml: `<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <backstage>
    <tab id="customTab" label="Custom File Menu">
      <firstColumn>
        <taskGroup id="customTaskGroup" label="Custom Tasks">
          <category id="tgCategory1" label="Category One">
            <task id="task1" label="Task 1" imageMso="FileOpen" onAction="onTask1Click"/>
            <task id="task2" label="Task 2" imageMso="FileSave" onAction="onTask2Click"/>
            <task id="task3" label="Task 3" imageMso="FileSaveAs" onAction="onTask3Click"/>
          </category>
        </taskGroup>
      </firstColumn>
    </tab>
  </backstage>
</customUI>`
    },
    repurpose: {
        xml: `<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <commands>
    <command idMso="Bold" enabled="false" />
    <command idMso="Save" onAction="MySave" />
  </commands>
</customUI>`
    },
    word_group_insert: {
        xml: `<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab idMso="TabInsert">
        <group id="customGroup" label="Custom Tools" insertAfterMso="GroupIllustrations">
          <button id="customButton1" label="Insert Tool" size="large" onAction="onInsertToolClick" imageMso="TableInsert" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>`
    }
};

function xmlToState(xmlText) {
    try {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlText, "text/xml");
        
        const parseError = xmlDoc.getElementsByTagName("parsererror");
        if (parseError.length > 0) {
            return { error: "XML 구문 오류가 발견되었습니다:\n" + parseError[0].textContent };
        }

        const root = xmlDoc.documentElement;
        if (!root || (root.nodeName.replace(/.*:/, '') !== 'customUI')) {
            return { error: "유효하지 않은 루트 노드입니다. 루트 태그는 <customUI>여야 합니다." };
        }

        const ns = root.getAttribute("xmlns") || "";
        let parsedSchema = 'customUI14';
        if (ns.includes("2006/01/customui")) {
            parsedSchema = 'customUI';
        }

        let tabs = [];
        let backstage = null;
        let commands = [];
        let hasNonSimulatable = false;

        // Context menus, sharedControls, etc. are still non-simulatable.
        if (xmlDoc.getElementsByTagName("contextMenus").length > 0) {
            hasNonSimulatable = true;
        }

        // Parse commands
        const commandsNode = xmlDoc.getElementsByTagName("commands")[0];
        if (commandsNode) {
            const cmdNodes = commandsNode.childNodes;
            for (let i = 0; i < cmdNodes.length; i++) {
                const child = cmdNodes[i];
                if (child.nodeType !== 1) continue;
                const tagName = child.nodeName.replace(/.*:/, '').toLowerCase();
                if (tagName === 'command') {
                    commands.push({
                        idMso: child.getAttribute("idMso") || "",
                        enabled: child.getAttribute("enabled") !== 'false',
                        onAction: child.getAttribute("onAction") || ""
                    });
                }
            }
        }

        // Parse backstage
        const backstageNode = xmlDoc.getElementsByTagName("backstage")[0];
        if (backstageNode) {
            backstage = {
                tabs: [],
                buttons: []
            };
            const children = backstageNode.childNodes;
            for (let i = 0; i < children.length; i++) {
                const child = children[i];
                if (child.nodeType !== 1) continue;
                const tagName = child.nodeName.replace(/.*:/, '').toLowerCase();
                if (tagName === 'tab') {
                    let tab = {
                        id: child.getAttribute("id") || child.getAttribute("idMso") || "backstage_tab_" + i,
                        label: child.getAttribute("label") || child.getAttribute("idMso") || "Tab",
                        visible: child.getAttribute("visible") !== 'false',
                        firstColumn: null,
                        secondColumn: null
                    };
                    const colNodes = child.childNodes;
                    for (let j = 0; j < colNodes.length; j++) {
                        const col = colNodes[j];
                        if (col.nodeType !== 1) continue;
                        const colName = col.nodeName.replace(/.*:/, '').toLowerCase();
                        if (colName === 'firstcolumn' || colName === 'secondcolumn') {
                            let colData = { groups: [] };
                            const grpNodes = col.childNodes;
                            for (let k = 0; k < grpNodes.length; k++) {
                                const grp = grpNodes[k];
                                if (grp.nodeType !== 1) continue;
                                const grpName = grp.nodeName.replace(/.*:/, '').toLowerCase();
                                if (grpName === 'taskgroup') {
                                    let taskGroup = {
                                        id: grp.getAttribute("id") || "taskgroup_" + k,
                                        label: grp.getAttribute("label") || "",
                                        categories: []
                                    };
                                    const catNodes = grp.childNodes;
                                    for (let l = 0; l < catNodes.length; l++) {
                                        const cat = catNodes[l];
                                        if (cat.nodeType !== 1) continue;
                                        const catName = cat.nodeName.replace(/.*:/, '').toLowerCase();
                                        if (catName === 'category') {
                                            let category = {
                                                id: cat.getAttribute("id") || "cat_" + l,
                                                label: cat.getAttribute("label") || "",
                                                tasks: []
                                            };
                                            const tskNodes = cat.childNodes;
                                            for (let m = 0; m < tskNodes.length; m++) {
                                                const tsk = tskNodes[m];
                                                if (tsk.nodeType !== 1) continue;
                                                const tskName = tsk.nodeName.replace(/.*:/, '').toLowerCase();
                                                if (tskName === 'task') {
                                                    category.tasks.push({
                                                        id: tsk.getAttribute("id") || "task_" + m,
                                                        label: tsk.getAttribute("label") || "",
                                                        imageMso: tsk.getAttribute("imageMso") || "",
                                                        onAction: tsk.getAttribute("onAction") || ""
                                                    });
                                                }
                                            }
                                            taskGroup.categories.push(category);
                                        }
                                    }
                                    colData.groups.push(taskGroup);
                                }
                            }
                            tab[colName === 'firstcolumn' ? 'firstColumn' : 'secondColumn'] = colData;
                        }
                    }
                    backstage.tabs.push(tab);
                } else if (tagName === 'button' || tagName === 'fastcommand') {
                    backstage.buttons.push({
                        id: child.getAttribute("id") || child.getAttribute("idMso") || "fast_btn_" + i,
                        label: child.getAttribute("label") || child.getAttribute("idMso") || "Command",
                        imageMso: child.getAttribute("imageMso") || "",
                        onAction: child.getAttribute("onAction") || ""
                    });
                }
            }
        }

        const tabNodes = xmlDoc.getElementsByTagName("tab");
        for (let t = 0; t < tabNodes.length; t++) {
            const tabNode = tabNodes[t];
            
            let parent = tabNode.parentNode;
            while (parent && parent.nodeName.replace(/.*:/, '') !== 'tabs' && parent.nodeName.replace(/.*:/, '') !== 'customUI') {
                parent = parent.parentNode;
            }
            if (parent && parent.nodeName.replace(/.*:/, '') !== 'tabs') {
                continue; 
            }

            const isStandard = tabNode.hasAttribute("idMso");
            let tab = {
                id: tabNode.getAttribute("id") || (isStandard ? tabNode.getAttribute("idMso") : "tab_" + t),
                label: tabNode.getAttribute("label") || tabNode.getAttribute("idMso") || "Tab",
                visible: tabNode.getAttribute("visible") !== 'false',
                isStandardTab: isStandard,
                idMso: tabNode.getAttribute("idMso") || "",
                positionType: tabNode.hasAttribute("insertBeforeMso") ? "before" : (tabNode.hasAttribute("insertAfterMso") ? "after" : ""),
                positionTarget: tabNode.getAttribute("insertBeforeMso") || tabNode.getAttribute("insertAfterMso") || "",
                groups: []
            };

            const groupNodes = tabNode.getElementsByTagName("group");
            for (let g = 0; g < groupNodes.length; g++) {
                const groupNode = groupNodes[g];
                if (groupNode.parentNode !== tabNode) continue;

                let group = {
                    id: groupNode.getAttribute("id") || (groupNode.hasAttribute("idMso") ? "" : "group_" + g),
                    idMso: groupNode.getAttribute("idMso") || "",
                    label: groupNode.getAttribute("label") || groupNode.getAttribute("idMso") || "Group",
                    visible: groupNode.getAttribute("visible") !== 'false',
                    controls: []
                };

                const childNodes = groupNode.childNodes;
                for (let c = 0; c < childNodes.length; c++) {
                    const child = childNodes[c];
                    if (child.nodeType !== 1) continue; 
                    parseAndFlattenControlNodes(child, group.controls);
                }

                tab.groups.push(group);
            }

            tabs.push(tab);
        }

        return { tabs, backstage, commands, schema: parsedSchema, hasNonSimulatable };
    } catch (err) {
        return { error: "XML 분석 중 예외가 발생했습니다: " + err.message };
    }
}

function parseAndFlattenControlNodes(node, flatControls) {
    if (node.nodeType !== 1) return;
    const type = node.nodeName.replace(/.*:/, '').toLowerCase();
    
    if (['box', 'buttongroup'].includes(type)) {
        let containerControl = {
            id: node.getAttribute("id") || type + "_" + Math.random().toString(36).substr(2, 5),
            type: type,
            label: node.getAttribute("label") || "",
            enabled: node.getAttribute("enabled") !== 'false',
            visible: node.getAttribute("visible") !== 'false'
        };
        
        if (type === 'box') {
            containerControl.boxStyle = node.getAttribute("boxStyle") || "vertical";
        }
        
        flatControls.push(containerControl);
        
        const childNodes = node.childNodes;
        let childCount = 0;
        for (let c = 0; c < childNodes.length; c++) {
            const child = childNodes[c];
            if (child.nodeType !== 1) continue;
            parseAndFlattenControlNodes(child, flatControls);
            childCount++;
        }
        containerControl.stackLimit = childCount;
    } else {
        let control = parseControlNode(node);
        if (control) {
            flatControls.push(control);
        }
    }
}

function parseControlNode(node) {
    const type = node.nodeName.replace(/.*:/, '').toLowerCase();
    
    if (!['button', 'togglebutton', 'checkbox', 'editbox', 'combobox', 'dropdown', 'separator', 'labelcontrol', 'menu', 'box', 'buttongroup', 'gallery', 'splitbutton', 'dynamicmenu'].includes(type)) {
        return null;
    }

    if (type === 'separator') {
        return {
            id: node.getAttribute("id") || "sep_" + Math.random().toString(36).substr(2, 5),
            type: 'separator'
        };
    }

    let control = {
        id: node.getAttribute("id") || type + "_" + Math.random().toString(36).substr(2, 5),
        type: type,
        label: node.getAttribute("label") || "",
        enabled: node.getAttribute("enabled") !== 'false',
        visible: node.getAttribute("visible") !== 'false'
    };

    if (['button', 'togglebutton', 'menu', 'gallery', 'splitbutton', 'dynamicmenu'].includes(type)) {
        control.size = node.getAttribute("size") || "normal";
        control.imageMso = node.getAttribute("imageMso") || "";
        if (type !== 'dynamicmenu') {
            control.onAction = node.getAttribute("onAction") || "";
        } else {
            control.getContent = node.getAttribute("getContent") || "";
        }
    }

    if (type === 'togglebutton') {
        control.checked = node.getAttribute("getPressed") === "true"; 
        control.toggleGroup = node.getAttribute("toggleGroup") || ""; 
    }

    if (type === 'checkbox') {
        control.onAction = node.getAttribute("onAction") || "";
        control.checked = node.getAttribute("getPressed") === "true";
    }

    if (type === 'editbox') {
        control.onAction = node.getAttribute("onChange") || "";
        control.text = node.getAttribute("getText") || "";
    }

    if (['combobox', 'dropdown'].includes(type)) {
        control.onAction = node.getAttribute("onChange") || node.getAttribute("onAction") || "";
        control.text = node.getAttribute("getText") || "";
        control.items = [];
        
        const itemNodes = node.getElementsByTagName("item");
        for (let i = 0; i < itemNodes.length; i++) {
            control.items.push({
                id: itemNodes[i].getAttribute("id") || "item_" + i,
                label: itemNodes[i].getAttribute("label") || ""
            });
        }
    }

    if (type === 'gallery') {
        control.onAction = node.getAttribute("onAction") || "";
        control.columns = parseInt(node.getAttribute("columns"), 10) || 5;
        control.itemWidth = parseInt(node.getAttribute("itemWidth"), 10) || 24;
        control.itemHeight = parseInt(node.getAttribute("itemHeight"), 10) || 24;
        control.items = [];
        
        const itemNodes = node.getElementsByTagName("item");
        for (let i = 0; i < itemNodes.length; i++) {
            control.items.push({
                id: itemNodes[i].getAttribute("id") || "item_" + i,
                label: itemNodes[i].getAttribute("label") || "",
                imageMso: itemNodes[i].getAttribute("imageMso") || ""
            });
        }
    }

    if (type === 'splitbutton') {
        control.items = [];
        control.size = node.getAttribute("size") || "large";
        
        const btnNode = node.querySelector("button, toggleButton");
        if (btnNode) {
            control.label = btnNode.getAttribute("label") || "";
            control.imageMso = btnNode.getAttribute("imageMso") || "";
            control.onAction = btnNode.getAttribute("onAction") || "";
        }
        
        const menuNode = node.querySelector("menu");
        if (menuNode) {
            const menuBtns = menuNode.querySelectorAll("button, toggleButton");
            for (let b = 0; b < menuBtns.length; b++) {
                control.items.push({
                    id: menuBtns[b].getAttribute("id") || "sub_" + b,
                    label: menuBtns[b].getAttribute("label") || "",
                    onAction: menuBtns[b].getAttribute("onAction") || ""
                });
            }
        }
    }

    if (type === 'menu') {
        control.items = [];
        const btnNodes = node.getElementsByTagName("button");
        for (let b = 0; b < btnNodes.length; b++) {
            control.items.push({
                id: btnNodes[b].getAttribute("id") || "sub_" + b,
                label: btnNodes[b].getAttribute("label") || "",
                onAction: btnNodes[b].getAttribute("onAction") || ""
            });
        }
    }

    if (type === 'box') {
        control.boxStyle = node.getAttribute("boxStyle") || "vertical";
    }

    return control;
}

function validateXmlText(xmlText) {
    try {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlText, "text/xml");
        const parseError = xmlDoc.getElementsByTagName("parsererror");
        
        if (parseError.length > 0) {
            return {
                valid: false,
                message: "XML 문법 오류: " + parseError[0].textContent
            };
        }
        
        const root = xmlDoc.documentElement;
        if (!root || (root.nodeName.replace(/.*:/, '') !== 'customUI')) {
            return {
                valid: false,
                message: "XML 구조 오류: 최상위 루트 태그는 <customUI>여야 합니다."
            };
        }

        const ns = root.getAttribute("xmlns") || "";
        if (!ns) {
            return {
                valid: false,
                message: "XML 구조 오류: xmlns 네임스페이스 선언이 없습니다."
            };
        }

        if (ns !== "http://schemas.microsoft.com/office/2009/07/customui" && 
            ns !== "http://schemas.microsoft.com/office/2006/01/customui") {
            return {
                valid: false,
                message: "XML 경고: 알 수 없는 xmlns 네임스페이스입니다. Excel Custom UI 표준 네임스페이스(2007 또는 2010+)를 사용해 주세요."
            };
        }

        const allNodes = xmlDoc.querySelectorAll("[id]");
        const ids = new Set();
        for (let i = 0; i < allNodes.length; i++) {
            const id = allNodes[i].getAttribute("id");
            if (ids.has(id)) {
                return {
                    valid: false,
                    message: `XML 구조 오류: 중복된 ID가 발견되었습니다: '${id}'. 모든 컨트롤 ID는 고유해야 합니다.`
                };
            }
            ids.add(id);
        }

        return {
            valid: true,
            message: "XML 유효성 검사 통과: 문법이 올바르고 Fluent UI 기본 규칙에 부합합니다!"
        };
    } catch (err) {
        return {
            valid: false,
            message: "오류 발생: " + err.message
        };
    }
}

function showValidationBanner(success, message) {
    const banner = document.getElementById("xml-validation-banner");
    const msg = document.getElementById("xml-validation-msg");
    if (!banner || !msg) return;

    banner.style.display = "flex";
    msg.textContent = message;
    
    if (success) {
        banner.className = "validation-banner success";
    } else {
        banner.className = "validation-banner error";
    }
}

function renderDocumentExplorer() {
    const container = document.getElementById("document-explorer-tree");
    if (!container) return;

    const uiFile = activeSchemaVersion === 'customUI' ? 'customUI.xml' : 'customUI14.xml';

    const tree = [
        { id: 'root', label: 'CustomWorkbook.xlsm (통합 문서)', icon: '📁', depth: 0 },
        { id: 'rels_folder', label: '_rels (관계 정보 폴더)', icon: '📁', depth: 1 },
        { id: 'rels', label: '.rels (전체 패키지 관계 파일)', icon: '📄', depth: 2 },
        { id: 'customui_folder', label: 'customUI (리본 메뉴 폴더)', icon: '📁', depth: 1 },
        { id: 'customui_file', label: uiFile + ' (사용자 정의 UI)', icon: '📄', depth: 2 },
        { id: 'xl_folder', label: 'xl (엑셀 데이터 폴더)', icon: '📁', depth: 1 },
        { id: 'xl_rels_folder', label: '_rels', icon: '📁', depth: 2 },
        { id: 'xl_rels', label: 'workbook.xml.rels', icon: '📄', depth: 3 },
        { id: 'workbook', label: 'workbook.xml (통합문서 구성)', icon: '📄', depth: 2 },
        { id: 'worksheets_folder', label: 'worksheets (워크시트 폴더)', icon: '📁', depth: 2 },
        { id: 'sheet1', label: 'sheet1.xml (시트 데이터)', icon: '📄', depth: 3 },
        { id: 'styles', label: 'styles.xml (시트 스타일)', icon: '📄', depth: 2 }
    ];

    container.innerHTML = "";
    tree.forEach(node => {
        const div = document.createElement("div");
        div.className = "tree-node";
        if (node.depth > 0) {
            div.classList.add("tree-indent-" + node.depth);
        }
        if (node.id === activeExplorerNode) {
            div.classList.add("active");
        }

        const iconSpan = document.createElement("span");
        iconSpan.className = "tree-node-icon";
        iconSpan.textContent = node.icon;

        const labelSpan = document.createElement("span");
        labelSpan.className = "tree-node-label";
        labelSpan.textContent = node.label;

        div.appendChild(iconSpan);
        div.appendChild(labelSpan);

        div.addEventListener("click", () => {
            activeExplorerNode = node.id;
            renderDocumentExplorer();
            handleExplorerNodeSelection(node.id);
        });

        container.appendChild(div);
    });
}

function handleExplorerNodeSelection(nodeId) {
    const xmlTab = document.querySelector(".tab-btn[data-tab='panel-xml']");
    const xmlBox = document.getElementById("code-xml-box");
    
    if (nodeId === 'customui_file') {
        xmlBox.readOnly = false;
        xmlBox.placeholder = "여기에 CustomUI XML 코드를 직접 수정하거나 작성한 후 [시뮬레이터 적용]을 클릭하세요.";
        compileCodes();
        if (xmlTab) xmlTab.click();
    } else {
        xmlBox.readOnly = true;
        xmlBox.value = getVirtualFileContent(nodeId);
        xmlBox.placeholder = "이 파일은 OpenXML 패키지의 읽기 전용 구조 파트입니다. 편집은 customUI XML만 가능합니다.";
        if (xmlTab) xmlTab.click();
    }
}

function getVirtualFileContent(nodeId) {
    const uiFile = activeSchemaVersion === 'customUI' ? 'customUI.xml' : 'customUI14.xml';
    const relsType = activeSchemaVersion === 'customUI'
        ? 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility'
        : 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility';

    switch (nodeId) {
        case 'rels':
            return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\n  <Relationship Id="rIdCustomUI" Type="${relsType}" Target="customUI/${uiFile}"/>\n</Relationships>`;
        
        case 'xl_rels':
            return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>\n  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n</Relationships>`;

        case 'workbook':
            return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n  <fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9302"/>\n  <workbookPr defaultThemeVersion="124226"/>\n  <bookViews>\n    <workbookView xWindow="0" yWindow="0" windowWidth="15000" windowHeight="10000"/>\n  </bookViews>\n  <sheets>\n    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>\n  </sheets>\n  <calcPr calcId="145621"/>\n</workbook>`;

        case 'sheet1':
            return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n  <dimension ref="A1"/>\n  <sheetViews>\n    <sheetView tabSelected="1" workbookViewId="0"/>\n  </sheetViews>\n  <sheetFormatPr defaultRowHeight="15"/>\n  <sheetData/>\n  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>\n</worksheet>`;

        case 'styles':
            return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n  <fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font></fonts>\n  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>\n  <borders count="1"><border><left/><right/><top/><bottom/></border></borders>\n  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>\n  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>\n  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>\n  <dxfs count="0"/>\n  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>\n</styleSheet>`;

        default:
            return ``;
    }
}
