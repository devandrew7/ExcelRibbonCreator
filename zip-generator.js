/**
 * Excel Custom Ribbon Creator - Client-side Excel ZIP Packaging Engine
 * Uses JSZip to generate standard OpenXML structures for .xlsm and .xlam files
 * with customUI injected.
 */

class ExcelZipGenerator {
    /**
     * Builds and packages a valid Excel Ribbon workbook ZIP structure.
     * @param {string} customUiXml The full generated customUI XML string.
     * @param {string} fileType The output file type: 'xlsm' (workbook) or 'xlam' (addin).
     * @returns {Promise<Blob>} The generated file binary Blob.
     */
    static async generate(customUiXml, fileType = 'xlsm') {
        if (typeof JSZip === 'undefined') {
            throw new Error("JSZip 라이브러리가 로드되지 않았습니다. 인터넷 연결을 확인해주세요.");
        }

        const zip = new JSZip();

        // 1. Determine Content Types and MIME settings
        const isAddin = fileType === 'xlam';
        const workbookContentType = isAddin 
            ? "application/vnd.ms-excel.addin.macroEnabled.main+xml" 
            : "application/vnd.ms-excel.sheet.macroEnabled.main+xml";

        // 2. [Content_Types].xml
        const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="${workbookContentType}"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/customUI/customUI14.xml" ContentType="application/xml"/>
</Types>`;
        zip.file("[Content_Types].xml", contentTypesXml);

        // 3. _rels/.rels
        const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rIdCustomUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>
</Relationships>`;
        zip.file("_rels/.rels", relsXml);

        // 4. xl/workbook.xml
        // For addins, they run in hidden mode by default, standard workbook xml is perfectly sufficient.
        const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9302"/>
  <workbookPr defaultThemeVersion="124226"/>
  <bookViews>
    <workbookView xWindow="0" yWindow="0" windowWidth="15000" windowHeight="10000"/>
  </bookViews>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <calcPr calcId="145621"/>
</workbook>`;
        zip.file("xl/workbook.xml", workbookXml);

        // 5. xl/_rels/workbook.xml.rels
        const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
        zip.file("xl/_rels/workbook.xml.rels", workbookRelsXml);

        // 6. xl/worksheets/sheet1.xml
        const sheet1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData/>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`;
        zip.file("xl/worksheets/sheet1.xml", sheet1Xml);

        // 7. xl/styles.xml
        const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>`;
        zip.file("xl/styles.xml", stylesXml);

        // 8. customUI/customUI14.xml
        zip.file("customUI/customUI14.xml", customUiXml);

        // Generate the ZIP file as a binary blob
        return await zip.generateAsync({ type: "blob" });
    }
}

// Export modules if running in Node/module context, otherwise bind to window
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { ExcelZipGenerator };
} else {
    window.ExcelZipGenerator = ExcelZipGenerator;
}
