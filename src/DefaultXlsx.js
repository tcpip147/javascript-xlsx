
export default class DefaultXlsx {

    static document = {
        "_rels/.rels": `
            <?xml version="1.0" encoding="UTF-8" standalone="no"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Target="xl/workbook.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>
                <Relationship Id="rId2" Target="docProps/app.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"/>
                <Relationship Id="rId3" Target="docProps/core.xml" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"/>
            </Relationships>
            `,
        "docProps/app.xml": `
            <?xml version="1.0" encoding="UTF-8"?>
            <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
                <Application>javascript-xlsx</Application>
            </Properties>
            `,
        "docProps/core.xml": `
            <?xml version="1.0" encoding="UTF-8" standalone="no"?>
            <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                <dcterms:created xsi:type="dcterms:W3CDTF">2023-03-12T22:55:40Z</dcterms:created>
                <dc:creator>javascript-xlsx</dc:creator>
            </cp:coreProperties>
            `,
        "xl/_rels/workbook.xml.rels": `
            <?xml version="1.0" encoding="UTF-8" standalone="no"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"/>
                <Relationship Id="rId2" Target="styles.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"/>
            </Relationships>
            `,
        "xl/sharedStrings.xml": `
            <?xml version="1.0" encoding="UTF-8"?>
            <sst count="0" uniqueCount="0" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
            `,
        "xl/styles.xml": `
            <?xml version="1.0" encoding="UTF-8"?>
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <numFmts count="0"/>
                <fonts count="1">
                    <font>
                        <sz val="11.0"/>
                        <color indexed="8"/>
                        <name val="Calibri"/>
                        <family val="2"/>
                        <scheme val="minor"/>
                    </font>
                </fonts>
                <fills count="2">
                    <fill>
                        <patternFill patternType="none"/>
                    </fill>
                    <fill>
                        <patternFill patternType="darkGray"/>
                    </fill>
                </fills>
                <borders count="1">
                    <border>
                        <left/>
                        <right/>
                        <top/>
                        <bottom/>
                        <diagonal/>
                    </border>
                </borders>
                <cellStyleXfs count="1">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
                </cellStyleXfs>
                <cellXfs count="1">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
                </cellXfs>
            </styleSheet>
            `,
        "xl/workbook.xml": `
            <?xml version="1.0" encoding="UTF-8"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <workbookPr date1904="false"/>
                <bookViews>
                    <workbookView activeTab="0"/>
                </bookViews>
                <sheets/>
            </workbook>
            `,
        "[Content_Types].xml": `
            <?xml version="1.0" encoding="UTF-8" standalone="no"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default ContentType="application/vnd.openxmlformats-package.relationships+xml" Extension="rels"/>
                <Default ContentType="application/xml" Extension="xml"/>
                <Override ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" PartName="/docProps/app.xml"/>
                <Override ContentType="application/vnd.openxmlformats-package.core-properties+xml" PartName="/docProps/core.xml"/>
                <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" PartName="/xl/sharedStrings.xml"/>
                <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" PartName="/xl/styles.xml"/>
                <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" PartName="/xl/workbook.xml"/>
            </Types>
            `
    };
}