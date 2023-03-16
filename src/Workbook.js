import _ from 'lodash';
import { XMLParser } from "fast-xml-parser";
import dayjs from 'dayjs';
import JSZip from "jszip";
import { saveAs } from 'file-saver';

import DefaultXlsx from './DefaultXlsx';
import Xlsx from './Xlsx';
import IndexedLinkedList from './IndexedLinkedList';
import Sheet from './Sheet';
import CellStyle from './CellStyle';

/** Class Workbook. */
export default class Workbook {

    #maxRId;
    #maxSheetId;
    #maxSheetNameId;

    /**
     * 
     * @param {object} option 
     */
    constructor(option) {
        option = option || {};
        const data = option.data || DefaultXlsx.document;

        this.parser = new XMLParser({
            ignoreAttributes: false,
            allowBooleanAttributes: true
        });
        const document = {};
        for (let filename in data) {
            document[filename] = this.parser.parse(data[filename]);
        }

        this.xlsx = new Xlsx(document);
        this.rels = new IndexedLinkedList();
        this.styles = new IndexedLinkedList();
        this.sharedStrings = new IndexedLinkedList();
        this.sheets = new IndexedLinkedList();
        this.#maxRId = 0;
        this.#maxSheetId = 0;
        this.#maxSheetNameId = 0;

        this.#loadRelationships();
        this.#loadStyles();
        this.#loadSharedStrings();
        this.#loadSheets();
    }

    #loadRelationships() {
        const xmlRels = this.xlsx.getNodes("xl/_rels/workbook.xml.rels|Relationships|Relationship");
        _.forEach(xmlRels, xmlRel => {
            this.rels.add(xmlRel["@_Id"], xmlRel);
            if (xmlRel["@_Id"].startsWith("rId")) {
                this.#maxRId = Math.max(this.#maxRId, Number(xmlRel["@_Id"].substring(3)));
            }
        });
    }

    #loadStyles() {
        const xmlStyles = this.xlsx.getNodes("xl/styles.xml|styleSheet|cellXfs|xf");
        _.forEach(xmlStyles, (xmlStyle, i) => {
            const style = new CellStyle({
                workbook: this
            });
            style["numFmtId"] = xmlStyle["@_numFmtId"];
            style["fontId"] = xmlStyle["@_fontId"];
            style["fillId"] = xmlStyle["@_fillId"];
            style["borderId"] = xmlStyle["@_borderId"];
            style["styleId"] = i;
            this.styles.add(i, style);
        });
    }

    #loadSharedStrings() {
        const xmlSharedStrings = this.xlsx.getNodes("xl/sharedStrings.xml|sst|si");
        _.forEach(xmlSharedStrings, (xmlSharedString, i) => {
            this.sharedStrings[xmlSharedString["t"]] = i;
        });
    }

    #loadSheets() {
        const xmlSheets = this.xlsx.getNodes("xl/workbook.xml|workbook|sheets|sheet");
        _.forEach(xmlSheets, xmlSheet => {
            const xmlRel = this.rels.get(xmlSheet["@_r:id"]).value;
            const xmlFile = this.xlsx.getNode("xl/" + xmlRel["@_Target"]);
            const match = xmlRel["@_Target"].match("worksheets/sheet([0-9]+)\.xml");
            if (match) {
                this.#maxSheetNameId = Math.max(this.#maxSheetNameId, Number(match[1]));
            }
            const sheet = new Sheet({
                workbook: this,
                xmlSheet: xmlSheet,
                xmlRel: xmlRel,
                xmlFile: xmlFile
            });
            this.sheets.add(xmlSheet["@_name"], sheet);
            this.#maxSheetId = Math.max(this.#maxSheetId, Number(xmlSheet["@_sheetId"]));
        });
    }

    addOlePackage(oleData, label, filename, command) {
        // TODO: addOlePackage
    }

    addPicture(pictureData, format) {
        // TODO: addPicture
    }

    addPivotCache(rId) {
        // TODO: addPivotCache
    }

    addToolPack(toolpack) {
        // TODO: addToolPack
    }

    beforeDocumentRead() {
        // TODO: beforeDocumentRead
    }

    cloneSheet(index, newName) {
        // TODO: cloneSheet
    }

    close() {
        // TODO: close
    }

    commit() {
        // TODO: commit
    }

    createCellStyle(styles) {
        const style = new CellStyle({
            workbook: this
        });
        style.createNodes(styles);
        this.styles[style.styleId] = style;
        return style;
    }

    createDataFormat() {
        // TODO: createDataFormat
    }

    createDialogsheet(sheetname, dialogsheet) {
        // TODO: createDialogsheet
    }

    createEvaluationWorkbook() {
        // TODO: createEvaluationWorkbook
    }

    createName() {
        // TODO: createName
    }

    createSheet(sheetname) {
        if (sheetname == null) {
            throw "Invalid sheetname";
        }

        this.#maxRId++;
        this.#maxSheetId++;
        this.#maxSheetNameId++;

        const xmlRel = {
            "@_Id": "rId" + this.#maxRId,
            "@_Target": "worksheets/sheet" + this.#maxSheetNameId + ".xml",
            "@_Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        };
        this.xlsx.appendNode("xl/_rels/workbook.xml.rels|Relationships|Relationship", xmlRel);

        const xmlSheet = {
            "@_name": sheetname,
            "@_r:id": 'rId' + this.#maxRId,
            "@_sheetId": this.#maxSheetId.toString()
        };
        this.xlsx.appendNode("xl/workbook.xml|workbook|sheets|sheet", xmlSheet);

        this.xlsx.appendNode("[Content_Types].xml|Types|Override", {
            "@_ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            "@_PartName": "/xl/worksheets/sheet" + this.#maxSheetNameId + ".xml"
        });

        const xmlFile = this.parser.parse(`
            <?xml version="1.0" encoding="UTF-8"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <dimension ref="A1"/>
                <sheetViews>
                    <sheetView workbookViewId="0"/>
                </sheetViews>
                <sheetFormatPr defaultRowHeight="15.0"/>
                <sheetData/>
                <pageMargins top="0.75" right="0.7" left="0.7" header="0.3" footer="0.3" bottom="0.75"/>
            </worksheet>
        `);
        this.xlsx.setNode("xl/worksheets/sheet" + this.#maxSheetNameId + ".xml", xmlFile);

        const sheet = new Sheet({
            workbook: this,
            xmlSheet: xmlSheet,
            xmlRel: xmlRel,
            xmlFile: xmlFile
        });

        this.rels.add("rId" + this.#maxRId, xmlRel);
        this.sheets.add(sheetname, sheet);

        return sheet;
    }

    getActiveSheetIndex() {
        return Number(this.xlsx.getNode("xl/workbook.xml|workbook|bookViews|workbookView|@_activeTab"));
    }

    getAllEmbeddedParts() {
        // TODO: getAllEmbeddedParts
    }

    getAllNames() {
        // TODO: getAllNames
    }

    getAllPictures() {
        // TODO: getAllPictures
    }

    getCalculationChain() {
        // TODO: getCalculationChain
    }

    getCellFormulaValidation() {
        // TODO: getCellFormulaValidation
    }

    getCellReferenceType() {
        // TODO: getCellReferenceType
    }

    getCellStyleAt(index) {
        // TODO: getCellStyleAt
    }

    getCreationHelper(index) {
        // TODO: getCreationHelper
    }

    getCTWorkbook() {
        // TODO: getCTWorkbook
    }

    getCustomXMLMappings() {
        // TODO: getCustomXMLMappings
    }

    getExternalLinksTable() {
        // TODO: getExternalLinksTable
    }

    getFirstVisibleTab() {
        // TODO: getFirstVisibleTab
    }

    getForceFormulaRecalculation() {
        // TODO: getForceFormulaRecalculation
    }

    getMapInfo() {
        // TODO: getMapInfo
    }

    getMissingCellPolicy() {
        // TODO: getMissingCellPolicy
    }

    getName(name) {
        // TODO: getName
    }

    getNames(name) {
        // TODO: getNames
    }

    getNumberOfNames() {
        // TODO: getNumberOfNames
    }

    getNumberOfSheets() {
        // TODO: getNumberOfSheets
    }

    getNumCellStyles() {
        // TODO: getNumCellStyles
    }

    getPivotTables() {
        // TODO: getPivotTables
    }

    getPrintArea(sheetIndex) {
        // TODO: getPivotTables
    }

    getSharedStringSource() {
        // TODO: getSharedStringSource
    }

    getSheet(name) {
        return this.sheets.get(name).value;
    }

    getSheetAt(index) {
        return this.getSheet(this.getSheetName(index));
    }

    getSheetIndex(nameOrSheet) {
        let index = -1;
        if (typeof nameOrSheet == "string") {
            this.sheets.each((key, value) => {
                index++;
                if (nameOrSheet == key) {
                    return index;
                }
            });
        } else {
            this.sheets.each((key, value) => {
                index++;
                if (nameOrSheet == value) {
                    return index;
                }
            });
        }
        return index;
    }

    getSheetName(index) {
        return this.xlsx.getNodes("xl/workbook.xml|workbook|sheets|sheet")[index]["@_name"];
    }

    getSheetVisibility(index) {
        // TODO: getSheetVisibility
    }

    getSpreadsheetVersion() {
        // TODO: getSpreadsheetVersion
    }

    getStylesSource() {
        // TODO: getStylesSource
    }

    getTable(name) {
        // TODO: getTable
    }

    getWorkbookType() {
        // TODO: getWorkbookType
    }

    getXssfFactory() {
        // TODO: getXssfFactory
    }

    isDate1904() {
        const date1904 = this.xlsx.getNode("xl/workbook.xml|workbook|workbookPr|@_date1904");
        if (date1904 == "true" || date1904 == "1") {
            return true;
        } else {
            return false;
        }
    }

    isHidden() {
        // TODO: isHidden
    }

    isMacroEnabled() {
        // TODO: isMacroEnabled
    }

    isRevisionLocked() {
        // TODO: isRevisionLocked
    }

    isSheetHidden(index) {
        // TODO: isSheetHidden
    }

    isSheetVeryHidden(index) {
        // TODO: isSheetVeryHidden
    }

    isStructureLocked() {
        // TODO: isStructureLocked
    }

    isWindowsLocked() {
        // TODO: isWindowsLocked
    }

    linkExternalWorkbook(name, workbook) {
        // TODO: linkExternalWorkbook
    }

    lockRevision() {
        // TODO: lockRevision
    }

    lockStructure() {
        // TODO: lockStructure
    }

    lockWindows() {
        // TODO: lockWindows
    }

    newPackage(workbookType) {
        // TODO: newPackage
    }

    onDeleteFormula(cell) {
        // TODO: onDeleteFormula
    }

    onDocumentRead() {
        // TODO: onDocumentRead
    }

    parseSheet(shIdMap, ctSheet) {
        // TODO: parseSheet
    }

    removeName(name) {
        // TODO: removeName
    }

    removePrintArea(sheetIndex) {
        // TODO: removePrintArea
    }

    removeSheetAt(index) {
        const sheet = this.getSheetAt(index);
        const rels = this.xlsx.getNodes("xl/_rels/workbook.xml.rels|Relationships|Relationship");
        const removing = _.find(rels, { "@_Id": sheet.xmlRel["@_Id"] });
        this.xlsx.removeNode("xl/" + removing["@_Target"]);
        this.xlsx.removeNode("xl/workbook.xml|workbook|sheets|sheet", { "@_r:id": removing["@_Id"] })
        this.xlsx.removeNode("[Content_Types].xml|Types|Override", { "@_PartName": "/xl/" + removing["@_Target"] })
        this.xlsx.removeNode("xl/_rels/workbook.xml.rels|Relationships|Relationship", { "@_Id": removing["@_Id"] });
    }

    setActiveSheet(index) {
        this.xlsx.setNode("xl/workbook.xml|workbook|bookViews|workbookView|@_activeTab", index.toString());
    }

    setCellFormulaValidation(value) {
        // TODO: setCellFormulaValidation
    }

    setCellReferenceType(cellReferenceType) {
        // TODO: setCellReferenceType
    }

    setFirstVisibleTab(index) {
        // TODO: setFirstVisibleTab
    }

    setForceFormulaRecalculation(value) {
        // TODO: setForceFormulaRecalculation
    }

    setHidden(hiddenFlag) {
        // TODO: setHidden
    }

    setMissingCellPolicy(missingCellPolicy) {
        // TODO: setMissingCellPolicy
    }

    setPivotTables(pivotTables) {
        // TODO: setPivotTables
    }

    setPrintArea(sheetIndex, reference) {
        // TODO: setPrintArea
    }

    setRevisionsPassword(password, hashAlgo) {
        // TODO: setRevisionsPassword
    }

    setSelectedTab(index) {
        // TODO: setSelectedTab
    }

    setSheetHidden(sheetIx, hidden) {
        // TODO: setSheetHidden
    }

    setSheetName(index, name) {
        // TODO: setSheetName
    }

    setSheetOrder(name, index) {
        // TODO: setSheetOrder
    }

    setSheetVisibility(sheetIx, visibility) {
        // TODO: setSheetVisibility
    }

    setVBAProject(vbaProjectStreamOrMacroWorkbook) {
        // TODO: setVBAProject
    }

    setWorkbookPassword(password, hashAlgo) {
        // TODO: setWorkbookPassword
    }

    setWorkbookType(type) {
        // TODO: setWorkbookType
    }

    sheetIterator() {
        // TODO: sheetIterator
    }

    spliterator() {
        // TODO: spliterator
    }

    unLock() {
        // TODO: unLock
    }

    unLockRevision() {
        // TODO: unLockRevision
    }

    unLockStructure() {
        // TODO: unLockStructure
    }

    unLockWindows() {
        // TODO: unLockWindows
    }

    validateRevisionsPassword(password) {
        // TODO: validateRevisionsPassword
    }

    validateWorkbookPassword(password) {
        // TODO: validateWorkbookPassword
    }

    write(filename) {
        const contents = this.#build();
        const zip = new JSZip();
        for (let key in contents) {
            zip.file(key, contents[key]);
        }
        for (let key in zip.files) {
            if (zip.files[key].dir) {
                delete zip.files[key];
            }
        }
        zip.generateAsync({ type: "blob", compression: "DEFLATE" }).then((content) => {
            saveAs(content, filename);
        });
    }

    #build() {
        const contents = {};
        for (let key in this.xlsx.document) {
            const builder = [];
            if (this.xlsx.document[key]["?xml"]) {
                builder.push("<?xml");
                for (let k in this.xlsx.document[key]["?xml"]) {
                    builder.push(" " + k.substring(2) + "=" + '"' + this.xlsx.document[key]["?xml"][k] + '"');
                }
                builder.push("?>");
            }
            for (let k in this.xlsx.document[key]) {
                if (k != "?xml") {
                    this.#buildXml(builder, k, this.xlsx.document[key][k]);
                }
            }
            contents[key] = builder.join("");
        }
        return contents;
    }

    #buildXml(builder, key, value) {
        if (typeof value == "object") {
            if (Array.isArray(value)) {
                _.forEach(value, (item) => {
                    this.#buildXml(builder, key, item);
                });
            } else {
                builder.push("<" + key);
                let attribute = [];
                let element = [];
                let text;
                for (let k in value) {
                    if (k.substring(0, 2) == "@_") {
                        attribute.push({
                            key: k,
                            value: value[k]
                        });
                    } else if (k == "#text") {
                        text = value[k];
                    } else {
                        element.push({
                            key: k,
                            value: value[k]
                        });
                    }
                }
                _.forEach(attribute, (item) => {
                    builder.push(" " + item.key.substring(2) + "=" + '"' + item.value + '"');
                });
                if (element.length > 0 || text != null) {
                    builder.push(">");
                    if (text) {
                        builder.push(text);
                    }
                    if (element.length > 0) {
                        _.forEach(element, (item) => {
                            this.#buildXml(builder, item.key, item.value);
                        });
                    }
                    builder.push("</" + key + ">");
                } else {
                    builder.push("/>");
                }
            }
        } else {
            builder.push("<" + key);
            if (value != null && value !== "") {
                builder.push(">");
                builder.push(value);
                builder.push("</" + key + ">");
            } else {
                builder.push("/>");
            }
        }
    }
}