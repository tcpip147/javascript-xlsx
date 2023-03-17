import _ from "lodash";
import IndexedLinkedList from "./IndexedLinkedList";
import Xlsx from "./Xlsx";
import Row from "./Row";
import Utils from "./Utils";

/**
 * @module Sheet
 */
export default class Sheet {

    constructor(option) {
        this.workbook = option.workbook;
        this.xmlSheet = option.xmlSheet;
        this.xmlRel = option.xmlRel;
        this.xmlFile = option.xmlFile;

        this.xlsx = new Xlsx(this.xmlFile);
        this.rows = new IndexedLinkedList();

        this.#loadRows();
    }

    #loadRows() {
        const xmlRows = this.xlsx.getNodes("worksheet|sheetData|row");
        _.forEach(xmlRows, xmlRow => {
            const row = new Row({
                workbook: this.workbook,
                sheet: this,
                xmlRow: xmlRow
            });
            this.rows.add(xmlRow["@_r"], row);
        });
    }

    addHyperlink(hyperlink) {
        // TODO: addHyperlink
    }

    addIgnoredErrors(region, ignoredErrorTypes) {
        // TODO: addIgnoredErrors
    }

    addIgnoredErrors(cell, ignoredErrorTypes) {
        // TODO: addIgnoredErrors
    }

    /**
     * @summary 셀 병합을 실행한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * sheet.addMergedRegion("A1:B3");
     * @param {String}
     * @returns {Void}
     */
    addMergedRegion(region) {
        this.xlsx.afterNodeKey("worksheet|sheetData", "mergeCells");
        this.xlsx.appendNode("worksheet|mergeCells|mergeCell", {
            "@_ref": region
        });
    }

    addValidationData(dataValidation) {
        // TODO: addValidationData
    }

    autoSizeColumn(column) {
        // TODO: autoSizeColumn
    }

    autoSizeColumn(column, useMergedCells) {
        // TODO: autoSizeColumn
    }

    commit() {
        // TODO: commit
    }

    copyRows(srcStartRow, srcEndRow, destStartRow, cellCopyPolicy) {
        // TODO: copyRows
    }

    copyRows(srcRows, destStartRow, policy) {
        // TODO: copyRows
    }

    createDrawingPatriarch() {
        // TODO: createDrawingPatriarch
    }

    createFreezePane(colSplit, rowSplit) {
        // TODO: createFreezePane
    }

    createFreezePane(colSplit, rowSplit, leftmostColumn, topRow) {
        // TODO: createFreezePane
    }

    createPivotTable(source, position) {
        // TODO: createPivotTable
    }

    createPivotTable(source, position, sourceSheet) {
        // TODO: createPivotTable
    }

    /**
     * @summary 행을 추가한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(0);
     * @param {Number}
     * @returns {Row}
     */
    createRow(rowIndex) {
        const xmlRow = {
            "@_r": (Number(rowIndex) + 1).toString()
        };
        this.xlsx.setNode("worksheet|sheetData|row|" + rowIndex, xmlRow, true);
        const row = new Row({
            workbook: this.workbook,
            sheet: this,
            xmlRow: xmlRow
        });
        this.rows.add(rowIndex, row);
        return row;
    }

    createSplitPane(xSplitPos, ySplitPos, leftmostColumn, topRow, activePane) {
        // TODO: createSplitPane
    }

    createTable(tableArea) {
        // TODO: createTable
    }

    disableLocking() {
        // TODO: disableLocking
    }

    enableLocking() {
        // TODO: enableLocking
    }

    findEndOfRowOutlineGroup(row) {
        // TODO: findEndOfRowOutlineGroup
    }

    /**
     * @summary 선택된 셀을 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * sheet.setActiveCell("A1");
     * console.log(sheet.getActiveCell()); // A1
     * @returns {String}
     */
    getActiveCell() {
        return this.xlsx.getNode("worksheet|sheetViews|sheetView|selection|@_activeCell");
    }

    getAutobreaks() {
        // TODO: getAutobreaks
    }

    getCellComment(address) {
        // TODO: getCellComment
    }

    getCellComments() {
        // TODO: getCellComments
    }

    getColumnBreaks() {
        // TODO: getColumnBreaks
    }

    getColumnHelper() {
        // TODO: getColumnHelper
    }

    getColumnOutlineLevel(columnIndex) {
        // TODO: getColumnOutlineLevel
    }

    getColumnStyle(column) {
        // TODO: getColumnStyle
    }

    /**
     * @summary 열의 너비를 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * sheet.setColumnWidth(1, 100);
     * console.log(sheet.getColumnWidth(1)); // 100
     * @param {Number}
     * @returns {Void}
     */
    getColumnWidth(columnIndex) {
        const columns = this.xlsx.getNodes("worksheet|cols|col");
        let width;
        _.forEach(columns, column => {
            if (column["@_min"] == columnIndex + 1) {
                width = column["@_width"];
                return false;
            }
        });
        if (width == null) {
            width = this.getDefaultColumnWidth();
        }
        return width;
    }

    getColumnWidthInPixels(columnIndex) {
        // TODO: getColumnWidthInPixels
    }

    getCommentsTable(create) {
        // TODO: getCommentsTable
    }

    getCTDrawing() {
        // TODO: getCTDrawing
    }

    getCTLegacyDrawing() {
        // TODO: getCTLegacyDrawing
    }

    getCTWorksheet() {
        // TODO: getCTWorksheet
    }

    getDataValidationHelper() {
        // TODO: getDataValidationHelper
    }

    getDataValidations() {
        // TODO: getDataValidations
    }

    /**
     * @summary 열의 너비 기본값을 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * sheet.setDefaultColumnWidth(100);
     * console.log(sheet.getDefaultColumnWidth()); // 100
     * @returns {Number}
     */
    getDefaultColumnWidth() {
        let width = this.xlsx.getNode("worksheet|sheetFormatPr|@_baseColWidth");
        if (width == null) {
            width = 8;
        }
        return width;
    }

    /**
     * @summary 행의 높이 기본값을 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * sheet.setDefaultRowHeight(50);
     * console.log(sheet.getDefaultRowHeight()); // 50
     * @returns {Number}
     */
    getDefaultRowHeight() {
        return Number(this.xlsx.getNode("worksheet|sheetFormatPr|@_defaultRowHeight"));
    }

    getDimension() {
        // TODO: getDimension
    }

    getDisplayGuts() {
        // TODO: getDisplayGuts
    }

    getDrawingPatriarch() {
        // TODO: getDrawingPatriarch
    }

    getEvenFooter() {
        // TODO: getEvenFooter
    }

    getEvenHeader() {
        // TODO: getEvenHeader
    }

    getFirstFooter() {
        // TODO: getFirstFooter
    }

    getFirstHeader() {
        // TODO: getFirstHeader
    }

    /**
     * @summary 시트에 존재하는 첫번째 행의 Index를 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * sheet.createRow(4);
     * sheet.createRow(3);
     * sheet.createRow(5);
     * console.log(sheet.getFirstRowNum()); // 3
     * @returns {Number}
     */
    getFirstRowNum() {
        let min;
        this.rows.each((key, value) => {
            if (min == null) {
                min = key;
            }
            min = Math.min(min, key);
        });
        return min;
    }

    getFitToPage() {
        // TODO: getFitToPage
    }

    getFooter() {
        // TODO: getFooter
    }

    getForceFormulaRecalculation() {
        // TODO: getForceFormulaRecalculation
    }

    getHeader() {
        // TODO: getHeader
    }

    getHeaderFooterProperties() {
        // TODO: getHeaderFooterProperties
    }

    getHorizontallyCenter() {
        // TODO: getHorizontallyCenter
    }

    getHyperlink(addr) {
        // TODO: getHyperlink
    }

    getHyperlink(row, column) {
        // TODO: getHyperlink
    }

    getHyperlinkList() {
        // TODO: getHyperlinkList
    }

    getIgnoredErrors() {
        // TODO: getIgnoredErrors
    }

    /**
     * @summary 시트에 존재하는 마지막 행의 Index를 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * sheet.createRow(4);
     * sheet.createRow(5);
     * sheet.createRow(3);
     * console.log(sheet.getLastRowNum()); // 5
     * @returns {Number}
     */
    getLastRowNum() {
        let max;
        this.rows.each((key, value) => {
            if (max == null) {
                max = key;
            }
            max = Math.max(max, key);
        });
        return max;
    }

    getLeftCol() {
        // TODO: getLeftCol
    }

    getMargin(margin) {
        // TODO: getMargin
    }

    getMergedRegion(index) {
        // TODO: getMergedRegion
    }

    getMergedRegions() {
        // TODO: getMergedRegions
    }

    getNumberOfComments() {
        // TODO: getNumberOfComments
    }

    getNumHyperlinks() {
        // TODO: getNumHyperlinks
    }

    getNumMergedRegions() {
        // TODO: getNumMergedRegions
    }

    getOddFooter() {
        // TODO: getOddFooter
    }

    getOddHeader() {
        // TODO: getOddHeader
    }

    getPaneInformation() {
        // TODO: getPaneInformation
    }

    getPhysicalNumberOfRows() {
        // TODO: getPhysicalNumberOfRows
    }

    getPivotTables() {
        // TODO: getPivotTables
    }

    getPrintSetup() {
        // TODO: getPrintSetup
    }

    getProtect() {
        // TODO: getProtect
    }

    getRepeatingColumns() {
        // TODO: getRepeatingColumns
    }

    getRepeatingRows() {
        // TODO: getRepeatingRows
    }

    /**
     * @summary 시트를 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(0);
     * console.log(row === sheet.getRow(0)); // true
     * @param {Number}
     * @returns {Row}
     */
    getRow(rowIndex) {
        const row = this.rows.get(rowIndex);
        if (row != null) {
            return row.value;
        }
        return undefined;
    }

    getRowBreaks() {
        // TODO: getRowBreaks
    }

    getRowSumsBelow() {
        // TODO: getRowSumsBelow
    }

    getRowSumsRight() {
        // TODO: getRowSumsRight
    }

    getScenarioProtect() {
        // TODO: getScenarioProtect
    }

    getSharedFormula(sid) {
        // TODO: getSharedFormula
    }

    getSheetConditionalFormatting() {
        // TODO: getSheetConditionalFormatting
    }

    /**
     * @summary 시트명을 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var sheetname = sheet.getSheetName(0);
     * console.log(sheetname); // Sheet1
     * @returns {String}
     */
    getSheetName() {
        return this.xmlSheet["@_name"];
    }

    getSheetTypeSheetViews(create) {
        // TODO: getSheetTypeSheetViews
    }

    getTabColor() {
        // TODO: getTabColor
    }

    getTables() {
        // TODO: getTables
    }

    getTopRow() {
        // TODO: getTopRow
    }

    getVerticallyCenter() {
        // TODO: getVerticallyCenter
    }

    getVMLDrawing(autoCreate) {
        // TODO: getVMLDrawing
    }

    getWorkbook() {
        return this.workbook;
    }

    groupColumn(fromColumn, toColumn) {
        // TODO: groupColumn
    }

    groupRow(fromRow, toRow) {
        // TODO: groupRow
    }

    hasComments() {
        // TODO: hasComments
    }

    isAutoFilterLocked() {
        // TODO: isAutoFilterLocked
    }

    isColumnBroken(column) {
        // TODO: isColumnBroken
    }

    isColumnHidden(columnIndex) {
        // TODO: isColumnHidden
    }

    isDeleteColumnsLocked() {
        // TODO: isDeleteColumnsLocked
    }

    isDeleteRowsLocked() {
        // TODO: isDeleteRowsLocked
    }

    isDisplayFormulas() {
        // TODO: isDisplayFormulas
    }

    isDisplayGridlines() {
        // TODO: isDisplayGridlines
    }

    isDisplayRowColHeadings() {
        // TODO: isDisplayRowColHeadings
    }

    isDisplayZeros() {
        // TODO: isDisplayZeros
    }

    isFormatCellsLocked() {
        // TODO: isFormatCellsLocked
    }

    isFormatColumnsLocked() {
        // TODO: isFormatColumnsLocked
    }

    isFormatRowsLocked() {
        // TODO: isFormatRowsLocked
    }

    isInsertColumnsLocked() {
        // TODO: isInsertColumnsLocked
    }

    isInsertHyperlinksLocked() {
        // TODO: isInsertHyperlinksLocked
    }

    isInsertRowsLocked() {
        // TODO: isInsertRowsLocked
    }

    isObjectsLocked() {
        // TODO: isObjectsLocked
    }

    isPivotTablesLocked() {
        // TODO: isPivotTablesLocked
    }

    isPrintGridlines() {
        // TODO: isPrintGridlines
    }

    isPrintRowAndColumnHeadings() {
        // TODO: isPrintRowAndColumnHeadings
    }

    isRightToLeft() {
        // TODO: isRightToLeft
    }

    isRowBroken(row) {
        // TODO: isRowBroken
    }

    isScenariosLocked() {
        // TODO: isScenariosLocked
    }

    isSelected() {
        // TODO: isSelected
    }

    isSelectLockedCellsLocked() {
        // TODO: isSelectLockedCellsLocked
    }

    isSelectUnlockedCellsLocked() {
        // TODO: isSelectUnlockedCellsLocked
    }

    isSheetLocked() {
        // TODO: isSheetLocked
    }

    isSortLocked() {
        // TODO: isSortLocked
    }

    lockAutoFilter(enabled) {
        // TODO: lockAutoFilter
    }

    lockDeleteColumns(enabled) {
        // TODO: lockDeleteColumns
    }

    lockDeleteRows(enabled) {
        // TODO: lockDeleteRows
    }

    lockFormatCells(enabled) {
        // TODO: lockFormatCells
    }

    lockFormatColumns(enabled) {
        // TODO: lockFormatColumns
    }

    lockFormatRows(enabled) {
        // TODO: lockFormatRows
    }

    lockInsertColumns(enabled) {
        // TODO: lockInsertColumns
    }

    lockInsertHyperlinks(enabled) {
        // TODO: lockInsertHyperlinks
    }

    lockInsertRows(enabled) {
        // TODO: lockInsertRows
    }

    lockObjects(enabled) {
        // TODO: lockObjects
    }

    lockPivotTables(enabled) {
        // TODO: lockPivotTables
    }

    lockScenarios(enabled) {
        // TODO: lockScenarios
    }

    lockSelectLockedCells(enabled) {
        // TODO: lockSelectLockedCells
    }

    lockSelectUnlockedCells(enabled) {
        // TODO: lockSelectUnlockedCells
    }

    lockSort(enabled) {
        // TODO: lockSort
    }

    onDeleteFormula(cell, evalWb) {
        // TODO: onDeleteFormula
    }

    onDocumentCreate() {
        // TODO: onDocumentCreate
    }

    onDocumentRead() {
        // TODO: onDocumentRead
    }

    onSheetDelete() {
        // TODO: onSheetDelete
    }

    protectSheet(password) {
        // TODO: protectSheet
    }

    read(is) {
        // TODO: read
    }

    readOleObject(shapeId) {
        // TODO: readOleObject
    }

    removeArrayFormula(cell) {
        // TODO: removeArrayFormula
    }

    removeColumnBreak(column) {
        // TODO: removeColumnBreak
    }

    removeHyperlink(row, column) {
        // TODO: removeHyperlink
    }

    removeHyperlink(hyperlink) {
        // TODO: removeHyperlink
    }

    removeMergedRegion(index) {
        // TODO: removeMergedRegion
    }

    removeMergedRegions(indices) {
        // TODO: removeMergedRegions
    }

    /**
     * @summary 행을 삭제한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(0);
     * sheet.removeRow(0);
     * console.log(sheet.getRow(0)); // undefined
     * @param {Number}
     * @returns {Void}
     */
    removeRow(rowIndex) {
        this.xlsx.removeNode("worksheet|sheetData|row|" + rowIndex);
        this.rows.remove(rowIndex);
    }

    removeRowBreak(row) {
        // TODO: removeRowBreak
    }

    removeTable(t) {
        // TODO: removeTable
    }

    rowIterator() {
        // TODO: rowIterator
    }

    /**
     * @summary 셀을 선택한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * sheet.setActiveCell("A1");
     * console.log(sheet.getActiveCell()); // A1
     * @returns {Cell}
     */
    setActiveCell(address) {
        this.xlsx.setNode("worksheet|sheetViews|sheetView|selection", {
            "@_activeCell": address,
            "@_sqref": address
        });
    }

    setArrayFormula(formula, range) {
        // TODO: setArrayFormula
    }

    setAutobreaks(value) {
        // TODO: setAutobreaks
    }

    setAutoFilter(range) {
        // TODO: setAutoFilter
    }

    setColumnBreak(column) {
        // TODO: setColumnBreak
    }

    setColumnGroupCollapsed(columnNumber, collapsed) {
        // TODO: setColumnGroupCollapsed
    }

    setColumnHidden(columnIndex, hidden) {
        // TODO: setColumnHidden
    }

    /**
     * @summary 열의 너비를 변경한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * sheet.setColumnWidth(1, 100);
     * console.log(sheet.getColumnWidth(1)); // 100
     * @param {Number}
     * @param {Number}
     * @returns {Void}
     */
    setColumnWidth(columnIndex, width) {
        this.xlsx.afterNodeKey("worksheet|sheetFormatPr", "cols");
        this.xlsx.appendNode("worksheet|cols|col", {
            "@_min": (columnIndex + 1).toString(),
            "@_max": (columnIndex + 1).toString(),
            "@_width": width,
            "@_customWidth": "true"
        });
    }

    setDefaultColumnStyle(column, style) {
        // TODO: setDefaultColumnStyle
    }

    /**
     * @summary 열의 너비 기본값을 변경한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * sheet.setDefaultColumnWidth(100);
     * console.log(sheet.getDefaultColumnWidth()); // 100
     * @param {Number}
     * @returns {Void}
     */
    setDefaultColumnWidth(width) {
        this.xlsx.setNode("worksheet|sheetFormatPr|@_baseColWidth", width);
    }

    /**
     * @summary 행의 높이 기본값을 변경한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * sheet.setDefaultRowHeight(50);
     * console.log(sheet.getDefaultRowHeight()); // 50
     * @param {Number}
     * @returns {Void}
     */
    setDefaultRowHeight(height) {
        return this.xlsx.setNode("worksheet|sheetFormatPr|@_defaultRowHeight", height);
    }

    setDimensionOverride(dimension) {
        // TODO: setDimensionOverride
    }

    setDisplayFormulas(show) {
        // TODO: setDisplayFormulas
    }

    setDisplayGridlines(show) {
        // TODO: setDisplayGridlines
    }

    setDisplayGuts(value) {
        // TODO: setDisplayGuts
    }

    setDisplayRowColHeadings(show) {
        // TODO: setDisplayRowColHeadings
    }

    setDisplayZeros(value) {
        // TODO: setDisplayZeros
    }

    setFitToPage(b) {
        // TODO: setFitToPage
    }

    setForceFormulaRecalculation(value) {
        // TODO: setForceFormulaRecalculation
    }

    setHorizontallyCenter(value) {
        // TODO: setHorizontallyCenter
    }

    setMargin(margin, size) {
        // TODO: setMargin
    }

    setPrintGridlines(value) {
        // TODO: setPrintGridlines
    }

    setPrintRowAndColumnHeadings(value) {
        // TODO: setPrintRowAndColumnHeadings
    }

    setRepeatingColumns(columnRangeRef) {
        // TODO: setRepeatingColumns
    }

    setRepeatingRows(rowRangeRef) {
        // TODO: setRepeatingRows
    }

    setRightToLeft(value) {
        // TODO: setRightToLeft
    }

    setRowBreak(row) {
        // TODO: setRowBreak
    }

    setRowGroupCollapsed(rowIndex, collapse) {
        // TODO: setRowGroupCollapsed
    }

    setRowSumsBelow(value) {
        // TODO: setRowSumsBelow
    }

    setRowSumsRight(value) {
        // TODO: setRowSumsRight
    }

    setSelected(value) {
        // TODO: setSelected
    }

    setSheetPassword(password, hashAlgo) {
        // TODO: setSheetPassword
    }

    setTabColor(color) {
        // TODO: setTabColor
    }

    setVerticallyCenter(value) {
        // TODO: setVerticallyCenter
    }

    /**
     * @summary 시트의 확대 비율을 변경한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * sheet.setZoom(50);
     * @param {Number}
     * @returns {Void}
     */
    setZoom(scale) {
        this.xlsx.setNode("worksheet|sheetViews|sheetView|@_zoomScale", scale);
    }

    shiftColumns(startColumn, endColumn, n) {
        // TODO: shiftColumns
    }

    shiftRows(startRow, endRow, n) {
        // TODO: shiftRows
    }

    shiftRows(startRow, endRow, n, copyRowHeight, resetOriginalRowHeight) {
        // TODO: shiftRows
    }

    showInPane(topRow, leftCol) {
        // TODO: showInPane
    }

    spliterator() {
        // TODO: spliterator
    }

    ungroupColumn(fromColumn, toColumn) {
        // TODO:ungroupColumn 
    }

    ungroupRow(fromRow, toRow) {
        // TODO: ungroupRow
    }

    validateMergedRegions() {
        // TODO: validateMergedRegions
    }

    validateSheetPassword(password) {
        // TODO: validateSheetPassword
    }

    write(out) {
        // TODO: write
    }
}