import _ from "lodash";
import IndexedLinkedList from "./IndexedLinkedList";
import Xlsx from "./Xlsx";
import Row from "./Row";
import Utils from "./Utils";

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

    getActiveCell() {
        const ref = this.xlsx.getNode("worksheet|sheetViews|sheetView|selection|@_activeCell");
        if (ref == null) {
            const row = this.rows.get(0);
            if (row != null) {
                return row.value.cells.get(0);
            }
        } else {
            const match = ref.match(/([A-Z]+)([0-9]+)/);
            return this.rows.get(match[2] - 1).value.cells.get(Utils.alphabetToIndex(match[1]) - 1).value;
        }
        return undefined;
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

    getColumnWidth(columnIndex) {
        const columns = this.xlsx.getNodes("worksheet|cols|col");
        let width;
        _.forEach(columns, column => {
            if (column["@_min"] == columnIndex) {
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

    getDefaultColumnWidth() {
        let width = this.xlsx.getNode("worksheet|sheetFormatPr|@_baseColWidth");
        if (width == null) {
            width = 8;
        }
        return width;
    }

    getDefaultRowHeight() {
        return this.xlsx.getNode("worksheet|sheetFormatPr|@_defaultRowHeight");
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

    getFirstRowNum() {
        return Number(this.rows.first().value.xmlRow["@_r"]) - 1;
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

    getLastRowNum() {
        return Number(this.rows.last().value.xmlRow["@_r"]) - 1;
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

    getRow(rowIndex) {
        return this.rows.get(rowIndex);
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
        if (this.rows.first() != null) {
            return this.rows.first().value;
        }
        return undefined;
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

    setDefaultColumnWidth(width) {
        this.xlsx.setNode("worksheet|sheetFormatPr|@_baseColWidth", width);
    }

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