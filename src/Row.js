import _ from "lodash";
import IndexedLinkedList from "./IndexedLinkedList";
import Xlsx from "./Xlsx";
import Cell from "./Cell";
import Utils from "./Utils";

/**
 * @module Row
 */
export default class Row {

    constructor(option) {
        this.workbook = option.workbook;
        this.sheet = option.sheet;
        this.xmlRow = option.xmlRow;

        this.xlsx = new Xlsx(this.xmlRow);
        this.cells = new IndexedLinkedList();

        this.#loadCells();
    }

    #loadCells() {
        const xmlcells = this.xlsx.getNodes("c");
        _.forEach(xmlcells, xmlCell => {
            const cell = new Cell({
                workbook: this.workbook,
                sheet: this.sheet,
                row: this,
                xmlCell: xmlCell
            });
            const match = xmlCell["@_r"].match(/[A-Z]+([0-9]+)/);
            this.cells.add(match[1], cell);
        });
    }

    compareTo(other) {
        // TODO: compareTo
    }

    copyRowFrom(srcRow, policy) {
        // TODO: copyRowFrom
    }

    /**
     * @summary 셀을 생성한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createrow(0);
     * var cell = row.createCell(0);
     * @param {Number}
     * @returns {Cell}
     */
    createCell(cellIndex) {
        const xmlCell = {
            "@_r": Utils.indexToAlphabet(cellIndex + 1) + this.xlsx.getNode("@_r"),
            "@_s": "0"
        };
        this.xlsx.setNode("c|" + cellIndex, xmlCell, true);
        const cell = new Cell({
            workbook: this.workbook,
            sheet: this.sheet,
            row: this,
            xmlCell: xmlCell
        });
        this.cells.add(cellIndex, cell);
        return cell;
    }

    /**
     * @summary 셀을 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(0);
     * var cell = row.createCell(0);
     * console.log(cell === row.getCell(0)); // true
     * @param {Number}
     * @returns {Cell}
     */
    getCell(cellIndex) {
        const cell = this.cells.get(cellIndex);
        if (cell != null) {
            return cell.value;
        }
        return undefined;        
    }

    /**
     * @summary 행에 존재하는 첫번째 셀의 Index를 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(0);
     * row.createCell(5);
     * row.createCell(3);
     * row.createCell(4);
     * console.log(row.getFirstCellNum()); // 3
     * @returns {Number}
     */
    getFirstCellNum() {
        let min;
        this.cells.each(key => {
            if (min == null) {
                min = key;
            }
            min = Math.min(min, key);
        });
        return min;
    }

    /**
     * @summary 행의 높이를 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(0);
     * row.setHeight(100);
     * console.log(row.getHeight()); // 100
     * @returns {Number}
     */
    getHeight() {
        if (this.xlsx.getNode("@_ht") != null) {
            return Number(this.xlsx.getNode("@_ht"));
        } else {
            return Number(this.sheet.xlsx.getNode("worksheet|sheetFormatPr|@_defaultRowHeight"));
        }
    }

    /**
     * @summary 행에 존재하는 마지막 셀의 Index를 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(0);
     * row.createCell(5);
     * row.createCell(3);
     * row.createCell(4);
     * console.log(row.getLastCellNum()); // 5
     * @returns {Number}
     */
    getLastCellNum() {
        let max;
        this.cells.each(key => {
            if (max == null) {
                max = key;
            }
            max = Math.max(max, key);
        });
        return max;
    }

    getOutlineLevel() {
        // TODO: getOutlineLevel
    }

    getPhysicalNumberOfCells() {
        // TODO: getPhysicalNumberOfCells
    }

    /**
     * @summary 행의 Index를 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(3);
     * console.log(row.getRowNum()); // 3
     * @returns {Number}
     */
    getRowNum() {
        return Number(this.xlsx.getNode("@_r")) - 1;
    }

    getRowStyle() {
        // TODO: getRowStyle
    }

    /**
     * @summary 행이 속한 시트 객체를 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(0);
     * console.log(sheet === row.getSheet()); // true
     * @returns {Number}
     */
    getSheet() {
        return this.sheet;
    }

    getZeroHeight() {
        // TODO: getZeroHeight
    }

    isFormatted() {
        // TODO: isFormatted
    }

    onDocumentWrite() {
        // TODO: onDocumentWrite
    }

    /**
     * @summary 셀을 삭제한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(0);
     * var cell1 = row.createCell(1);
     * var cell2 = row.createCell(2);
     * var cell3 = row.createCell(3);
     * row.removeCell(2);
     * console.log(row.getCell(2)); // undefined
     * @param {Number}
     * @returns {Void}
     */
    removeCell(cellIndex) {
        this.xlsx.removeNode("c|" + cellIndex);
        this.cells.remove(cellIndex);
    }

    /**
     * @summary 행의 높이를 변경한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(0);
     * row.setHeight(50);
     * console.log(row.getHeight()); // 50
     * @param {Number}
     * @returns {Void}
     */
    setHeight(height) {
        this.xlsx.setNode("@_ht", height.toString());
        this.xlsx.setNode("@_customHeight", "true");
    }

    /**
     * @summary 행의 위치를 변경한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(0);
     * row.setRowNum(5);
     * console.log(row.getRowNum()); // 5
     * @param {Number}
     * @returns {Void}
     */
    setRowNum(rowIndex) {
        const origin = this.xlsx.getNode("@_r");
        const index = (rowIndex + 1).toString();
        const diff = index - origin;
        this.xlsx.setNode("@_r", index);
        if (diff > 0) {
            this.cells.eachRight((key, value) => {
                this.cells[key + diff] = value;
                delete this.cells[key];
                value.xlsx.setNode("@_r", value.xlsx.getNode("@_r").replace(/[0-9]+/, "") + index);
            });
        } else if (index - origin < 0) {
            this.cells.each((key, value) => {
                this.cells[key - diff] = value;
                delete this.cells[key];
                value.xlsx.setNode("@_r", value.xlsx.getNode("@_r").replace(/[0-9]+/, "") + index);
            });
        }
    }

    setRowStyle(style) {
        // TODO: setRowStyle
    }

    setZeroHeight(height) {
        // TODO: setZeroHeight
    }

    shift(n) {
        // TODO: shift
    }

    shiftCellsLeft(firstShiftColumnIndex, lastShiftColumnIndex, step) {
        // TODO: shiftCellsLeft
    }

    shiftCellsRight(firstShiftColumnIndex, lastShiftColumnIndex, step) {
        // TODO: shiftCellsRight
    }

    spliterator() {
        // TODO: spliterator
    }
}