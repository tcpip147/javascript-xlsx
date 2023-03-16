import _ from "lodash";
import IndexedLinkedList from "./IndexedLinkedList";
import Xlsx from "./Xlsx";
import Cell from "./Cell";
import Utils from "./Utils";

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

    getCell(cellIndex) {
        return this.cells.get(cellIndex).value;
    }

    getFirstCellNum() {
        return this.cells.first().value.getColumnIndex();
    }

    getHeight() {
        if (this.xlsx.getNode("@_ht") != null) {
            return this.xlsx.getNode("@_ht");
        } else {
            return this.sheet.xlsx.getNode("worksheet|sheetFormatPr|@_defaultRowHeight");
        }
    }

    getLastCellNum() {
        return this.cells.last().value.getColumnIndex();
    }

    getOutlineLevel() {
        // TODO: getOutlineLevel
    }

    getPhysicalNumberOfCells() {
        // TODO: getPhysicalNumberOfCells
    }

    getRowNum() {
        return Number(this.xlsx.getNode("@_r")) - 1;
    }

    getRowStyle() {
        // TODO: getRowStyle
    }

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

    removeCell(cellIndex) {
        this.xlsx.removeNode("c|" + cellIndex);
        this.cells.remove(cellIndex);
    }

    setHeight(height) {
        this.xlsx.setNode("@_ht", height);
        this.xlsx.setNode("@_customHeight", "true");
    }

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