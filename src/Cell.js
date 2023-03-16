import Utils from "./Utils";
import Xlsx from "./Xlsx";

export default class Cell {

    constructor(option) {
        this.workbook = option.workbook;
        this.sheet = option.sheet;
        this.row = option.row;
        this.xmlCell = option.xmlCell;
        
        this.xlsx = new Xlsx(this.xmlCell);
        this.style;
    }

    copyCellFrom(srcCell, policy) {
        // TODO: copyCellFrom
    }

    getArrayFormulaRange() {
        // TODO: getArrayFormulaRange
    }

    getCellComment() {
        // TODO: getCellComment
    }

    getCellStyle() {
        return this.workbook.styles.get(this.xlsx.getNode("@_s")).value;
    }

    getCellType() {
        return this.xlsx.getNode("@_t");
    }

    getCellValue() {
        if (this.xlsx.getNode("@_t") == "s") {
            return this.workbook.xlsx.getNodes("xl/sharedStrings.xml|sst|si")[this.xlsx.getNode("v")]["t"];
        } else if (this.xlsx.getNode("f") != null) {
            return this.xlsx.getNode("f");
        } else {
            return this.xlsx.getNode("v");
        }
    }

    getColumnIndex() {
        return Utils.alphabetToIndex(this.xlsx.getNode("@_r").replace(/[0-9]+/, "")) - 1;
    }

    getHyperlink() {
        // TODO: getHyperlink
    }

    getReference() {
        return this.xlsx.getNode("@_r");
    }

    getRow() {
        return this.row;
    }

    getRowIndex() {
        return Number(this.row.xlsx.getNode("@_r")) - 1;
    }

    getSheet() {
        return this.sheet;
    }

    isPartOfArrayFormulaGroup() {
        // TODO: isPartOfArrayFormulaGroup
    }

    removeCellComment() {
        // TODO: removeCellComment
    }

    removeHyperlink() {
        // TODO: removeHyperlink
    }

    setAsActiveCell() {
        this.sheet.xlsx.setNode("worksheet|sheetViews|sheetView|selection", {
            "@_activeCell": this.xlsx.getNode("@_r"),
            "@_sqref": this.xlsx.getNode("@_r")
        });
    }

    setCellComment(comment) {
        // TODO: setCellComment
    }

    setCellStyle(style) {
        this.style = style;
        this.xlsx.setNode("@_s", style.styleId.toString());
    }

    setCellValue(value) {
        if (isNaN(value)) {
            if (value.substring(0, 1) == "=") {
                this.xlsx.setNode("f", value);
            } else {
                if (this.workbook.sharedStrings[value] == null) {
                    this.workbook.xlsx.appendNode("xl/sharedStrings.xml|sst|si", {
                        "t": value
                    });
                    this.workbook.sharedStrings[value] = this.workbook.xlsx.getNodes("xl/sharedStrings.xml|sst|si").length - 1;
                }
                this.xlsx.setNode("@_t", "s");
                this.xlsx.setNode("v", this.workbook.sharedStrings[value].toString());
            }
        } else {
            this.xlsx.setNode("@_t", "n");
            this.xlsx.setNode("v", value);
        }
    }

    setHyperlink(hyperlink) {
        // TODO: setHyperlink
    }
}