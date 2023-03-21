import { _ } from "core-js";
import Utils from "./Utils";
import Xlsx from "./Xlsx";

/**
 * @module Cell
 */
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

    /**
     * @summary 셀에 적용된 스타일을 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(0);
     * var cell = row.createCell(0);
     * var style = workbook.createCellStyle({
     *     font: {
     *         size: 20
     *     },
     *     border: {
     *         left: {
     *             style: "thin",
     *             color: "FF0000"
     *         }
     *     }
     * });
     * cell.setCellValue("Hello");
     * cell.setCellStyle(style);
     * console.log(cell.getCellStyle() === style); // true
     * @returns {CellStyle}
     */
    getCellStyle() {
        const style = this.workbook.styles.get(this.xlsx.getNode("@_s"));
        if (style != null) {
            return style.value;
        }
        return undefined;
    }

    getCellType() {
        return this.xlsx.getNode("@_t");
    }

    /**
     * @summary 셀의 값을 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(0);
     * var cell = row.createCell(0);
     * cell.setCellValue("Hello");
     * console.log(cell.getCellValue()); // Hello
     * @returns {Void}
     */
    getCellValue() {
        if (this.xlsx.getNode("@_t") == "s") {
            return this.workbook.xlsx.getNodes("xl/sharedStrings.xml|sst|si")[this.xlsx.getNode("v")]["t"];
        } else if (this.xlsx.getNode("f") != null) {
            return this.xlsx.getNode("f");
        } else {
            return this.xlsx.getNode("v");
        }
    }

    /**
     * @summary 셀의 Index를 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(0);
     * var cell = row.createCell(3);
     * console.log(cell.getColumnIndex()); // 3
     * @returns {Number}
     */
    getColumnIndex() {
        return Utils.alphabetToIndex(this.xlsx.getNode("@_r").replace(/[0-9]+/, "")) - 1;
    }

    getHyperlink() {
        // TODO: getHyperlink
    }

    /**
     * @summary 셀의 Reference를 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(0);
     * var cell = row.createCell(3);
     * console.log(cell.getReference()); // D1
     * @returns {String}
     */
    getReference() {
        return this.xlsx.getNode("@_r");
    }

    /**
     * @summary 셀이 속한 행을 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(0);
     * var cell = row.createCell(3);
     * console.log(row === cell.getRow()); // true
     * @returns {Row}
     */
    getRow() {
        return this.row;
    }

    /**
     * @summary 셀이 속한 행의 Index를 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(1);
     * var cell = row.createCell(3);
     * console.log(cell.getRowIndex()); // 1
     * @returns {Number}
     */
    getRowIndex() {
        return Number(this.row.xlsx.getNode("@_r")) - 1;
    }

    /**
     * @summary 셀이 속한 시트를 반환한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(1);
     * var cell = row.createCell(3);
     * console.log(sheet === cell.getSheet()); // true
     * @returns {Sheet}
     */
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

    /**
     * @summary 셀을 선택한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(1);
     * var cell = row.createCell(3);
     * cell.setAsActiveCell();
     * console.log(sheet.getActiveCell() === cell.getReference()); // true
     * @returns {Void}
     */
    setAsActiveCell() {
        this.sheet.xlsx.setNode("worksheet|sheetViews|sheetView|selection", {
            "@_activeCell": this.xlsx.getNode("@_r"),
            "@_sqref": this.xlsx.getNode("@_r")
        });
    }

    setCellComment(comment) {
        // TODO: setCellComment
    }

    /**
     * @summary 셀에 스타일을 적용한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(0);
     * var cell = row.createCell(0);
     * var style = workbook.createCellStyle({
     *     font: {
     *         size: 20
     *     },
     *     border: {
     *         left: {
     *             style: "thin",
     *             color: "FF0000"
     *         }
     *     }
     * });
     * cell.setCellValue("Hello");
     * cell.setCellStyle(style);
     * console.log(cell.getCellStyle() === style); // true
     * @param {Object}
     * @returns {Void}
     */
    setCellStyle(style) {
        this.style = style;
        this.xlsx.setNode("@_s", style.styleId.toString());
    }

    /**
     * @summary 셀의 값을 변경한다.
     * @example
     * var workbook = JavascriptXlsx.createWorkbook();
     * var sheet = workbook.createSheet("Sheet1");
     * var row = sheet.createRow(0);
     * var cell = row.createCell(0);
     * cell.setCellValue("Hello");
     * console.log(cell.getCellValue()); // Hello
     * @param {Anything}
     * @returns {Void}
     */
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
        if (this.sheet.xmlRel["@_Target"].lastIndexOf("/") > -1) {
            const filename = this.sheet.xmlRel["@_Target"].substring(this.sheet.xmlRel["@_Target"].lastIndexOf("/") + 1);
            if (this.workbook.xlsx.getNode("xl/worksheets/" + filename + ".rels") == null) {
                this.workbook.xlsx.setNode("xl/worksheets/" + filename + ".rels|?xml", {
                    "@_version": "1.0",
                    "@_encoding": "UTF-8",
                    "@_standalone": "no"
                });
            }

            if (this.workbook.xlsx.getNode("xl/worksheets/" + filename + ".rels|Relationships|@_xmlns") == null) {
                this.workbook.xlsx.appendNode("xl/worksheets/" + filename + ".rels|Relationships", {
                    "@_xmlns": "http://schemas.openxmlformats.org/package/2006/relationships"
                });
            }

            const rels = this.workbook.xlsx.getNodes("xl/worksheets/" + filename + ".rels|Relationships|Relationship");
            let maxRId = 0;
            _.each(rels, rel => {
                const match = rel["@_Id"].match(/rId([0-9]+)/);
                if (match) {
                    maxRId = Math.max(maxRId, Number(match[1]));
                }
            });

            this.workbook.xlsx.appendNode("xl/worksheets/" + filename + ".rels|Relationships|Relationship", {
                "@_Id": "rId" + (maxRId + 1),
                "@_Target": hyperlink.address,
                "@_TargetMode": "External",
                "@_Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
            });

            //
            
            this.sheet.xlsx.afterNodeKey("worksheet|sheetData", "hyperlinks");
            this.sheet.xlsx.appendNode("worksheet|hyperlinks|hyperlink", {
                "@_ref": this.xlsx.getNode("@_r"),
                "@_r:id": "rId" + (maxRId + 1)
            });
        }
    }
}