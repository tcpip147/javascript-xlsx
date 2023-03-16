import _ from 'lodash';

export default class CellStyle {

    // TODO: Styles have to be more implemented.

    constructor(option) {
        this.workbook = option.workbook;

        this.numFmtId = 0;
        this.fontId = 0;
        this.fillId = 0;
        this.borderId = 0;
        this.xfId = 0;
        this.styleId = 0;
    }

    createNodes(styles) {
        for (let key in styles) {
            if (key == "format") {
                this.numFmtId = this.#createNumFmt(styles[key]);
            } else if (key == "font") {
                this.fontId = this.#createFont(styles[key]);
            } else if (key == "fill") {
                this.fillId = this.#createFill(styles[key]);
            } else if (key == "border") {
                this.borderId = this.#createBorder(styles[key]);
            }
        }
        this.styleId = this.#createCellXf(styles);
    }

    #createNumFmt(format) {
        let id = 165;
        const xmlNumFmts = this.workbook.xlsx.getNodes("xl/styles.xml|styleSheet|numFmts|numFmt");
        _.forEach(xmlNumFmts, (xmlNumFmt) => {
            id = Math.max(id, Number(xmlNumFmt["@_numFmtId"]));
        });
        const style = {
            "@_numFmtId": id,
            "@_formatCode": format
        };
        this.workbook.xlsx.appendNode("xl/styles.xml|styleSheet|numFmts|numFmt", style);
        return id;
    }

    #createFont(font) {

        const style = {};
        for (let key in font) {
            if (key == "size") {
                style["sz"] = {
                    "@_val": font[key]
                };
            } else if (key == "name") {
                if (font[key]) {
                    style["name"] = font[key];
                }
            } else if (key == "bold") {
                if (font[key]) {
                    style["b"] = {};
                }
            } else if (key == "italic") {
                if (font[key]) {
                    style["i"] = {};
                }
            } else if (key == "strike") {
                if (font[key]) {
                    style["strike"] = {};
                }
            } else if (key == "color") {
                style["color"] = {
                    "@_rgb": font[key]
                };
            } else if (key == "underline") {
                style["u"] = {
                    "@_val": font[key]
                };
            } else if (key == "script") {
                style["vertAlign"] = {
                    "@_val": font[key]
                };
            }
        }
        this.workbook.xlsx.appendNode("xl/styles.xml|styleSheet|fonts|font", style);
        return this.workbook.xlsx.getNodes("xl/styles.xml|styleSheet|fonts|font").length - 1;
    }

    #createFill(fill) {
        const style = {};
        for (let key in fill) {
            if (key == "pattern") {
                style["patternFill"] = {};
                if (fill[key]["foregroundColor"]) {
                    style["patternFill"]["fgColor"] = {
                        "@_rgb": fill[key]["foregroundColor"]
                    };
                }
                if (fill[key]["backgroundColor"]) {
                    style["patternFill"]["bgColor"] = {
                        "@_rgb": fill[key]["backgroundColor"]
                    };
                }
                if (fill[key]["type"]) {
                    style["patternFill"]["@_patternType"] = fill[key]["type"];
                }
            } else if (key == "gradient") {

            }
        }
        this.workbook.xlsx.appendNode("xl/styles.xml|styleSheet|fills|fill", style);
        return this.workbook.xlsx.getNodes("xl/styles.xml|styleSheet|fills|fill").length - 1;
    }

    #createBorder(border) {
        const style = {};
        if (border["left"]) {
            style["left"] = {};
            if (border["left"]["style"]) {
                style["left"]["@_style"] = border["left"]["style"];
            }
            if (border["left"]["color"]) {
                style["left"]["color"] = {
                    "@_rgb": border["left"]["color"]
                };
            }
        }
        if (border["right"]) {
            style["right"] = {};
            if (border["right"]["style"]) {
                style["right"]["@_style"] = border["right"]["style"];
            }
            if (border["right"]["color"]) {
                style["right"]["color"] = {
                    "@_rgb": border["right"]["color"]
                };
            }
        }
        if (border["top"]) {
            style["top"] = {};
            if (border["top"]["style"]) {
                style["top"]["@_style"] = border["top"]["style"];
            }
            if (border["top"]["color"]) {
                style["top"]["color"] = {
                    "@_rgb": border["top"]["color"]
                };
            }
        }
        if (border["bottom"]) {
            style["bottom"] = {};
            if (border["bottom"]["style"]) {
                style["bottom"]["@_style"] = border["bottom"]["style"];
            }
            if (border["bottom"]["color"]) {
                style["bottom"]["color"] = {
                    "@_rgb": border["bottom"]["color"]
                };
            }
        }
        if (border["diagonal"]) {
            style["diagonal"] = {};
            if (border["diagonal"]["style"]) {
                style["diagonal"]["@_style"] = border["diagonal"]["style"];
            }
            if (border["diagonal"]["color"]) {
                style["diagonal"]["color"] = {
                    "@_rgb": border["diagonal"]["color"]
                };
            }
            if (border["diagonal"]["direction"] == "up") {
                style["@_diagonalUp"] = "1";
            } else if (border["diagonal"]["direction"] == "down") {
                style["@_diagonalDown"] = "1";
            }
        }
        this.workbook.xlsx.appendNode("xl/styles.xml|styleSheet|borders|border", style);
        return this.workbook.xlsx.getNodes("xl/styles.xml|styleSheet|borders|border").length - 1;
    }

    #createCellXf(styles) {
        const xf = {
            "@_borderId": this.borderId,
            "@_fillId": this.fillId,
            "@_fontId": this.fontId,
            "@_numFmtId": this.numFmtId,
            "@_xfId": this.xfId,
        };
        this.workbook.xlsx.appendNode("xl/styles.xml|styleSheet|cellXfs|xf", xf);

        if (this.numFmtId > 0) {
            xf["@_applyNumberFormat"] = "true";
        }

        if (styles["alignment"]) {
            xf["@_applyAlignment"] = "true";
            xf["alignment"] = {};
            if (styles["alignment"]["horizontal"]) {
                xf["alignment"]["@_horizontal"] = styles["alignment"]["horizontal"];
            }
            if (styles["alignment"]["vertical"]) {
                xf["alignment"]["@_vertical"] = styles["alignment"]["vertical"];
            }
        }
        return this.workbook.xlsx.getNodes("xl/styles.xml|styleSheet|cellXfs|xf").length - 1;
    }
}