import Workbook from "./Workbook";

class JavascriptXlsx {

    static createWorkbook(option) {
        return new Workbook(option);
    }
}

window.JavascriptXlsx = JavascriptXlsx;