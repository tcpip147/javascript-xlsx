이 프로젝트는 ApachePOI 인터페이스를 참고하여 만들어졌습니다.

현재는 텍스트 수정 및 일부 스타일만 구현되어 있습니다.

지속적으로 기능을 추가할 예정입니다.

### Getting Stared

```
<script src="javascript-xlsx.js"></script>
```

```
<!-- IE11 브라우저 호환 버전 -->
<script src="javascript-xlsx-ie11.js"></script>
```

### Example

```
var workbook = JavascriptXlsx.createWorkbook();
var sheet = workbook.createSheet("Sheet1");

for (var r = 0; r < 10; r++) {
    var row = sheet.createRow(r);
    var style = workbook.createCellStyle({
        font: {
            size: 10 + r
        }
    });
    for (var c = 0; c < 5; c++) {
        var cell = row.createCell(c);
        cell.setCellValue("row" + r + " : col" + c);
        cell.setCellStyle(style);
    }
}

for (var c = 0; c < 5; c++) {
    sheet.setColumnWidth(c, 30);
}

workbook.write("example.xlsx");
```

### API

<a href="https://tcpip147.github.io/javascript-xlsx/index.html">API Reference</a>를 참조하십시요.

<h3 id="cellStyleAttributes">CellStyle Attributes</h3>

```
{
    /* 표시형식 */
    format: string,
    /* 글꼴 */
    font: {
        name: string // 글꼴명
        size: number, // 크기
        bold: boolean, // 굵게
        italic: boolean, // 기울임꼴
        strike: boolean, // 취소선
        color: [0-9A-F]{8}, // 색상
        /* 첨자 */
        script: superscript // 위 첨자
              | subscript // 아래 첨자
        /* 밑줄 */
        underline: none // 없음
                 | single // 실선
                 | double // 이중 실선
                 | singleAccounting // 실선(회계용)
                 | doubleAccounting // 이중 실선(회계용)
    },
    /* 채우기 */
    fill: {
        pattern : {
            foregroundColor: [0-9A-F]{8}, // 무늬색
            backgroundColor: [0-9A-F]{8}, // 배경색
            /* 무늬스타일 */
            type: none // 없음
                | solid // 단색
                | darkGray // 75% 회색
                | mediumGray // 50% 회색
                | lightGray // 25% 회색
                | gray125 // 12.5% 회색
                | gray0625 // 6.25% 회색
                | darkHorizontal // 가로줄
                | darkVertical // 세로줄
                | darkDown // 역대각선줄
                | darkUp // 대각선줄
                | darkGrid // 대각선교차무늬
                | darkTrellis // 굵은 실선 대각선교차무늬
                | lightHorizontal // 가는 실선 가로줄
                | lightVertical // 가는 실선 세로줄
                | lightDown // 가는 실선 역대각선줄
                | lightUp // 가는 실선 대각선줄
                | lightGrid // 가는 실선 가로교차무늬
                | lightTrellis // 가는 실선 대각선교차무늬
        },
        /* 채우기 효과 */
        gradient : {
            // 미구현
        }
    },
    /* 테두리 */
    border: {
        left: {
            style: thin
                 | medium
                 | dashed
                 | dotted
                 | thick
                 | double
                 | hair
                 | mediumDashed
                 | dashDotDot
                 | mediumDashDotDot
                 | slantDashDot,
            color: [0-9A-F]{8}
        },
        right: {
            style: thin
                 | medium
                 | dashed
                 | dotted
                 | thick
                 | double
                 | hair
                 | mediumDashed
                 | dashDotDot
                 | mediumDashDotDot
                 | slantDashDot,
            color: [0-9A-F]{8}
        },
        bottom: {
            style: thin
                 | medium
                 | dashed
                 | dotted
                 | thick
                 | double
                 | hair
                 | mediumDashed
                 | dashDotDot
                 | mediumDashDotDot
                 | slantDashDot,
            color: [0-9A-F]{8}
        },
        top: {
            style: thin
                 | medium
                 | dashed
                 | dotted
                 | thick
                 | double
                 | hair
                 | mediumDashed
                 | dashDotDot
                 | mediumDashDotDot
                 | slantDashDot,
            color: [0-9A-F]{8}
        },
        /* 대각선 */
        diagonal: {
            style: thin
                 | medium
                 | dashed
                 | dotted
                 | thick
                 | double
                 | hair
                 | mediumDashed
                 | dashDotDot
                 | mediumDashDotDot
                 | slantDashDot,
            color: [0-9A-F]{8},
            direction: up|down
        }
    },
    /* 맞춤 */
    alignment: {
        /* 가로정렬 */
        horizontal: general // 일반
                  | left // 왼쪽(들여쓰기)
                  | center // 가운데
                  | right // 오른쪽(들여쓰기)
                  | fill // 채우기
                  | justify // 양쪽맞춤
                  | centerContinuous // 선택영역의 가운데로
                  | distributed, // 균등분할(들여쓰기)
        /* 세로정렬 */
        vertical: top // 위쪽
                | center // 가운데
                | bottom // 아래쪽
                | justify // 양쪽맞춤
                | distributed // 균등분할
    }
}
```
