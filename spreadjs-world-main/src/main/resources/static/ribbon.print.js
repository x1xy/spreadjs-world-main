

async function setFitZoomFactor(designer) {
    var spread = designer.getWorkbook();
    var sheet = spread.getActiveSheet(), sheetIndex = spread.getSheetIndex(sheet.name());
    var sheetPageInfo = spread.pageInfo(sheetIndex);
    if (sheetPageInfo && sheetPageInfo.pages && sheetPageInfo.pages.length > 1) {
        var pages = sheetPageInfo.pages;
        var remainPageWide = 0;
        for (var pageIndex = 0; pageIndex < pages.length; pageIndex++) {
            if(pages[pageIndex].row === pages[0].row){
                for (var col = pages[pageIndex].column; col < pages[pageIndex].column + pages[pageIndex].columnCount; col++) {
                    remainPageWide += sheet.getColumnWidth(col);
                }
            }
        }
        var printInfo = sheet.printInfo();
        var printWidth = printInfo.paperSize().width() - (printInfo.margin() ? printInfo.margin().left + printInfo.margin().right : 140)
        if(printInfo.showRowHeader() !== 1){
            var rowHeaderWidth = 0;
            for(var i = 0; i < sheet.getColumnCount(GC.Spread.Sheets.SheetArea.rowHeader); i++){
                rowHeaderWidth += sheet.getColumnWidth(i, GC.Spread.Sheets.SheetArea.rowHeader);
            }
        }
        var zoom = printWidth / (rowHeaderWidth + remainPageWide) - 0.05;
        printInfo.zoomFactor(zoom.toFixed(1));
        sheet.printInfo(printInfo);
    }
}

var ribbonPrintCommands = {
    "showHidePrintLine": {
        text: "打印分页预览线",
        type: "checkbox",
        commandName: "showHidePrintLine",
        execute: async (context) => {
            setTimeout(() => {
                let sheet = context.Spread.getActiveSheet();
                var isVisible = sheet.isPrintLineVisible();
                sheet.isPrintLineVisible(!isVisible);
            }, 0)
        },
        getState: (context) => {
            let sheet = context.Spread.getActiveSheet();
            return sheet.isPrintLineVisible();
        }
    },
    "insertPageBreak": {
        iconClass: "ribbon-button-namemanager",
        text: "插入分页符",
        commandName: "insertPageBreak",
        execute: async (context) => {
            let sheet = context.Spread.getActiveSheet();
            sheet.setColumnPageBreak(sheet.getActiveColumnIndex(), true)
            sheet.setRowPageBreak(sheet.getActiveRowIndex(), true)
        }
    },
    "deletePageBreak": {
        iconClass: "ribbon-button-namemanager",
        text: "删除分页符",
        commandName: "deletePageBreak",
        execute: async (context) => {
            let sheet = context.Spread.getActiveSheet();
            sheet.setColumnPageBreak(sheet.getActiveColumnIndex(), false)
            sheet.setRowPageBreak(sheet.getActiveRowIndex(), false)
        }
    },
    "deleteAllPageBreak": {
        iconClass: "ribbon-button-namemanager",
        text: "删除所有分页符",
        commandName: "deleteAllPageBreak",
        execute: async (context) => {
            let sheet = context.Spread.getActiveSheet();
            sheet.suspendPaint()
            for (var row = 0; row < sheet.getRowCount(); row++) {
                sheet.setRowPageBreak(row, false)
            }
            for (var col = 0; col < sheet.getColumnCount(); col++) {
                sheet.setColumnPageBreak(col, false)
            }
            sheet.resumePaint();
        }
    },
    "setFitPagesWide": {
        text: "自动",
        type: "comboBox",
        commandName: "setFitPagesWide",
        dropdownList: [
            {
                text: "自动",
                value: "-1"
            },
            {
                text: "1",
                value: "1"
            },
            {
                text: "2",
                value: "2"
            }
        ],
        execute: async (context, selectValue) => {
            if (selectValue != null && selectValue != undefined) {
                setTimeout(() => {
                    let sheet = context.Spread.getActiveSheet();
                    var printInfo = sheet.printInfo();
                    printInfo.fitPagesWide(parseInt(selectValue));
                    sheet.printInfo(printInfo)
                }, 0);
            }
        },
        getState: (context) => {
            let sheet = context.Spread.getActiveSheet();
            var printInfo = sheet.printInfo()
            var value = printInfo.fitPagesWide();
            return value.toString();
        }

    },
    "setFitPagesTall": {
        text: "高度",
        type: "comboBox",
        commandName: "setFitPagesTall",
        dropdownList: [
            {
                text: "自动",
                value: "-1"
            },
            {
                text: "1",
                value: "1"
            },
            {
                text: "2",
                value: "2"
            }
        ],
        execute: async (context, selectValue) => {
            if (selectValue != null && selectValue != undefined) {
                setTimeout(() => {
                    let sheet = context.Spread.getActiveSheet();
                    var printInfo = sheet.printInfo();
                    printInfo.fitPagesTall(parseInt(selectValue));
                    sheet.printInfo(printInfo)
                }, 0);
            }
        },
        getState: (context) => {
            let sheet = context.Spread.getActiveSheet();
            var printInfo = sheet.printInfo()
            var value = printInfo.fitPagesTall();
            return value.toString();
        }
    },
    "setFitZoomFactor": {
        iconClass: "ribbon-button-namemanager",
        //                        bigButton: true,
        text: "强制分页自适应",
        commandName: "setFitZoomFactor",
        execute: setFitZoomFactor,
    },
    "resetZoomFactor": {
        iconClass: "ribbon-button-namemanager",
        //                        bigButton: true,
        text: "重置打印缩放",
        commandName: "resetZoomFactor",
        execute: async function (designer) {
            var spread = designer.getWorkbook();
            var sheet = spread.getActiveSheet();
            var printInfo = sheet.printInfo();
            printInfo.zoomFactor(1);
            sheet.printInfo(printInfo);
        },
    },
}


var ribbonPrintConfig = {
    "label": "页面布局",
    "thumbnailClass": "ribbon-thumbnail-spreadsettings",
    "commandGroup": {
        "children": [
            {
                "direction": "vertical",
                "commands": ["showHidePrintLine"]
            },
            {
                "type": "separator"
            },
            {
                "direction": "vertical",
                "commands": [
                    "insertPageBreak",
                    "deletePageBreak",
                    "deleteAllPageBreak"
                ]
            },
            {
                "type": "separator"
            },
            {
                "direction": "vertical",
                "commands": [
                    {
                        iconClass: "ribbon-button-namemanager",
                        text: "宽度",
                    },
                    {
                        iconClass: "ribbon-button-namemanager",
                        text: "高度",
                    }
                ]
            },
            {
                "direction": "vertical",
                "commands": [
                    "setFitPagesWide",
                    "setFitPagesTall"
                ]
            },
            {
                "type": "separator"
            },
            {
                "direction": "vertical",
                "commands": [
                    "setFitZoomFactor",
                    "resetZoomFactor"
                ]
            }
        ]
    }
};
