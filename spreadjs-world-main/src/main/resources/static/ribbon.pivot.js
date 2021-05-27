
var PivotPanelVisible = "PivotPanelVisible";

function showPivotPanel(context){
            var spread = context.getWorkbook();
            var sheet = spread.getActiveSheet();
            if(sheet.pivotTables.all().length){
                var pivotTable = sheet.pivotTables.all()[0];
                var panel = new GC.Spread.Pivot.PivotPanel("myPivotPanel", pivotTable, pivotPanel);
                pivotPanel.style.display = "block"
                designer.setData(PivotPanelVisible, true)
                context.refresh();
            }
            else{
                pivotPanel.style.display = "none";
            }
}

var ribbonPivotCommands = {
    "showPivotPanel": {
        text: "字段列表",
        iconClass: "ribbon-button-namemanager",
        bigButton: true,
        type: "checkbox",
        commandName: "showPivotPanel",
        execute: async (context) => {
            var pivotPanel = document.getElementById("pivotPanel");
            if(pivotPanel.style.display == "block"){
                pivotPanel.style.display = "none";
                designer.setData(PivotPanelVisible, false)
                return
            }
            showPivotPanel(context);

        },
        getState: (context) => {
            return !!context.getData(PivotPanelVisible);
        }
    },
    "newPivotTable": {
        text: "新建",
        iconClass: "ribbon-button-namemanager",
        bigButton: true,
        commandName: "newPivotTable",
        execute: async (context) => {


            var spread = context.getWorkbook();
            var sheet = spread.getActiveSheet();
            var sheetIndex = spread.getSheetIndex(sheet.name())
            var range = "'" + sheet.name() + "'!"+ GC.Spread.Sheets.CalcEngine.rangeToFormula(sheet.getSelections()[0], 0, 0, GC.Spread.Sheets.CalcEngine.RangeReferenceRelative.allAbsolute, false);


            var dialogOption = {
                range: range,
            };
            GC.Spread.Sheets.Designer.showDialog("newPivotTable", dialogOption, (result) => {
                if(result && result.range){
                    var newSheet;
                    try{
                        spread.addSheet(sheetIndex);
                        newSheet = spread.getSheet(sheetIndex);
                        var pivotTable = newSheet.pivotTables.add("PivotTable_" + newSheet.name() + (sheet.pivotTables.all().length + 1), result.range, 1, 1, GC.Spread.Pivot.PivotTableLayoutType.outline, GC.Spread.Pivot.PivotTableThemes.medium2);
                        spread.setActiveSheetIndex(sheetIndex)
                        showPivotPanel(context);
                    }
                    catch(ex){
                        if(newSheet){
                            spread.removeSheet(spread.getSheetIndex(newSheet.name()))
                            GC.Spread.Sheets.Designer.showMessageBox("添加失败", "错误", GC.Spread.Sheets.Designer.MessageBoxIcon.error);
                        }
                    }
                }
            });

        }
    }
}


var newPivotTableTemplate = {
    title: "创建数据透视表",
    content: [
        {
            type: "FlexContainer",
            children: [
                {
                    type: "TextBlock",
                    margin: "5px -4px",
                    text: "请选择要分析的数据"
                },
                {
                    type: "RangeSelect",
                    title: "表/区域：",
                    absoluteReference: true,
                    needSheetName: true,
                    margin: "5px -5px",
                    bindingPath: "range"
                }
            ]
        }
    ]
};
GC.Spread.Sheets.Designer.registerTemplate("newPivotTable", newPivotTableTemplate);



var ribbonPivotConfig = {
    "label": "数据透视表",
    "thumbnailClass": "ribbon-thumbnail-spreadsettings",
    "commandGroup": {
        "children": [
            {
                "direction": "vertical",
                "commands": ["newPivotTable"]
            },
            {
                "direction": "vertical",
                "commands": ["showPivotPanel"]
            }
        ]
    }
}