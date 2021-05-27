async function addChart(designer, type) {
    setTimeout(function () {
        var spread = designer.getWorkbook();
        var sheet = spread.getActiveSheet();
        // var range = sheet.getSelections();
        // var rangeStr = GC.Spread.Sheets.CalcEngine.rangeToFormula(range[0]);
        // sheet.charts.add("", GC.Spread.Sheets.Charts.ChartType[type], 0, 100, 400, 300, rangeStr);

        var commandManager = spread.commandManager();
        commandManager.execute({cmd: "customAddChart", sheetName: sheet.name(), chartType: type});

    },10)
}

var ribbonChartCommands = {
    "columnChart": {
        iconClass: "ribbon-button-column-little",
        type: "dropdown",
        commandName: "chart_List",
        dropdownList:[
            {
                groupName: "二维柱形图",
                groupItems: [
                    {
                        iconClass: "ribbon-button-columnClustered",
                        iconHeight: 50,
                        iconWidth: 50,
                        tip: "簇状柱形图",
                        value: "columnClustered"
                    },
                    {
                        iconClass: "ribbon-button-columnStacked",
                        iconHeight: 50,
                        iconWidth: 50,
                        tip: "堆积柱形图",
                        value: "columnStacked"
                    },
                    {
                        iconClass: "ribbon-button-columnStacked100",
                        iconHeight: 50,
                        iconWidth: 50,
                        tip: "百分比堆积柱形图",
                        value: "columnStacked100"
                    }
                ]
            },
            {
                groupName: "二维条形图",
                groupItems: [
                    {
                        iconClass: "ribbon-button-barClustered",
                        iconHeight: 50,
                        iconWidth: 50,
                        tip: "簇状条形图",
                        value: "barClustered"
                    },
                    {
                        iconClass: "ribbon-button-barStacked",
                        iconHeight: 50,
                        iconWidth: 50,
                        tip: "堆积条形图",
                        value: "barStacked"
                    },
                    {
                        iconClass: "ribbon-button-barStacked100",
                        iconHeight: 50,
                        iconWidth: 50,
                        tip: "百分比堆积条形图",
                        value: "barStacked100"
                    }
                ]
            }
        ],
        dropdownMaxHeight: 500,
        dropdownMaxWidth: 260,
        execute: async function (context, selectValue) {
            if(!selectValue) return;
            addChart(context, selectValue);
        }
    },
    "lineChart": {
        iconClass: "ribbon-button-line-little",
        type: "dropdown",
        commandName: "chart_List",
        dropdownList:[
            {
                groupName: "折线图",
                groupItems: [
                    {
                        iconClass: "ribbon-button-line",
                        iconHeight: 50,
                        iconWidth: 50,
                        tip: "折线图",
                        value: "line"
                    },
                    {
                        iconClass: "ribbon-button-lineStacked",
                        iconHeight: 50,
                        iconWidth: 50,
                        tip: "堆积折线图",
                        value: "lineStacked"
                    },
                    {
                        iconClass: "ribbon-button-lineMarkers",
                        iconHeight: 50,
                        iconWidth: 50,
                        tip: "数据点折线图",
                        value: "lineMarkers"
                    }
                ]
            },
            {
                groupName: "面积图",
                groupItems: [
                    {
                        iconClass: "ribbon-button-area",
                        iconHeight: 50,
                        iconWidth: 50,
                        tip: "面积图",
                        value: "area"
                    }
                ]
            }
        ],
        dropdownMaxHeight: 500,
        dropdownMaxWidth: 260,
        execute: async function (context, selectValue) {
            if(!selectValue) return;
            addChart(context, selectValue);
        }
    }
}

