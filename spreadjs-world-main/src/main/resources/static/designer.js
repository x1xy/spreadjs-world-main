const FirstLoad = "firstLoad"


var designerConfig = JSON.parse(JSON.stringify(GC.Spread.Sheets.Designer.DefaultConfig))
var customerRibbon = {
    "id": "operate",
    "text": "操作",
    "buttonGroups": [
    ],
//    visibleWhen: FirstLoad
}
var layoutRibbon = {
    "id": "layout",
    "text": "页面布局",
    "buttonGroups": [
    ]
}

var pivotRibbon = {
    "id": "pivot",
    "text": "数据透视表分析",
    "buttonGroups": [
    ]
}

var tabRibbon = {
    "id": "tab",
    "text": "名称标签",
    "buttonGroups": [
    ]
}

function initStates(designer) {
    spreadTemplateJSON = undefined;
    designer.setData(AllowBindingData, true)
    designer.setData(AllowBackTemplate, false)

    setTimeout(()=>
    designer.setData(FirstLoad, true), 500)
//    designer.refresh()
//    designer.setData(FirstLoad, false)
}

function initEvents(designer){
    var spread = designer.getWorkbook();
    spread.setSheetCount(1);
//    spread.commandManager().addListener('Designer.mergeCenter', function() {
//             var sheet = spread.getActiveSheet();
//
//             sheet.removeSpan(sheet.getActiveRowIndex(), sheet.getActiveColumnIndex(), GC.Spread.Sheets.SheetArea.viewport);
//
//            });
    spread.bind(GC.Spread.Sheets.Events.ActiveSheetChanged, ()=>{

        var pivotPanel = document.getElementById("pivotPanel");
            pivotPanel.style.display = "none";
        designer.setData(PivotPanelVisible, false)
    })



    			spread.bind(GC.Spread.Sheets.Events.SelectionChanged, function(e, info) {
    				if (isFormatPainting) {
    					var sheet = spread.getActiveSheet();
    					resetFormatPainting(sheet);
    					sheet.isPaintSuspended(true);
    					var toRange = sheet.getSelections()[0];


    					//toRange biger than fromRange
    					if (fromRange.rowCount > toRange.rowCount) {
    						toRange.rowCount = fromRange.rowCount;
    					}
    					if (fromRange.colCount > toRange.colCount) {
    						toRange.colCount = fromRange.colCount;
    					}
    					//toRange must in Sheet
    					if (toRange.row + toRange.rowCount > sheet.getRowCount()) {
    						toRange.rowCount = sheet.getRowCount() - toRange.row;
    					}
    					if (toRange.col + toRange.colCount > sheet.getColumnCount()) {
    						toRange.colCount = sheet.getColumnCount() - toRange.col;
    					}

    					var rowStep = fromRange.rowCount,
    						colStep = fromRange.colCount;
    					var endRow = toRange.row + toRange.rowCount - 1,
    						endCol = toRange.col + toRange.colCount - 1;

    					// if toRange bigger than fromRange, repeat paint
    					for (var startRow = toRange.row; startRow <= endRow; startRow = startRow + rowStep) {
    						for (var startCol = toRange.col; startCol <= endCol; startCol = startCol + colStep) {

    							var rowCount = startRow + rowStep > endRow + 1 ? endRow - startRow + 1 : rowStep;
    							var colCount = startCol + colStep > endCol + 1 ? endCol - startCol + 1 : colStep;
    							// sheet.copyTo(fromRange.row,fromRange.col, startRow, startCol, rowCount, colCount,GC.Spread.Sheets.CopyToOptions.style | GC.Spread.Sheets.CopyToOptions.span);
    							var fromRanges = new GC.Spread.Sheets.Range(fromRange.row, fromRange.col, rowCount, colCount);
    							var pastedRange = new GC.Spread.Sheets.Range(startRow, startCol, rowCount, colCount);
    							spread.commandManager().execute({
    								cmd: "clipboardPaste",
    								sheetName: sheet.name(),
    								fromSheet: fromSheet,
    								fromRanges: [fromRanges],
    								pastedRanges: [pastedRange],
    								isCutting: false,
    								clipboardText: "",
    								pasteOption: GC.Spread.Sheets.ClipboardPasteOptions.formatting
    							});
    						}
    					}

    					sheet.isPaintSuspended(false);
    				}
           		});


           		document.addEventListener('mousemove', function (e) {
                				var e = e ? e : window.event;
                				if (!isFormatPainting) {
                					$("#brushIcon").hide();
                					return;
                				}
                				var posx = e.pageX;
                				var posy = e.pageY;

                				var offset = $("#ssDesigner").offset();
                				if (posx > offset.left && posy > offset.top &&
                					posx < offset.left + $("#ssDesigner").width() && posy < offset.top + $("#ssDesigner").height()) {
                					var brushIcon = $("#brushIcon");
                					brushIcon.show();
                					brushIcon.css("left", posx + 12 + "px");
                					brushIcon.css("top", posy + "px");
                				} else {
                					$("#brushIcon").hide();
                				}

                			})
}

function initCommands(designer){
    var spread = designer.getWorkbook();

    var addChartCommand = {
        canUndo: true,
        execute: function (context, options, isUndo) {
            var Commands = GC.Spread.Sheets.Commands;
            if (isUndo) {
                Commands.undoTransaction(context, options);
                return true;
            } else {
                Commands.startTransaction(context, options);
                var sheet = context.getSheetFromName(options.sheetName);
                var range = sheet.getSelections();
                var rangeStr = GC.Spread.Sheets.CalcEngine.rangeToFormula(range[0]);
                var type = options.chartType;
                sheet.charts.add("", GC.Spread.Sheets.Charts.ChartType[type], 0, 100, 400, 300, rangeStr);
                Commands.endTransaction(context, options);
                return true;
            }
        }
    };

    var commandManager = spread.commandManager();
    commandManager.register("customAddChart", addChartCommand);
}

function initDesigner() {

    customerRibbon.buttonGroups.push(ribbonFileConfig);
    customerRibbon.buttonGroups.push(ribbonDataConfig);
    layoutRibbon.buttonGroups.push(ribbonPivotConfig1);
    designerConfig.ribbon[0].buttonGroups.unshift(brushGroup);


    layoutRibbon.buttonGroups.push(ribbonPrintConfig);
    pivotRibbon.buttonGroups.push(ribbonPivotConfig);
//    tabRibbon.buttonGroups.push(ribbonPivotConfig1);

    designerConfig.commandMap = {}

    Object.assign(designerConfig.commandMap, ribbonChartCommands)
    //注册命令
    Object.assign(designerConfig.commandMap, ribbonChartCommands)

    Object.assign(designerConfig.commandMap, ribbonFileCommands, ribbonPrintCommands)
    Object.assign(designerConfig.commandMap, ribbonPivotCommands, ribbonDataCommands)
    Object.assign(designerConfig.commandMap, ribbonTabCommands)
    Object.assign(designerConfig.commandMap, brushCommand)
    Object.assign(designerConfig.commandMap, sidePanelsPictureCommands)
    Object.assign(designerConfig.commandMap, ribbonViewCommands );


    designerConfig.ribbon.push(customerRibbon);
    designerConfig.ribbon.push(layoutRibbon);
    designerConfig.ribbon.push(pivotRibbon);
//    designerConfig.ribbon.push(tabRibbon);
    designerConfig.sidePanels.push(sidePanelsPictureConfig)
    // designerConfig.contextMenu.push("myCmd");

    //视图：滚动条
    designerConfig.ribbon[4].buttonGroups[0].commandGroup.children.push(
        {
            type: "separator"
        },
        {
            commands:["showHideScrollbar"]
        }
    )


    designerConfig.ribbon[1].buttonGroups[1].commandGroup = {};
        designerConfig.ribbon[1].buttonGroups[1].commandGroup.children = [
            {
                commands: ["insertChart"]
            },
            {
                commands: ["columnChart", "lineChart"],
                direction: "vertical"
            },
        ]
    var fontFamilyCommand = GC.Spread.Sheets.Designer.getCommand(GC.Spread.Sheets.Designer.CommandNames.FontFamily);

    var command1 = GC.Spread.Sheets.Designer.getCommand(GC.Spread.Sheets.Designer.CommandNames.MergeCenter);


    fontFamilyCommand.dropdownList.push({
        text: "微软雅黑",
        value: "微软雅黑"
    });
    fontFamilyCommand.dropdownList.push({
        text: "华文楷体",
        value: "华文楷体"
    });


    designerConfig.commandMap[GC.Spread.Sheets.Designer.CommandNames.FontFamily] = fontFamilyCommand

    var formatDialogTemplate = GC.Spread.Sheets.Designer.getTemplate(GC.Spread.Sheets.Designer.TemplateNames.FormatDialogTemplate)
    formatDialogTemplate.content[0].children[2].children[0].children[0].children[0].children[1].items.push({text: "微软雅黑", value: "微软雅黑"})
    GC.Spread.Sheets.Designer.registerTemplate(GC.Spread.Sheets.Designer.TemplateNames.FormatDialogTemplate, formatDialogTemplate)

    var fileMenuPanelTemplate =  GC.Spread.Sheets.Designer.getTemplate(GC.Spread.Sheets.Designer.TemplateNames.FileMenuPanelTemplate)
    fileMenuPanelTemplate.content[0].children[0].children[0].children[0].children[1].items.push({text: "保存", value: "Save"})
    GC.Spread.Sheets.Designer.registerTemplate(GC.Spread.Sheets.Designer.TemplateNames.FileMenuPanelTemplate, fileMenuPanelTemplate)


    //    designerConfig.quickAccessBar = ["reset",{iconClass: "gc-accessbar-save",
    //                                                            commandName: "quickSave",
    //                                                            execute: async function(context){
    //                                                                alert("ss")
    //                                                            }}]

    initContextMenu();
    var designer = new GC.Spread.Sheets.Designer.Designer(
        document.getElementById('ssDesigner'), /**designer host */
        designerConfig, // designerConfigJson /**If you want to change Designer's layout or rewrite it, you can pass in your own Config JSON */ ,
        undefined /**If you want to use the spread you already created instead of generating a new spread, pass in */
    );
    $(window).bind('beforeunload',function(){
        return "";
    });
    initStates(designer);
    initEvents(designer);
    initCommands(designer);
    return designer
}




