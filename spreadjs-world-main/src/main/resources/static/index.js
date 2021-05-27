var designer
window.onload = function () {
	GC.Spread.Common.CultureManager.culture("zh-cn");
	GC.Spread.Sheets.Themes.Office.bodyFont("宋体")
    designer = initDesigner();
	var spread = designer.getWorkbook();
    var commandManager = spread.commandManager();
    var _executeRotation = {
        canUndo: true,
        execute: function (context, options, isUndo) {
            var Commands = GC.Spread.Sheets.Commands;
            if (isUndo) {
                Commands.undoTransaction(context, options);
                return true;
            } else {
                Commands.startTransaction(context, options);
                var sheet = context.getSheetFromName(options.sheetName);
                sheet.suspendPaint();
                var selections = sheet.getSelections();
                for(var i=0;i<selections.length;i++){
                    sheet.getRange(selections[i].row,
                        selections[i].col,
                        selections[i].rowCount,
                        selections[i].colCount,
                        GC.Spread.Sheets.SheetArea.viewport).textOrientation(options.degree);
                }
                sheet.resumePaint();
                Commands.endTransaction(context, options);
                return true;
            }
        }
    };
    commandManager.register("executeRotation", _executeRotation);
	
    // designer.dispose();
};


		function ungzipString(str) {
			if(!str){
				return "";
			}
			try {
				var restored = pako.ungzip(str, { to: "string" }); //解压
				return restored;
			} catch (err) {
				console.log(err);
			}
			return "";
		}

		function gzipString(str) {
			if(!str){
				return "";
			}
			try {
				var restored = pako.gzip(str, { to: "string" }); //压缩
				return restored;
			} catch (err) {
				console.log(err);
			}
			return "";
		}