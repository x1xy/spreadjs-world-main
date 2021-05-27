 var fromRange, fromSheet, isFormatPainting = false;

function resetFormatPainting(sheet) {
      isFormatPainting = false;
}

var brushCommand = {
    "brush":{
    		title: "格式刷",
    		iconClass: "ribbon-button-brush",
    		commandName: "brush",
    		execute: async (context, propertyName, fontItalicChecked) => {

                    var spread = context.getWorkbook();
                    var sheet = spread.getActiveSheet();
                    var selectionRange = sheet.getSelections();
    				if (selectionRange.length > 1) {
    					alert("无法对多重选择区域执行此操作");
    					return;
    				}

    				if (isFormatPainting) {
    					resetFormatPainting(sheet);
    					return;
    				}

    				fromRange = selectionRange[0];
    				    fromSheet = sheet;
    					isFormatPainting = true;
    		}
    }
}

var brushGroup = {
    "label": "格式刷",
    "thumbnailClass": "ribbon-button-brush",
    "commandGroup": {
        "children": [
            {
                "direction": "vertical",
                "commands": ["brush"]
            }
        ]
    }
}
