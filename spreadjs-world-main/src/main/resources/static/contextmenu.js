function initContextMenu(){
    designerConfig.commandMap["textRotation"]= {
        text: "文字旋转",
        commandName: "textRotation",
        subCommands:[
            "negativeVertical",
            "horizontal",
            "vertical"
        ],
        visibleContext:"ClickViewport"
    }

    designerConfig.commandMap["negativeVertical"]= {
        text:"-90度",
        commandName: "negativeVertical",
        //enableContext: "allowEditObject",
        execute: async (context) => {
            var spread = context.getWorkbook();
            var commandManager = spread.commandManager();
            commandManager.execute({cmd: "executeRotation", sheetName: spread.getActiveSheet().name(),degree:-90});
        }
    }

    designerConfig.commandMap["horizontal"]= {
        text:"水平",
        commandName: "horizontal",
        //enableContext: "allowEditObject",
        execute: async (context) => {
            var spread = context.getWorkbook();
            var commandManager = spread.commandManager();
            commandManager.execute({cmd: "executeRotation", sheetName: spread.getActiveSheet().name(),degree:0});
        }
    }

    designerConfig.commandMap["vertical"]= {
        text:"90度",
        commandName: "vertical",
        //enableContext: "allowEditObject",
        execute: async (context) => {
            var spread = context.getWorkbook();
            var commandManager = spread.commandManager();
            commandManager.execute({cmd: "executeRotation", sheetName: spread.getActiveSheet().name(),degree:90});
        }
    }

    designerConfig.contextMenu.push("textRotation");




}
