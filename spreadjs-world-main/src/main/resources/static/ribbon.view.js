var ribbonViewCommands = {
    "showHideScrollbar": {
        text: "显示滚动条",
        type: "checkbox",
        iconClass: "ribbon-button-namemanager",
        bigButton: true,
        commandName: "showHideScrollbar",
        execute: async (context) => {
            setTimeout(() => {
                let vsb = context.Spread.options.showVerticalScrollbar;
                let hsb = context.Spread.options.showHorizontalScrollbar;
                context.Spread.options.showVerticalScrollbar = !vsb;
                context.Spread.options.showHorizontalScrollbar  = !hsb;
            }, 0)
        },
        getState: (context) => {
            let sheet = context.Spread.getActiveSheet();
            return context.Spread.options.showVerticalScrollbar;
        }
    }
}

