
async function getTemplates(designer) {
    var spread = designer.getWorkbook();

    $.get("spread/getTemplates", function (data) {
        if (data != "0") {
            data = ungzipString(data);
            var files = data.split(";")
            var sheet = spread.getActiveSheet();
            sheet.reset();
            sheet.defaults.colWidth = 180
            var fileList = [];
            for (var i = 0; i < files.length; i++) {
                fileList.push([files[i]])
            }
            sheet.setArray(0, 0, fileList)
            var h = new GC.Spread.Sheets.CellTypes.HyperLink();
            h.onClickAction(function (e) {
                var sheet = e.sheet, row = e.row, col = e.col;
                var value = sheet.getValue(row, col);
                loadTemplate(designer, value)
            })
            sheet.getRange(0, 0, files.length, 1).cellType(h)
        }
    })
}
async function loadTemplate(designer, fileName) {
    var spread = designer.getWorkbook();

    let formData = new FormData();
    formData.append("fileName", fileName);

    $.ajax({
        url: "spread/loadTemplate",
        type: "POST",
        contentType: false,
        processData: false,
        data: formData,
        success: function (data) {
            if (data != "0") {
                console.time("ungzip")
                data = ungzipString(data);
                console.timeEnd("ungzip")

                console.time("parse")
                var json = JSON.parse(data);
                console.timeEnd("parse")
                console.time("fromJSON")

                if (json.designerBindingPathSchema) {
                    spreaddesignerBindingPathSchema = json.designerBindingPathSchema;
                    designer.setData("treeNodeFromJson", JSON.stringify(json.designerBindingPathSchema));
                    designer.setData("oldTreeNodeFromJson", JSON.stringify(json.designerBindingPathSchema));
                    delete json.designerBindingPathSchema;
                }

                json.calcOnDemand = true;
                spread.fromJSON(json, { doNotRecalculateAfterLoad: true });
                console.timeEnd("fromJSON")
                initStates(designer)
            }
        }
    })
}
async function uploadFile(designer) {
    var spread = designer.getWorkbook();
    var ribbon_file_selector = document.getElementById("ribbon_file_selector");
    ribbon_file_selector.onchange = function (event) {
        console.log(event)
        var file = event.target.files[0];
        let formData = new FormData();
        formData.append("file", file);
        formData.append("fileName", file.name);

        $.ajax({
            url: "spread/uploadFile",
            type: "POST",
            contentType: false,
            processData: false,
            data: formData,
            success: function (data) {
                if (data != "上传失败！") {
                    data = ungzipString(data);
                    var json = JSON.parse(data);
                    json.calcOnDemand = true;
                    spread.fromJSON(json, { doNotRecalculateAfterLoad: true });

                    initStates(designer)

                }
                else {
                    alert("上传失败")
                }
            }
        })
    }
    ribbon_file_selector.value = "";
    ribbon_file_selector.click();

}

async function exportPDF(designer, bindingType, isActiveSheet) {
    console.log(arguments)
    var spread = designer.getWorkbook();
    var spreadJSON = JSON.stringify(spread.toJSON({ includeBindingSource: true }));
    var uploadData = gzipString(spreadJSON);

    var formData = new FormData();
    formData.append("data", uploadData)
    formData.append("bindingType", bindingType ? bindingType : 0)
    formData.append("isActiveSheet", isActiveSheet ? true : false)

    fetch("spread/getpdf", {
        body: formData,
        method: "POST"
    }).then(resp => resp.blob())
        .then(blob => {
            const url = window.URL.createObjectURL(blob);
            window.open(url, "_blank")
        })
}

async function exportImage(designer, isSelectedRange, isPDFMode) {
    var spread = designer.getWorkbook();
    var spreadJSON = JSON.stringify(spread.toJSON());
    var uploadData = gzipString(spreadJSON);

    var formData = new FormData();
    formData.append("data", uploadData)
    if (isSelectedRange) {
        formData.append("isSelectedRange", isSelectedRange)
    }
    if (isPDFMode) {
        formData.append("isPDFMode", isPDFMode)
    }

    fetch("spread/getImage", {
        body: formData,
        method: "POST"
    }).then(resp => resp.blob())
        .then(blob => {
            const url = window.URL.createObjectURL(blob);
            window.open(url, "_blank")
        })
}

async function getHTML(designer) {
    var spread = designer.getWorkbook();
    var spreadJSON = JSON.stringify(spread.toJSON());
    var uploadData = gzipString(spreadJSON);
    var formData = new FormData();
    formData.append("data", uploadData)

    fetch("spread/getHTML", {
        body: formData,
        method: "POST"
    }).then(resp => resp.blob())
        .then(blob => {
            const url = window.URL.createObjectURL(blob);
            window.open(url, "_blank")
        })
}

var ribbonFileCommands = {
    "getTemplates": {
        iconClass: "ribbon-button-download",
        text: "加载",
        commandName: "getTemplates",
        execute: getTemplates,
    },
    "uploadFile":
    {
        iconClass: "ribbon-button-upload",
        text: "上传",
        commandName: "uploadFile",
        execute: uploadFile,
    },
    "pdf0": {
         iconClass: "ribbon-button-pdf",
//        iconClass: "ribbon-button-pdf-font",
        bigButton: true,
        text: "PDF",
        commandName: "pdf0",
        execute: async (context, selectValue) => {
            await exportPDF(context, 0, false);
        }
    },
    "pdf1": {
        text: "导出当前Sheet",
        iconClass: "ribbon-button-pdf-small",
        commandName: "pdf1",
        execute: async (context, selectValue) => {
            await exportPDF(context, 0, true);
        }
    },
    "pdf2": {
        text: "绑定数据导出",
        iconClass: "ribbon-button-pdf-small",
        commandName: "pdf2",
        execute: async (context, selectValue) => {
            await exportPDF(context, 1, false);
        }
    },
    "pdf3": {
        text: "GCExcel模板绑定导出",
        iconClass: "ribbon-button-pdf-small",
        commandName: "pdf3",
        execute: async (context, selectValue) => {
            await exportPDF(context, 1, true);
        }
    },
    "pdf4": {
        text: "GCExcel BindingPath绑定导出",
        iconClass: "ribbon-button-pdf-small",
        commandName: "pdf4",
        execute: async (context, selectValue) => {
            await exportPDF(context, 2, true);
        }
    },
    "PDF_List": {
        title: "PDF_List",
        type: "dropdown",
        commandName: "pdf_sub",
        subCommands: [
            "pdf1",
        ]
    },
    "exportImage_List": {
        title: "exportImage_List",
        bigButton: true,
        iconClass: "ribbon-button-namemanager",
        text:"导出图片",
        type: "dropdown",
        commandName: "exportImage_List",
        subCommands: [
            "exportSheetPDFImage",
            "exportRangePDFImage",
            "exportSheetImage",
            "exportRangeImage"
        ]
    },
    "exportSheetPDFImage": {
        iconClass: "ribbon-button-namemanager",
        text: "Sheet导出PDF图片",
        commandName: "exportSheetPDFImage",
        execute: async function (context) {
        exportImage(context, false, true)
        }
    },
    "exportRangePDFImage": {
        iconClass: "ribbon-button-namemanager",
        text: "选择区域导出PDF图片",
        commandName: "exportRangePDFImage",
        execute: async function (context) {
        exportImage(context, true, true)
        }
    },
    "exportSheetImage": {
        iconClass: "ribbon-button-namemanager",
        text: "Sheet导出图片",
        commandName: "exportSheetImage",
        execute: async function (context) {
            exportImage(context, false, false)
        }
    },
    "exportRangeImage": {
        iconClass: "ribbon-button-namemanager",
        text: "选择区域导出图片",
        commandName: "exportRangeImage",
        execute: async function (context) {
            exportImage(context, true, false)
        }
    },
    "getHTML": {
        iconClass: "ribbon-button-namemanager",
        text: "HTML",
        bigButton: true,
        commandName: "getHTML",
        execute: getHTML
    }
}

var ribbonFileConfig =
{
    "label": "文件操作",
    "thumbnailClass": "ribbon-thumbnail-spreadsettings",
    "commandGroup": {
        "children": [
            {
                "direction": "vertical",
                "commands": ["getTemplates", "uploadFile"]
            },
            {
                "type": "separator"
            },
            {
                "commands": [
                    "pdf0",
                    "PDF_List"
                ]
            },
            {
                "type": "separator"
            },
            {
                "direction": "vertical",
                "commands": ["pdf4", "pdf3"]
            },
            {
                "type": "separator"
            },
//             {
//                 "direction": "vertical",
//                 "commands": [
//                     "exportSheetImage",
//                     "exportRangeImage"
//                 ]
//             },
             {
                 "direction": "vertical",
                 "commands": [
                     "exportImage_List"
                 ]
             },
            {
                "type": "separator"
            },
            {
                "direction": "vertical",
                "commands": [
                    "getHTML"
                ]
            }
        ]
    }
};