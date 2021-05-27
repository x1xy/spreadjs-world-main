var pictureTemplate = {
    "templateName":"pictureOptionTemplate",
    "content":[
        {"type":"TextBlock","style":"margin:10px;font-size: 20px;font-weight: lighter;color: #08892c","text":"图片属性"},

        {
            "type": "Container",
            "children": [
                {
                    "type": "FlexContainer",
                    "children": [
                        {
                            "type": "Container",
                            "visibleWhen": "series.selectedValue=1",
                            "children": [
                                {
                                    "type": "Container",
                                    "visibleWhen": "",
                                    "children": [
                                        {
                                            "type": "TabControl",
                                            "id": "c_tabControl",
                                            "className": "tab-control",
                                            "children": [
                                                {
                                                    "key": "fillAndLine",
                                                    "tip": "填充&线条",
                                                    "iconClass": "FillLine",
                                                    "selectedClass": "FillLine-selected",
                                                    "children": [
                                                        {
                                                            "type": "CollapsePanel",
                                                            "children": [
                                                                {
                                                                    "key": "fill",
                                                                    "text": "填充",
                                                                    "active": true,
                                                                    "children": [
                                                                        {
                                                                            "type": "Container",
                                                                            "margin": "10px 5px",
                                                                            "children": [
                                                                                {
                                                                                    "type": "ColumnSet",
                                                                                    "margin": "5px 0px",
                                                                                    "children": [
                                                                                        {
                                                                                            "type": "Column",
                                                                                            "children": [
                                                                                                {
                                                                                                    "type": "Radio",
                                                                                                    "bindingPath": "pictureFill",
                                                                                                    "items": [
                                                                                                        {
                                                                                                            "text": "无填充",
                                                                                                            "value": 0
                                                                                                        },
                                                                                                        {
                                                                                                            "text": "纯色填充",
                                                                                                            "value": 1
                                                                                                        }
                                                                                                    ]
                                                                                                }
                                                                                            ]
                                                                                        }
                                                                                    ]
                                                                                },
                                                                                {
                                                                                    "type": "ColumnSet",
                                                                                    "visibleWhen": "pictureFill=1",
                                                                                    "margin": "5px 0px",
                                                                                    "children": [
                                                                                        {
                                                                                            "type": "Column",
                                                                                            "width": "110px",
                                                                                            "children": [
                                                                                                {
                                                                                                    "type": "TextBlock",
                                                                                                    "text": "背景色"
                                                                                                }
                                                                                            ]
                                                                                        },
                                                                                        {
                                                                                            "type": "Column",
                                                                                            "width": "150px",
                                                                                            "children": [
                                                                                                {
                                                                                                    "type": "ColorComboEditor",
                                                                                                    "bindingPath": "backColor"
                                                                                                }
                                                                                            ]
                                                                                        }
                                                                                    ]
                                                                                }
                                                                            ]
                                                                        }
                                                                    ]
                                                                },
                                                                {
                                                                    "key": "line",
                                                                    "active": true,
                                                                    "text": "线条",
                                                                    "children": [
                                                                        {
                                                                            "type": "Container",
                                                                            "margin": "10px 5px",
                                                                            "children": [
                                                                                {
                                                                                    "type": "ColumnSet",
                                                                                    "margin": "5px 0px",
                                                                                    "children": [
                                                                                        {
                                                                                            "type": "Column",
                                                                                            "children": [
                                                                                                {
                                                                                                    "type": "Radio",
                                                                                                    "bindingPath": "pictureLine",
                                                                                                    "items": [
                                                                                                        {
                                                                                                            "text": "无线条",
                                                                                                            "value": 0
                                                                                                        },
                                                                                                        {
                                                                                                            "text": "实线",
                                                                                                            "value": 1
                                                                                                        }
                                                                                                    ]
                                                                                                }
                                                                                            ]
                                                                                        }
                                                                                    ]
                                                                                },
                                                                                {
                                                                                    "type": "ColumnSet",
                                                                                    "visibleWhen": "pictureLine=1",
                                                                                    "margin": "5px 0px",
                                                                                    "children": [
                                                                                        {
                                                                                            "type": "Column",
                                                                                            "width": "110px",
                                                                                            "children": [
                                                                                                {
                                                                                                    "type": "TextBlock",
                                                                                                    "text": "颜色"
                                                                                                }
                                                                                            ]
                                                                                        },
                                                                                        {
                                                                                            "type": "Column",
                                                                                            "width": "150px",
                                                                                            "children": [
                                                                                                {
                                                                                                    "type": "ColorComboEditor",
                                                                                                    "bindingPath": "borderColor"
                                                                                                }
                                                                                            ]
                                                                                        }
                                                                                    ]
                                                                                },
                                                                                {
                                                                                    "type": "ColumnSet",
                                                                                    "visibleWhen": "pictureLine=1",
                                                                                    "margin": "5px 0px",
                                                                                    "children": [
                                                                                        {
                                                                                            "type": "Column",
                                                                                            "width": "110px",
                                                                                            "children": [
                                                                                                {
                                                                                                    "type": "TextBlock",
                                                                                                    "text": "宽度"
                                                                                                }
                                                                                            ]
                                                                                        },
                                                                                        {
                                                                                            "type": "Column",
                                                                                            "width": "150px",
                                                                                            "children": [
                                                                                                {
                                                                                                    "type": "NumberEditor",
                                                                                                    "bindingPath": "borderWidth",
                                                                                                    "min": 1
                                                                                                }
                                                                                            ]
                                                                                        }
                                                                                    ]
                                                                                }
                                                                            ]
                                                                        }
                                                                    ]
                                                                }
                                                            ]
                                                        }
                                                    ]
                                                },
                                                {
                                                    "key": "size",
                                                    "tip": "大小",
                                                    "iconClass": "Size",
                                                    "selectedClass": "Size-selected",
                                                    "children": [
                                                        {
                                                            "type": "CollapsePanel",
                                                            "children": [
                                                                {
                                                                    "key": "size",
                                                                    "text": "大小",
                                                                    "active": true,
                                                                    "children": [
                                                                        {
                                                                            "type": "Container",
                                                                            "margin": "10px 5px",
                                                                            "children": [
                                                                                {
                                                                                    "type": "ColumnSet",
                                                                                    "margin": "5px 0px",
                                                                                    "children": [
                                                                                        {
                                                                                            "type": "Column",
                                                                                            "width": "110px",
                                                                                            "children": [
                                                                                                {
                                                                                                    "type": "TextBlock",
                                                                                                    "text": "高度"
                                                                                                }
                                                                                            ]
                                                                                        },
                                                                                        {
                                                                                            "type": "Column",
                                                                                            "width": "150px",
                                                                                            "children": [
                                                                                                {
                                                                                                    "type": "NumberEditor",
                                                                                                    "bindingPath": "size.height",
                                                                                                    "ruleType": "Defaults"
                                                                                                }
                                                                                            ]
                                                                                        }
                                                                                    ]
                                                                                },
                                                                                {
                                                                                    "type": "ColumnSet",
                                                                                    "margin": "5px 0px",
                                                                                    "children": [
                                                                                        {
                                                                                            "type": "Column",
                                                                                            "width": "110px",
                                                                                            "children": [
                                                                                                {
                                                                                                    "type": "TextBlock",
                                                                                                    "text": "宽度"
                                                                                                }
                                                                                            ]
                                                                                        },
                                                                                        {
                                                                                            "type": "Column",
                                                                                            "width": "150px",
                                                                                            "children": [
                                                                                                {
                                                                                                    "type": "NumberEditor",
                                                                                                    "bindingPath": "size.width",
                                                                                                    "ruleType": "Defaults"
                                                                                                }
                                                                                            ]
                                                                                        }
                                                                                    ]
                                                                                }
                                                                            ]
                                                                        }
                                                                    ]
                                                                },
                                                                {
                                                                    "key": "properties",
                                                                    "text": "属性",
                                                                    "active": true,
                                                                    "children": [
                                                                        {
                                                                            "type": "Container",
                                                                            "margin": "10px 5px",
                                                                            "children": [
                                                                                {
                                                                                    "type": "ColumnSet",
                                                                                    "margin": "5px 0px",
                                                                                    "children": [
                                                                                        {
                                                                                            "type": "Column",
                                                                                            "children": [
                                                                                                {
                                                                                                    "type": "Radio",
                                                                                                    "bindingPath": "properties.moveSizeRelationship",
                                                                                                    "items": [
                                                                                                        {
                                                                                                            "text": "随单元格改变位置和大小",
                                                                                                            "value": 0
                                                                                                        },
                                                                                                        {
                                                                                                            "text": "随单元格改变位置，但不改变大小",
                                                                                                            "value": 1
                                                                                                        },
                                                                                                        {
                                                                                                            "text": "不随单元格改变位置和大小",
                                                                                                            "value": 2
                                                                                                        }
                                                                                                    ]
                                                                                                }
                                                                                            ]
                                                                                        }
                                                                                    ]
                                                                                },
                                                                                {
                                                                                    "type": "ColumnSet",
                                                                                    "margin": "5px 0px",
                                                                                    "children": [
                                                                                        {
                                                                                            "type": "Column",
                                                                                            "width": "110px",
                                                                                            "children": [
                                                                                                {
                                                                                                    "type": "CheckBox",
                                                                                                    "text": "锁定",
                                                                                                    "bindingPath": "properties.locked"
                                                                                                }
                                                                                            ]
                                                                                        }
                                                                                    ]
                                                                                },
                                                                                {
                                                                                    "type": "ColumnSet",
                                                                                    "margin": "5px 0px",
                                                                                    "children": [
                                                                                        {
                                                                                            "type": "Column",
                                                                                            "width": "150px",
                                                                                            "children": [
                                                                                                {
                                                                                                    "type": "CheckBox",
                                                                                                    "text": "允许调整大小",
                                                                                                    "bindingPath": "properties.allowResize"
                                                                                                }
                                                                                            ]
                                                                                        }
                                                                                    ]
                                                                                },
                                                                                {
                                                                                    "type": "ColumnSet",
                                                                                    "margin": "5px 0px",
                                                                                    "children": [
                                                                                        {
                                                                                            "type": "Column",
                                                                                            "width": "150px",
                                                                                            "children": [
                                                                                                {
                                                                                                    "type": "CheckBox",
                                                                                                    "text": "允许移动",
                                                                                                    "bindingPath": "properties.allowMove"
                                                                                                }
                                                                                            ]
                                                                                        }
                                                                                    ]
                                                                                }
                                                                            ]
                                                                        }
                                                                    ]
                                                                }
                                                            ]
                                                        }
                                                    ]
                                                }
                                            ]
                                        }
                                    ]
                                }
                            ]
                        }
                    ]
                }
            ]
        }

    ]
}

GC.Spread.Sheets.Designer.registerTemplate("pictureOptionTemplate",pictureTemplate)

function getSelectedPicture (sheet) {
    let allPictures = sheet.pictures.all();
    for (let i = 0; i < allPictures.length; i++) {
        let picture = allPictures[i];
        if (picture.isSelected()) {
            return picture;
        }
    }
}

var sidePanelsPictureCommands = {pictureOptionPanel: {
    commandName: "pictureOptionPanel",
    enableContext: "AllowEditObject",
    visibleContext: "pictureSelected && !pictureAltTextPanel_Visible",
    execute: function(context, propertyName, newValue){
        if(!propertyName){
            return;
        }
        let sheet = context.Spread.getActiveSheet();
        let activePicture = getSelectedPicture(sheet);
        switch(propertyName){
            case "pictureFill":
                if(newValue === 0){
                    activePicture.backColor(null)
                }
                else{
                    activePicture.backColor("white")
                }
                break;
            case "pictureLine":
                if(newValue === 0){
                    activePicture.borderStyle("none")
                }
                else{
                    activePicture.borderStyle("double")
                }
                break;
            default:
                Reflect.apply(activePicture[propertyName], activePicture, [newValue])
                break;
        }
    },
    getState: function(context){
        let sheet = context.Spread.getActiveSheet();
        let activePicture = getSelectedPicture(sheet);
        const pictureStatus = {
            pictureFill: 0,
            backColor: "",
            pictureLine: 0,
            borderWidth: 0,
            borderColor: ""
        };
        if (activePicture) {

            for(let key in pictureStatus) {
                switch(key){
                    case "backColor":
                        let backColor = activePicture.backColor();
                        if(backColor){
                            pictureStatus.backColor = backColor;
                            pictureStatus.pictureFill = 1;
                        }else {
                            pictureStatus.backColor = "#000000";
                            pictureStatus.pictureFill = 0;
                        }
                        break;
                    case "borderWidth":
                        pictureStatus.borderWidth = activePicture.borderWidth();
                        break;
                    case "borderColor":
                        pictureStatus.borderColor = activePicture.borderColor();
                        break;
                }
            }

            let borderStyle = activePicture.borderStyle();
            if(borderStyle && borderStyle !== "none"){
                pictureStatus.pictureLine = 1
            }else {
                pictureStatus.pictureLine = 0
            }
        }
        return pictureStatus;
    },
}
}


var sidePanelsPictureConfig =
    {
        "position": "right",
        "width": "315px",
        "command": "pictureOptionPanel",
        "uiTemplate": "pictureOptionTemplate",
        "showCloseButton": true
    }

