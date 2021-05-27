
var ribbonTabCommands = {

    "newTab": {
        text: "页面设置",
        iconClass: "ribbon-button-sheetgeneral",
        bigButton: true,
        commandName: "newTab",
        execute: async (context) => {


            var spread = context.getWorkbook();
            var sheet = spread.getActiveSheet();
            var sheetIndex = spread.getSheetIndex(sheet.name())

             var option = {
                tag: sheet.tag(),
                type: "rowColId",
                row: true,
//                colClassName: false,
                dialogOption: {
                    activeTab: "Test2",
                    showTabList: ["Test2"],
                },
            }
            GC.Spread.Sheets.Designer.showDialog("newTab", option, (result) => {
                var printInfo = sheet.printInfo();
                console.log(result)
                console.log(result.type)
                if(result.type === "uuid"){
                    printInfo.orientation(GC.Spread.Sheets.Print.PrintPageOrientation.landscape);


                     if(result.row === true){
                           printInfo.showBorder(true);
                           printInfo.showGridLine(false);
                     }

                     if(result.col === true){

                           if(result.row === false){
                                printInfo.showBorder(false);
                           }
                          printInfo.showGridLine(true);
                     }
                     spread.print();
                }



                console.log(result)
                if(result.cellId === false){

                    printInfo.showRowHeader(GC.Spread.Sheets.Print.PrintVisibilityType.hide);

                    if(result.colClassName === false){
                        printInfo.showColumnHeader(GC.Spread.Sheets.Print.PrintVisibilityType.hide);
                    }

                     if(result.colClassName === true){
                        printInfo.showColumnHeader(GC.Spread.Sheets.Print.PrintVisibilityType.show);
                     }

                    spread.print();
                }

                if(result.cellId === true){

                     printInfo.showRowHeader(GC.Spread.Sheets.Print.PrintVisibilityType.show);

                     if(result.colClassName === true){
                        printInfo.showColumnHeader(GC.Spread.Sheets.Print.PrintVisibilityType.show);
                     }
                     if(result.colClassName === false){
                        printInfo.showColumnHeader(GC.Spread.Sheets.Print.PrintVisibilityType.hide);
                     }

                     spread.print();
                }

                spread.commandManager().execute({
                    cmd: "Designer.setTag",
                    type: "sheet",
                    tag: result.tag,
                    sheetName: sheet.name()
                })
            });

        }
    }
}



var CreateId = {
  type: "FlexContainer",
  margin: "10px",
  children: [
    {
      type: "FlexContainer",
      children: [
        {
          type: "Radio",
          items: [
            {
              text: '横向打印',
              value: 'uuid'
            },
            {
              text: '纵向打印',
              value: 'rowColId',
            }],
          bindingPath: "type"
        },

        {
          type: 'FlexContainer',
          enableWhen: "type=uuid",
          margin: '10px 0',
          children: [
            {
              type: 'ColumnSet',
              children: [
                {
                  type: 'Column',
                  children: [
                    {
                      type: "CheckBox",
                      text: '显示边框',
                      bindingPath: 'row'
                    }
                  ],
                  width: '120px'
                },
                {
                  type: 'Column',
                  children: [
                    {
                      type: "CheckBox",
                      text: '显示网格线',
                      bindingPath: 'col'
                    }
                  ],
                  width: '120px'
                }
              ]
            }
          ]
        },
        {
          type: 'FlexContainer',
          enableWhen: "type=rowColId",
          margin: '10px 0',
          children: [
            {
              type: 'ColumnSet',
              children: [
                {
                  type: 'Column',
                  children: [
                    {
                      type: "CheckBox",
                      text: '显示行头',
//                      text: '显示行头(生成单元格ID)',
                      bindingPath: 'cellId'
                    }
                  ],
                  width: '120px'
                },
                {
                  type: 'Column',
                  children: [
                    {
                      type: "CheckBox",
                      text: '显示列头',
//                      text: '显示列头(用列ID生成className)',
                      bindingPath: 'colClassName'
                    }
                  ],
                  width: '150px'
                }
              ]
            }
          ]
        },
            {
              type: "Column",
              children: [
                {
                  type: "ColumnSet",
                  children: [
                    {
                      type: "Column",
                      children: [

                        {
                          type: "TextBlock",
                          text: "宽度",
                          margin: "0px 0 0 30px",
                        }
                      ],
                      width: "120px"
                    },
                    {
                      type: "Column",
                      children: [
                        {
                          type: "TextEditor",
                          bindingPath: "width",
                        }
                      ],
                      width: "120px"
                    },
                    {
                          type: "Column",
                          children: [

                                {
                                   type: "TextBlock",
                                   text: "高度",
                                   margin: "0px 0 0 30px",
                                }
                           ],
                             width: "120px"
                          },
                       {
                          type: "Column",
                          children: [
                            {
                               type: "TextEditor",
                               bindingPath: "height",
                            }
                          ],
                          width: "120px"
                        }
                  ]
                }
              ]
            }
      ]
    },

  ]
}

var richSheetTagTemplate = {
    title: "自定义打印",
      content: [
        {
          type: "TabControl",
          width: 670,
          height: 530,
          bindingPath: "dialogOption",
          children: [
            {
              key: "Test2",
              text: "打印",
              children: [
                CreateId
              ]
            },
          ]
        }
      ]
}
GC.Spread.Sheets.Designer.registerTemplate("newTab", richSheetTagTemplate);



var ribbonPivotConfig1 = {
    "label": "打印",
    "thumbnailClass": "ribbon-thumbnail-spreadsettings",
    "commandGroup": {
        "children": [
            {
                "direction": "vertical",
                "commands": ["newTab"]
            }
        ]
    }
}