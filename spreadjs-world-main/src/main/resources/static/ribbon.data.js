

async function sjsBinding(designer) {

    function Company(name, logo, slogan, address, city, phone, email) {
        this.name = name;
        this.logo = logo;
        this.slogan = slogan;
        this.address = address;
        this.city = city;
        this.phone = phone;
        this.email = email;
    }
    function Customer(id, name, company) {
        this.id = id;
        this.name = name;
        this.company = company;
    }
    function Record(description, quantity, amount) {
        this.description = description;
        this.quantity = quantity;
        this.amount = amount;
    }
    function Invoice(company, number, date, customer, receiverCustomer, records) {
        this.company = company;
        this.number = number;
        this.date = date;
        this.customer = customer;
        this.receiverCustomer = receiverCustomer;
        this.records = records;
    }
    var company1 = new Company("GrapeCity", null, "We know everything!", "Gaoxin 6th road", "Xi'an", "029-88331988", "grapecity@grapecity.com"),
        company2 = new Company("Tecent", null, "We have everything!", "Shenzhen 2st road", "Shenzhen", "0755-12345678", "tecent@qq.com"),
        company3 = new Company("Alibaba", null, "We sale everything!", "Hangzhou 3rd road", "Hangzhou", "0571-12345678", "alibaba@alibaba.com"),
        customer1 = new Customer("A1", "employee 1", company2),
        customer2 = new Customer("A2", "employee 2", company3),
        records1 = [],
        invoice1 = new Invoice(company1, "00001", new Date(2014, 0, 1), customer1, customer2, records1);

    for (var i = 0; i < Math.random() * 20; i++) {
        records1.push(new Record("Description" + i, i, Math.round(Math.random() * 150)))
    }

    dataSource = new GC.Spread.Sheets.Bindings.CellBindingSource(invoice1);

    var sheet = designer.getWorkbook().getActiveSheet();

    sheet.setDataSource(dataSource);
}

async function sjsGCBinding(designer) {
    var spread = designer.getWorkbook();
    var jsonString = JSON.stringify(spread.toJSON())
    var uploadData = gzipString(jsonString);
    $.post("spread/bindingSJSData", { data: uploadData }, function (data) {
        if (data != "0") {

            data = ungzipString(data);
            var json = JSON.parse(data);
            spread.fromJSON(json);
        }
    })
}

var spreadTemplateJSON;
var AllowBindingData = "AllowBindingData", AllowBackTemplate = "AllowBackTemplate";
async function bindingData(designer) {
    var spread = designer.getWorkbook();
    if (!spreadTemplateJSON) {
        spreadTemplateJSON = JSON.stringify(spread.toJSON())
    }

    console.log(JSON.parse(spreadTemplateJSON))
    var uploadData = gzipString(spreadTemplateJSON);
    $.post("spread/bindingData", { data: uploadData }, function (data) {
        if (data != "0") {

            designer.setData(AllowBackTemplate, true)
            data = ungzipString(data);
            var json = JSON.parse(data);
            //            json.calcOnDemand = true;
            //            console.log(json)
            //            spread.fromJSON(json, { doNotRecalculateAfterLoad: true });
            spread.fromJSON(json);
        }
        else {
            spreadTemplateJSON = undefined;
            designer.setData(AllowBackTemplate, false)
        }
    })
}

async function backTemplate(designer) {
    if (spreadTemplateJSON) {
        var spread = designer.getWorkbook();
        spread.fromJSON(JSON.parse(spreadTemplateJSON))
        spreadTemplateJSON = undefined;
    }
    designer.setData(AllowBindingData, true)
    designer.setData(AllowBackTemplate, false)
}






var ribbonDataCommands = {

    "sjsBinding": {
        iconClass: "ribbon-button-namemanager",
        text: "前端 bindingPath 绑定",
        bigButton: true,
        commandName: "sjsBinding",
        execute: sjsBinding
    },
    "sjsGCBinding": {
        iconClass: "ribbon-button-namemanager",
        text: "GCExcel bindingPath 绑定",
        bigButton: true,
        commandName: "sjsGCBinding",
        execute: sjsGCBinding
    },
    "bindingData": {
        iconClass: "ribbon-button-namemanager",
        text: "GCExcel模板绑定",
        commandName: "bindingData",
        execute: bindingData,
        enableContext: AllowBindingData
    },
    "backTemplate": {
        iconClass: "ribbon-button-namemanager",
        text: "返回模板",
        commandName: "backTemplate",
        execute: backTemplate,
        enableContext: "AllowBackTemplate"
    }

}




var ribbonDataConfig = {
    "label": "数据绑定",
    "thumbnailClass": "ribbon-thumbnail-spreadsettings",
    "commandGroup": {
        "children": [
            {
                "direction": "vertical",
                "commands": ["sjsBinding"]
            },
            {
                "type": "separator"
            },
            {
                "direction": "vertical",
                "commands": ["sjsGCBinding"]
            },
            {
                "type": "separator"
            },
            {
                "direction": "vertical",
                "commands": [
                    "bindingData",
                    "backTemplate"
                ]
            }
        ]
    }
};