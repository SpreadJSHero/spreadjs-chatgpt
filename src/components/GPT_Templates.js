import * as GC from "@grapecity/spread-sheets";


import "@grapecity/spread-sheets-designer-resources-cn"
import "@grapecity/spread-sheets-designer"
import {openai} from './openai'

import { ElLoading } from "element-plus";

let globalDesigner;

let formulaGenerateTemplate = {
    "templateName": "formulaGenerateTemplate",
    "title": "智能公式生成",
    "content": [
        {
            "type": "FlexContainer",
            "children": [
                { "type": "TextBlock", "text": "描述:", "style": "line-height:25px" },
                { "type": "TextEditor", "multiLine": true, "resize": false, "bindingPath": "description", "style": "width: 300px;height:60px","placeholder": "请输入公式描述，例如：对B列求和，当C列的时间大于2023年1月1日并且小于2023年3月31日。" },
                { "type": "TextBlock", "text": "公式:", "style": "line-height:25px" },
                { "type": "TextBlock", "bindingPath": "formula", "style": "width: 300px; table-layout:fixed; word-break:break-all;", "enableWhen": "never=true"}]
        }],
    "buttons": [
        { "text": "生成", "buttonType": "Ok", "closeAfterClick": true },
        { "text": "插入", "visibleWhen": "canInsert=true", "bindingPath": "canInsert", "closeAfterClick": true, "click": function(data){
            if(!globalDesigner){
                GC.Spread.Sheets.Designer.showMessageBox("没有注册Designer对象", "错误", GC.Spread.Sheets.Designer.MessageBoxIcon.error)
            }
            let spread = globalDesigner.getWorkbook(), sheet = spread.getActiveSheet();
            spread.commandManager().execute({cmd: "editCell", row: sheet.getActiveRowIndex(),col: sheet.getActiveColumnIndex(), newValue: data.formula, sheetName: sheet.name() });
            let formulaTextBox = globalDesigner.getData("formulaBar");
            if(formulaTextBox){
                formulaTextBox.refresh(true);
            }
            // 用完销毁防止内存泄露
            globalDesigner = undefined;
        }},
        { "text": "取消", "buttonType": "Cancel" }]
}
GC.Spread.Sheets.Designer.registerTemplate("formulaGenerateTemplate", formulaGenerateTemplate);

var dataAnalyzeTemplate = {
    "templateName": "dataAnalyzeTemplate",
    "title": "选择区域数据智能分析",
    "content": [
        {
            "type": "FlexContainer",
            "children": [
                { "type": "TextBlock", "text": "描述:", "style": "line-height:25px" },
                { "type": "TextEditor", "multiLine": true, "resize": false, "bindingPath": "description", "style": "width: 300px;height:60px","placeholder": "请输入分析描述，例如谁的销售情况最好？" },
                { "type": "TextBlock", "text": "结果:", "style": "line-height:25px" },
                { "type": "TextBlock", "bindingPath": "result", "multiLine": true, "resize": false, "style": "width: 300px; table-layout:fixed; word-break:break-all;", "enableWhen": "never=true"}]
        }],
    "buttons": [
        { "text": "分析", "buttonType": "Ok", "closeAfterClick": true },
        { "text": "关闭", "buttonType": "Cancel" }]
}
GC.Spread.Sheets.Designer.registerTemplate("dataAnalyzeTemplate", dataAnalyzeTemplate);

function checkDialogDescription(value) {
    if (!value.description) {
        GC.Spread.Sheets.Designer.showMessageBox("请输入描述！", "Warning", GC.Spread.Sheets.Designer.MessageBoxIcon.warning);
        return false;
    } else {
        return true
    }
}

function showFormulaGenerateDialog(data, designer){
    if(!designer){
        GC.Spread.Sheets.Designer.showMessageBox("没有注册Designer对象", "错误", GC.Spread.Sheets.Designer.MessageBoxIcon.error)
    }
    globalDesigner = designer;
    GC.Spread.Sheets.Designer.showDialog("formulaGenerateTemplate", data, function(bindingData){
        let description = bindingData.description;
        if(description){
            let loading = ElLoading.service({ lock: true, text: "Loading", background: "rgba(0, 0, 0, 0.7)"});
            const response = openai.createCompletion({
                model: "text-davinci-003",
                prompt: "帮我写一个Excel公式， " + description,
                max_tokens: 100,
                temperature: 0.5
            });
            response.then(function(completion){
                loading.close();
                let formula = completion.data.choices[0].text.trim();
                bindingData.formula = formula;
                if(formula.indexOf("=") === 0){
                    bindingData.canInsert = true;
                }
                showFormulaGenerateDialog(bindingData, designer);
            }).catch(function(error){
                loading.close();
            });
        }
    }, undefined, checkDialogDescription);
}


function showDataAnalyzeDialog(data, designer){
    if(!designer){
        GC.Spread.Sheets.Designer.showMessageBox("没有注册Designer对象", "错误", GC.Spread.Sheets.Designer.MessageBoxIcon.error)
    }
    GC.Spread.Sheets.Designer.showDialog("dataAnalyzeTemplate", data, function(bindingData){
        let description = bindingData.description;
        if(description){
            let spread = designer.getWorkbook();
            let sheet = spread.getActiveSheet(), selection = sheet.getSelections()[0];
            let data = sheet.getCsv(selection.row, selection.col, selection.rowCount, selection.colCount, "\n", ",");
            let loading = ElLoading.service({ lock: true, text: "Loading", background: "rgba(0, 0, 0, 0.7)"});
            const response = openai.createCompletion({
                model: "text-davinci-003",
                prompt: "表格数据如下:\n" + data + "\n\n分析" + description,
                max_tokens: 100,
                temperature: 0.5
            });
            response.then(function(completion){
                loading.close();
                let result = completion.data.choices[0].text.trim();
                bindingData.result = result;
                showDataAnalyzeDialog(bindingData, designer);
            }).catch(function(error){
                loading.close();
            });
        }

    }, undefined, checkDialogDescription);
}


var createPivotTableTemplate = {
    "templateName": "createPivotTableTemplate",
    "title": "创建数据透视表",
    "content": [
        {
            "type": "FlexContainer",
            "children": [
                { "type": "TextBlock", "text": "数据区域:", "style": "line-height:25px" },
                { "type": "TextBlock", "bindingPath": "dataRange", "multiLine": true, "resize": false, "style": "width: 300px; table-layout:fixed; word-break:break-all;", "enableWhen": "never=true"},
                { "type": "TextBlock", "text": "描述:", "style": "line-height:25px" },
                { "type": "TextEditor", "multiLine": true, "resize": false, "bindingPath": "description", "style": "width: 300px;height:80px","placeholder": "请输入分析描述，例如销售人员各品牌的销售数量,\n无描述直接创建建议透视表" },
            ]
        }],
    "buttons": [
        { "text": "创建", "buttonType": "Ok", "closeAfterClick": true },
        { "text": "关闭", "buttonType": "Cancel" }]
}
GC.Spread.Sheets.Designer.registerTemplate("createPivotTableTemplate", createPivotTableTemplate);


function showPivotTableCreateDialog(data, designer){
    if(!designer){
        GC.Spread.Sheets.Designer.showMessageBox("没有注册Designer对象", "错误", GC.Spread.Sheets.Designer.MessageBoxIcon.error)
    }
    let spread = designer.getWorkbook(), sheet = spread.getActiveSheet();
    let table = sheet.tables.find(sheet.getActiveRowIndex(), sheet.getActiveColumnIndex());
    let pivotRange, headerList;
    if(table){
        let tableRange = table.range();
        headerList = sheet.getArray(tableRange.row, tableRange.col, 1, tableRange.colCount)[0].join(",");
        pivotRange = table.name();
    }
    else{
        let selection = sheet.getSelections()[0];
        headerList = sheet.getArray(selection.row, selection.col, 1, selection.colCount)[0].join(",");
        pivotRange = "'" + sheet.name() +  "'!" + GC.Spread.Sheets.CalcEngine.rangeToFormula(selection, 0, 0);
    }
    data.dataRange = pivotRange;
    GC.Spread.Sheets.Designer.showDialog("createPivotTableTemplate", data, async function(bindingData){
        debugger
        if(!bindingData){
            return;
        }
        let loading = ElLoading.service({ lock: true, text: "Loading", background: "rgba(0, 0, 0, 0.7)"});
     
        let messages = [
            {"role": "user", "content": "表格包含" + headerList + "这些列，创建数据透视表，返回行，列，值字段，用来分析" + bindingData.description}
        ];
        let functions = [{
            "name": "pivot_talbe_analyze",
            "description": "对数据创建数据透视表，返回数据透视表结果",
            "parameters": {
                "type": "object",
                "properties": {
                    "rowFieldName": {
                        "type": "string",
                        "description": "行字段名称"
                    },
                    "columnFieldName": {
                        "type": "string",
                        "description": "列段名称"
                    },
                    "dataFieldName": {
                        "type": "string",
                        "description": "值字段名称"
                    },
                },
            },
            "required": ["rowFieldName", "dataFieldName"]
        }]
        try {
            var completion = await openai.createChatCompletion({
                "model": "gpt-3.5-turbo-0613",
                "messages": messages,
                "functions": functions,
                "function_call": {"name": "pivot_talbe_analyze"}
            });
            console.log(completion.data.choices[0].message.function_call); 
            if(completion.data.choices[0].message.function_call){
                let args = JSON.parse(completion.data.choices[0].message.function_call.arguments);
                spread.suspendPaint();
                let activeSheetIndex = spread.getActiveSheetIndex();
                spread.addSheet(activeSheetIndex);
                spread.setActiveSheetIndex(activeSheetIndex);
                let newSheet = spread.getSheet(activeSheetIndex);
                let pivotTable = newSheet.pivotTables.add(getUniquePivotName(newSheet), pivotRange, 2, 0, GC.Spread.Pivot.PivotTableLayoutType.outline, GC.Spread.Pivot.PivotTableThemes.medium2);
                pivotTable.add(args.rowFieldName, args.rowFieldName, GC.Spread.Pivot.PivotTableFieldType.rowField);
                if(args.columnFieldName){
                    pivotTable.add(args.columnFieldName, args.columnFieldName, GC.Spread.Pivot.PivotTableFieldType.columnField);
                }
                pivotTable.add(args.dataFieldName, "求和项：" + args.dataFieldName, GC.Spread.Pivot.PivotTableFieldType.valueField, GC.Pivot.SubtotalType.sum);

                spread.resumePaint();
            }
        }
        catch(err){
            console.log(err)
        }
        finally{
            globalDesigner = undefined;
            loading.close();
        }
    }, function(error){
        loading.close();
    });
}

function getUniquePivotName(sheet) {
    debugger
    var count = 1;
    if (!sheet.pivotTables) {
        return;
    }
    var sheets = sheet.getParent().sheets;
    var length = sheets.length;
    for (var index = 0; index < length; index++) {
        while (sheets[index].pivotTables.get("PivotTable" + count)) {
            count++;
            index = 0;
        }
    }
    return "PivotTable" + count;
}

export {showFormulaGenerateDialog, showDataAnalyzeDialog, showPivotTableCreateDialog}