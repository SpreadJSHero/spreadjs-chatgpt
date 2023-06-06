import * as GC from "@grapecity/spread-sheets";


import "@grapecity/spread-sheets-designer-resources-cn"
import "@grapecity/spread-sheets-designer"
import {openai} from './openai'

import { ElLoading } from "element-plus";

let globalSpread;

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
            if(!globalSpread){
                GC.Spread.Sheets.Designer.showMessageBox("没有注册Spread对象", "错误", GC.Spread.Sheets.Designer.MessageBoxIcon.error)
            }
            let sheet = globalSpread.getActiveSheet();
            globalSpread.commandManager().execute({cmd: "editCell", row: sheet.getActiveRowIndex(),col: sheet.getActiveColumnIndex(), newValue: data.formula, sheetName: sheet.name() })
            // 用完销毁防止内存泄露
            globalSpread = undefined;
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
        { "text": "生成", "buttonType": "Ok", "closeAfterClick": true },
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

function showFormulaGenerateDialog(data, spread){
    globalSpread = spread;
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
                showFormulaGenerateDialog(bindingData, spread);
            }).catch(function(error){
                loading.close();
            });
        }
    }, undefined, checkDialogDescription);
}


function showDataAnalyzeDialog(data, spread){
    GC.Spread.Sheets.Designer.showDialog("dataAnalyzeTemplate", data, function(bindingData){
        let description = bindingData.description;
        if(description){
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
                showDataAnalyzeDialog(bindingData, spread);
            }).catch(function(error){
                loading.close();
            });
        }

    }, undefined, checkDialogDescription);
}

export {showFormulaGenerateDialog, showDataAnalyzeDialog}