import * as GC from "@grapecity/spread-sheets";

import {openai} from './openai'

var GPT_Query = function () {};
GPT_Query.prototype = new GC.Spread.CalcEngine.Functions.AsyncFunction('GPT.QUERY', 1, 1,  {
        description: "向GPT提问，直接返回结果",
        parameters: [
            {
                name: "问题"
            }]});
GPT_Query.prototype.defaultValue = function () { return 'Loading...'; };
GPT_Query.prototype.evaluateAsync = function (context, arg) {
    if (!arg) {
        return GC.Spread.CalcEngine.Errors.NotAvailable;
    }

    const response = openai.createCompletion({
        model: "text-davinci-003",
        prompt: arg + "？只返回结果。",
        max_tokens: 100,
        temperature: 0.5
    });
    response.then(function(completion){
        context.setAsyncResult(completion.data.choices[0].text.trim());
    })
};
GC.Spread.CalcEngine.Functions.defineGlobalCustomFunction("GPT.QUERY", new GPT_Query());




var GPT_Filter = function () {};
GPT_Filter.prototype = new GC.Spread.CalcEngine.Functions.AsyncFunction('GPT.FILTER', 2, 2,  {
        description: "对选择的数据区域做描述行的过滤",
        parameters: [
            {
                name: "数据区域"
            },
            {
                name: "过滤条件描述"
            }]});
GPT_Filter.prototype.defaultValue = function () { return 'Loading...'; };
GPT_Filter.prototype.acceptsArray = function () {
    return true;
}
GPT_Filter.prototype.acceptsReference = function () {
    return true;
}
GPT_Filter.prototype.evaluateAsync = function (context, range, desc) {
    if (!range || !desc) {
        return GC.Spread.CalcEngine.Errors.NotAvailable;
    }
    
    let tempArray = range.toArray && range.toArray();
    if (!Array.isArray(tempArray)) {
        return GC.Spread.CalcEngine.Errors.NotAvailable;
    }

    const response = openai.createCompletion({
        model: "text-davinci-003",
        prompt: "对最后的JSON数据，过滤" + desc + "，返回和原来结构一样的双层JSON数组。\n" + JSON.stringify(tempArray),
        max_tokens: 500,
        temperature: 0.5
    });
    response.then(function(completion){
        let array = JSON.parse(completion.data.choices[0].text.trim())
        context.setAsyncResult(new GC.Spread.CalcEngine.CalcArray(array));
    })
};


GC.Spread.CalcEngine.Functions.defineGlobalCustomFunction("GPT.FILTER", new GPT_Filter());