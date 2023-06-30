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
    
    let tempArray = range.toArray && range.toArray(undefined, false);
    if (!Array.isArray(tempArray)) {
        return GC.Spread.CalcEngine.Errors.NotAvailable;
    }
    desc = desc.replaceAll("```", "").replaceAll('"""').replaceAll("\n", "");
    let delimiter = "####";
    let messages = [
        {"role": "system", "content": `你将根据用户的要求对数据做处理。`
                                    + `用户的处理要求将用${delimiter}分隔.`
                                    + `如果处理要求和数据无关，直接返回"NA".`
                                    + `先按用户的要求处理数据，然后将处理的数据通过JSON的数组返回，并且JSON 符合JSON schema { "type": "array", "items": { "type": "array", "items": { "type": [ "string", "number" ] } } } `
                                    + `数据是：${JSON.stringify(tempArray)}`},
        { "role": "user", "content": `${delimiter}${desc}${delimiter}`
    }];

    let response = openai.createChatCompletion({
        "model": "gpt-3.5-turbo-0613",
        "messages": messages,
    });
    response.then(function(completion){
        let text = completion.data.choices[0].message.content.trim();
        // text = text[0] === "[" ? text : text.substring(text.indexOf("["));
        let array = JSON.parse(text)
        context.setAsyncResult(new GC.Spread.CalcEngine.CalcArray(array));
    })
    /* 
    const response = openai.createCompletion({
        model: "text-davinci-003",
        prompt: '按照如下步骤处理数据：\n' 
        + '1. 按照处理要求对JSON数据进行处理，\n' 
        + '2. 格式化你的回复为JSON数组，数组符合JSON schema { "type": "array", "items": { "type": "array", "items": { "type": ["string", "number"] } } } 。\n'
        + '3. 直接返回JSON文本，不要添加提示内容\n'
        + 'JSON数据：'+ JSON.stringify(tempArray)
        + '\n处理要求：\n'
        + ' """\n'
        + desc
        + '\n"""', 
        max_tokens: 500,
        temperature: 0.3
    });
    response.then(function(completion){
        let text = completion.data.choices[0].text.trim();
        // text = text[0] === "[" ? text : text.substring(text.indexOf("["));
        let array = JSON.parse(text)
        context.setAsyncResult(new GC.Spread.CalcEngine.CalcArray(array));
    })
    */


};


GC.Spread.CalcEngine.Functions.defineGlobalCustomFunction("GPT.FILTER", new GPT_Filter());