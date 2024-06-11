const adm_zip = require('adm-zip');
const xml2js = require('xml2js');
const util =  require('node:util');
const default_opt = {
    delimiter: ["{{", "}}"],
    loop_stack: []
}
function escapeRegExp(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    //$&表示整个被匹配的字符串
}

function flat(obj = {}, preKey = "", res = {}) {
    //空值判断，如果obj是空，直接返回
    if(!obj) return
    //获取obj对象的所有[key,value]数组并且遍历，forEach的箭头函数中用了解构
    Object.entries(obj).forEach(([key,value])=>{
        if(Array.isArray(value)){
            //如果obj是数组，那么key就是数组的index，value就是对应的value
            //obj是数组的话就用[]引起来
            //因为value是数组，数组后面是直接跟元素的，不需要.号
            let temp = Array.isArray(obj) ? `${preKey}[${key}]` : `${preKey}${key}`
            flat(value,temp,res)
        }else if(typeof value === 'object'){
            //因为value是对象类型，所以在末尾需要加.号
            let temp = Array.isArray(obj) ? `${preKey}[${key}].` : `${preKey}${key}.`
            flat(value,temp,res)
        }else{
            let temp = Array.isArray(obj) ? `${preKey}[${key}]` : `${preKey}${key}`
            res[temp] = value
        }
    })
    return res;
}

let body_arr_process = function (body_arr, data, opt, func) {
    if (body_arr.length === 0) {
        return;
    }
    for (let body_obj of body_arr) {
        body_process(body_obj, data, opt, func);
    }
}
let body_process = function (body_obj, data, opt, func) {
    if (body_obj.hasOwnProperty("w:p")) {
        paragraph_arr_process(body_obj["w:p"], data, opt, func);
    }
}

let paragraph_arr_process = function (paragraph_arr, data, opt, func) {
    if (paragraph_arr.length === 0) {
        return;
    }
    for (let paragraph_obj of paragraph_arr) {
        paragraph_process(paragraph_obj, data, opt, func);
    }
}

let paragraph_process = function (paragraph_obj, data, opt, func) {
    if (paragraph_obj.hasOwnProperty("w:r")) {
        paragraph_obj["w:r"] = run_arr_process(paragraph_obj["w:r"], data, opt, func);
    }
}

let run_arr_process = function (run_arr, data, opt, func) {
    if (run_arr.length === 0) {
        return;
    }

    return func(run_arr, data, opt);
}

let run_process = function (run_arr, data, opt) {
    for (let run_obj of run_arr) {
        if (run_obj.hasOwnProperty("w:t") && typeof run_obj["w:t"][0] !== "object") {
            run_obj["w:t"][0] = text_process(run_obj["w:t"][0], data, opt, func);
        }
    }
    return run_arr;
}

// 合并属性一样的run，产生新的数组
let merge_runs_by_rPr = function (run_arr, data, opt) {
    let new_run_arr = [run_arr[0]];
    for (let index = 1; index < run_arr.length; ++ index) {
        if (!cmp_run_rPr(new_run_arr[new_run_arr.length - 1], run_arr[index])) {
            new_run_arr.push(run_arr[index]);
            continue;
        }
        if (!new_run_arr[new_run_arr.length - 1]["w:t"][0].hasOwnProperty("_")
            && new_run_arr[new_run_arr.length - 1]["w:t"][0]["$"].hasOwnProperty("xml:space")
            && new_run_arr[new_run_arr.length - 1]["w:t"][0]["$"]["xml:space"] === "preserve") {
            new_run_arr[new_run_arr.length - 1]["w:t"][0]["_"] = " ";
        }
        if (!run_arr[index]["w:t"][0].hasOwnProperty("_")
            && run_arr[index]["w:t"][0]["$"].hasOwnProperty("xml:space")
            && run_arr[index]["w:t"][0]["$"]["xml:space"] === "preserve") {
            run_arr[index]["w:t"][0]["_"] = " ";
        }

        new_run_arr[new_run_arr.length - 1]["w:t"][0]["_"] += run_arr[index]["w:t"][0]["_"];
        new_run_arr[new_run_arr.length - 1]["w:t"][0]["$"] = Object.assign(new_run_arr[new_run_arr.length - 1]["w:t"][0]["$"] || {}, run_arr[index]["w:t"][0]["$"] || {});
    }
    return new_run_arr;
}

let cmp_run_rPr = function (run_a, run_b) {
    return cmp_rPr(run_a["w:rPr"], run_b["w:rPr"]);
}

let cmp_rPr = function (rPr_a_arr, rPr_b_arr) {
    if(util.isDeepStrictEqual(rPr_a_arr, rPr_b_arr)) {
        return true;
    }

    if (rPr_a_arr === undefined || rPr_b_arr === undefined) {
        return false;
    }

    let rPr_a = rPr_a_arr[0];
    let rPr_b = rPr_b_arr[0];
    if (rPr_a.hasOwnProperty("w:rFonts") + rPr_b.hasOwnProperty("w:rFonts") == 1) {
        return false;
    }

    // w:rPr里可能没有w:rFonts。只针对两者都有的话，忽略hint值使之相同。
    if (rPr_a.hasOwnProperty("w:rFonts") + rPr_b.hasOwnProperty("w:rFonts") == 2) {
        if (!rPr_a["w:rFonts"][0]["$"].hasOwnProperty("w:hint") && rPr_b["w:rFonts"][0]["$"].hasOwnProperty("w:hint")) {
            rPr_a["w:rFonts"][0]["$"]["w:hint"] = rPr_b["w:rFonts"][0]["$"]["w:hint"];
        }

        if (!rPr_b["w:rFonts"][0]["$"].hasOwnProperty("w:hint") && rPr_a["w:rFonts"][0]["$"].hasOwnProperty("w:hint")) {
            rPr_b["w:rFonts"][0]["$"]["w:hint"] = rPr_a["w:rFonts"][0]["$"]["w:hint"];
        }
    }

    if (rPr_a.hasOwnProperty("w:lang") && !rPr_b.hasOwnProperty("w:lang")) {
        rPr_b["w:lang"] = rPr_a["w:lang"];
    }
    if (rPr_b.hasOwnProperty("w:lang") && !rPr_a.hasOwnProperty("w:lang")) {
        rPr_a["w:lang"] = rPr_b["w:lang"];
    }

    let result = util.isDeepStrictEqual(rPr_a, rPr_b);
    return result;
}

let text_process = function (text_obj, data, opt, func) {
    // find loop
    let reg_loop = escapeRegExp(opt.delimiter[0]) + "\#.*?" + escapeRegExp(opt.delimiter[1]);
    let reg = new RegExp(reg_loop, "g");
    let loop_matched_arr = text_obj.match(reg);
    if (!loop_matched_arr) {
        return replace_hold(text_obj);
    }
    let left_str = text_obj.slice(0, text_obj.indexOf(loop_matched_arr[0]));
    let right_str = text_obj.slice(text_obj.indexOf(loop_matched_arr) + loop_matched_arr.length);
    let loop_variable_name = loop_matched_arr.slice(opt.delimiter[0].length + 1, 0 - opt.delimiter[0].length);
    if (!data.hasOwnProperty(loop_matched_arr)) {
        return replace_hold(text_obj);
    }

}

let replace_hold = function (template, data, opt) {
    let reg_str = escapeRegExp(opt.delimiter[0]) + ".*?" + escapeRegExp(opt.delimiter[1]);
    let reg = new RegExp(reg_str, "g");
    let text_matched_arr = template.match(reg);
    if (!text_matched_arr) {
        return template;
    }
    for (let text_matched of text_matched_arr) {
        variable_name = text_matched.slice(opt.delimiter[0].length, 0 - opt.delimiter[1].length);
        if (variable_name === "" || !data.hasOwnProperty(variable_name)) {
            continue;
        }
        if (!data.hasOwnProperty(variable_name)) {
            continue;
        }
        template = template.replaceAll(text_matched, data[variable_name]);
    }
    return template;
}

let extend_loop = function (template, data, opt) {
    if (template === undefined || template === "") {
        return template;
    }
    let reg_str = escapeRegExp(opt.delimiter[0]) + "\#.*?" + escapeRegExp(opt.delimiter[1]);
    let reg = new RegExp(reg_str, "g");
    let loop_matched_arr = template.match(reg);
    if (!loop_matched_arr) {
        return replace_hold(template, data, opt);
    }
    let result = "";
    let loop_matched = loop_matched_arr[0];
    let left_str = template.slice(0, template.indexOf(loop_matched));
    let right_str = template.substring(template.indexOf(loop_matched) + loop_matched.length);
    let loop_variable_name = loop_matched.slice(opt.delimiter[0].length + 1, 0 - opt.delimiter[1].length);
    let delimiter_tmp = opt.delimiter[0] + "\/" + loop_variable_name + opt.delimiter[1]
    let pos = right_str.indexOf(delimiter_tmp);
    let middle_str = right_str.slice(0, pos);
    right_str = right_str.slice(pos + delimiter_tmp.length);

    left_str = replace_hold(left_str, data, opt);
    right_str = extend_loop(right_str, data, opt);
    let middle_result = "";
    if (data.hasOwnProperty(loop_variable_name)) {
        let count = data[loop_variable_name].length
        for (let index = 0; index < count; ++ index) {
            middle_result += extend_loop(middle_str, data[loop_variable_name][index], opt);
        }
    }
    return left_str + middle_result + right_str;
}

let render_docx = async function (template, data, opt){
    opt = Object.assign(default_opt, opt);
    const zip = new adm_zip(template);
    let document_xml = zip.readAsText("word/document.xml");
    let parser = new xml2js.Parser({
        trim: false,
        normalize: false,
        explicitCharkey: true,
    });
    let document_obj = await parser.parseStringPromise(document_xml).then(function (result) {
        return result;
    });

    // 先合并相同的run
    if (document_obj.hasOwnProperty("w:document") && document_obj["w:document"].hasOwnProperty("w:body")) {
        body_arr_process(document_obj["w:document"]["w:body"], data, opt, merge_runs_by_rPr);
    }

    // 再还原出xml
    let xml2js_builder = new xml2js.Builder({
        renderOpts: {
            pretty: false
        }
    });
    document_xml = xml2js_builder.buildObject(document_obj);

    //
    data = Object.assign(data, flat(data));
    document_xml = extend_loop(document_xml, data, opt);

    // download image
    // let [err, buf_image] = await utils.download("https://pics1.baidu.com/feed/9d82d158ccbf6c8176395c9eed85af3b32fa4087.jpeg");
    // if (err) {
    //     return res.send({"errno": 1, "message": "No files were uploaded."});
    // }
    // zip.addFile("word/media/download.jpeg", Buffer.from(buf_image, "binary"));

    // build document_xml


    // update document_xml
    zip.updateFile("word/document.xml", document_xml);
    // zip.addLocalFile("/home/zhangke/Myprojects/gostudy-server/zk.jpeg", "word/media");
    return zip.toBuffer();
    // return new Promise((resolve, reject) => {
    //     resolve(zip.toBuffer());
    // }).then(data => {
    //     return [null, data];
    // }).catch(err => {
    //     return [err, null];
    // });
}

let test_msoffice_templater = function () {
    console.log("test for test_msoffice_templater");
}

exports.render_docx = render_docx;
exports.test_msoffice_templater = test_msoffice_templater;