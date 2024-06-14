const adm_zip = require('adm-zip');
const xml2js = require('xml2js');
const util =  require('node:util');
const https = require('https');
const http = require('http');
const uuid = require('uuid');
const image_size = require('image-size');
const default_opt = {
    delimiter: ["{{", "}}"],
    line_break_index: []
}

// async function download_img(url) {
//     return new Promise((resolve, reject) => {
//         let http_sender;
//         if (url.substring(0, "https".length) === "https") {
//             http_sender = https;
//         } else {
//             http_sender = http;
//         }
//         http_sender.get(url, (res) => {
//             let image_data_arr = [];
//             // res.setEncoding("binary");
//             // res.setEncoding("utf8");
//             res.on('data', (chunk) => {
//                 image_data_arr.push(chunk);
//             });
//             res.on('end', () => {
//                 // 以Unit8Array的格式，来读取图片的宽高
//                 let buffer = Buffer.concat(image_data_arr)
//                 console.log(image_size(buffer));
//
//                 // 再转换成二进制来存放本地
//                 buffer = Buffer.from(buffer.buffer);
//                 return resolve(buffer);
//             });
//         }).on("error", function () {
//             reject("error");
//         });
//     }).then(data => {
//         return [null, data];
//     }).catch(error => {
//         return [error, null];
//     })
// }

async function download_img(url) {
    return new Promise((resolve, reject) => {
        let http_sender;
        if (url.substring(0, "https".length) === "https") {
            http_sender = https;
        } else {
            http_sender = http;
        }
        http_sender.get(url, (res) => {
            let image_data = [];
            res.setEncoding("binary");
            // res.setEncoding("utf8");
            res.on('data', (chunk) => {
                image_data += chunk;
            });
            res.on('end', () => {
                const uint8Array = new Uint8Array(image_data.length);
                for (let i = 0; i < image_data.length; i++) {
                    uint8Array[i] = image_data.charCodeAt(i);
                }
                return resolve({
                    image_data: image_data,
                    info: image_size(uint8Array)
                });
            });
        }).on("error", function () {
            reject("error");
        });
    }).then(data => {
        return [null, data];
    }).catch(error => {
        return [error, null];
    })
}

function escapeXml(unsafe) {
    unsafe = unsafe.toString();
    return unsafe.replace(/[<>&'"]/g, function (c) {
        switch (c) {
            case '<': return '&lt;';
            case '>': return '&gt;';
            case '&': return '&amp;';
            case '\'': return '&apos;';
            case '"': return '&quot;';
        }
    });
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

let body_arr_process = async function (body_arr, data, opt, func) {
    if (body_arr.length === 0) {
        return;
    }
    for (let body_obj of body_arr) {
        await body_process(body_obj, data, opt, func);
    }
}
let body_process = async function (body_obj, data, opt, func) {
    if (body_obj.hasOwnProperty("w:p")) {
        body_obj["w:p"] = await paragraph_arr_process(body_obj["w:p"], data, opt, func);
    }
}

let paragraph_arr_process = async function (paragraph_arr, data, opt, func) {
    if (paragraph_arr.length === 0) {
        return paragraph_arr;
    }

    if (func.hasOwnProperty("paragraph_arr_process")) {
        return await func.paragraph_arr_process(paragraph_arr, data, opt, func);
    }
    for (let paragraph_obj of paragraph_arr) {
        await paragraph_process(paragraph_obj, data, opt, func);

        // // 处理可能出现的段落调整
        // if (opt.line_break_index.length > 0) {
        //     let new
        //     for (let index = 0; index < paragraph_obj["w:r"].length; ++ index) {
        //         paragraph_obj["w:r"][index]
        //     }
        // }
    }
    return paragraph_arr;
}

let paragraph_arr_process_for_html_paragraph_split = async function (paragraph_arr, data, opt, func) {
    let new_paragraph_arr = [];
    for (let paragraph_obj of paragraph_arr) {
        if (!paragraph_obj.hasOwnProperty("w:r")) {
            new_paragraph_arr.push(paragraph_obj);
            continue;
        }
        await paragraph_process(paragraph_obj, data, opt, func);

        // 处理可能出现的段落调整
        let new_run_arr = [];
        for (let index = 0; index < paragraph_obj["w:r"].length; ++ index) {
            new_run_arr.push(paragraph_obj["w:r"][index]);
            if (paragraph_obj["w:r"][index].hasOwnProperty("line_break")
            && paragraph_obj["w:r"][index]["line_break"]) {
                new_paragraph_arr.push({
                    "w:r": new_run_arr
                });
                new_run_arr = [];
                delete paragraph_obj["w:r"][index]["line_break"];
            }
        }
        if (new_run_arr.length > 0) {
            new_paragraph_arr.push({
                "w:r": new_run_arr
            });
        }
    }
    return new_paragraph_arr;
}

let paragraph_process = async function (paragraph_obj, data, opt, func) {
    if (paragraph_obj.hasOwnProperty("w:r")) {
        paragraph_obj["w:r"] = await run_arr_process(paragraph_obj["w:r"], data, opt, func);
    }
}

let run_arr_process = async function (run_arr, data, opt, func) {
    if (run_arr.length === 0) {
        return run_arr;
    }

    if (func.hasOwnProperty("run_arr_process")) {
        return await func.run_arr_process(run_arr, data, opt, func);
    }
}
// 检查每一个run，查看里面是否需要分段
let run_arr_process_for_html_paragraph_split = async function (run_arr, data, opt) {
    opt.line_break_index = [];
    let new_run_arr = [];
    for (let index = 0; index < run_arr.length; ++ index) {
        if (!run_arr[index].hasOwnProperty("w:t") || run_arr[index].length === 0 || !run_arr[index]["w:t"][0].hasOwnProperty("_")) {
            new_run_arr.push(run_arr[index]);
            continue;
        }

        let matched_arr = run_arr[index]["w:t"][0]["_"].match(/(?<=<p>).*?(?=<\/p>)/g);
        if (!matched_arr) {
            new_run_arr.push(run_arr[index]);
            continue;
        }

        let text_arr = run_arr[index]["w:t"][0]["_"].split('</p><p>');
        for (let i = 0; i < text_arr.length; ++ i) {
            let text = text_arr[i].replaceAll('<p>', "").replaceAll('</p>', "").replaceAll('&nbsp;', ' ');
            let run = {};
            if (run_arr[index].hasOwnProperty("w:rPr")) {
                run["w:rPr"] = run_arr[index]["w:rPr"];
            }
            run["w:t"] = [{"_": text}];
            if (i < text_arr.length - 1) {
                run["line_break"] = true;
            }
            new_run_arr.push(run);
        }

        //
        // let pos = run_arr[index]["w:t"][0]["_"].indexOf("<p>");
        // let left = run_arr[index]["w:t"][0]["_"].substring(0, pos);
        // pos = run_arr[index]["w:t"][0]["_"].lastIndexOf("<\/p>");
        // let right = run_arr[index]["w:t"][0]["_"].substring(pos + "<\/p>".length);
        //
        // if (left !== "") {
        //     let run = {};
        //     if (run_arr[index].hasOwnProperty("w:rPr")) {
        //         run["w:rPr"] = run_arr[index]["w:rPr"];
        //     }
        //     run["w:t"] = [{"_": left}];
        //     new_run_arr.push(run);
        // }
        // for (let i = 0; i < matched_arr.length; ++ i) {
        //     let run = {};
        //     if (run_arr[index].hasOwnProperty("w:rPr")) {
        //         run["w:rPr"] = run_arr[index]["w:rPr"];
        //     }
        //     run["w:t"] = [{"_": matched_arr[i]}];
        //     if (i < matched_arr.length - 1) {
        //         run["line_break"] = true;
        //     }
        //     new_run_arr.push(run);
        // }
        // if (right !== "") {
        //     let run = {};
        //     if (run_arr[index].hasOwnProperty("w:rPr")) {
        //         run["w:rPr"] = run_arr[index]["w:rPr"];
        //     }
        //     run["w:t"] = [right];
        //     new_run_arr.push(run);
        // }
    }
    return new_run_arr;
}

let replace_html_img = async function (run, data, opt) {
    if (!run.hasOwnProperty("w:t") || run["w:t"].length === 0 || !run["w:t"][0].hasOwnProperty("_")) {
        return [run];
    }
    let new_run_arr = [];
    let matched_arr = run["w:t"][0]["_"].match(/(<img\ssrc=).*?(\/>)/g);
    if (!matched_arr) {
        new_run_arr.push(run);
        return [run];
    }

    let left_pos = run["w:t"][0]["_"].indexOf(matched_arr[0]);
    let left = run["w:t"][0]["_"].substring(0, left_pos);
    let right_pos = left_pos + matched_arr[0].length;
    let right = run["w:t"][0]["_"].substring(right_pos);

    // for left
    if (left !== "") {
        let run_tmp = {};
        if (run.hasOwnProperty("w:rPr")) {
            run_tmp["w:rPr"] = JSON.parse(JSON.stringify(run["w:rPr"]));
        }
        run_tmp["w:t"] = [{"_": left}];
        new_run_arr.push(run_tmp);
    }

    // for middle matched
    let run_tmp = {};
    run_tmp["w:rPr"] = [{"w:noProof":[""]}]

    let src = matched_arr[0].match(/(?<=src=").*?(?="\s)/g);
    let img_filename = "";
    let rid = "rId" + (opt["document_rels_obj"]["Relationships"]["Relationship"].length + 1).toString();
    if (src !== undefined && src.length > 0) {
        let width = matched_arr[0].match(/(?<=width:\s).*?(?=%)/g);
        if (width === undefined) {
            width = 100;
        } else {
            width = Number(width[0]);
        }

        let [err, image] = await download_img(src[0]);
        image.info.width = parseInt(image.info.width * 8325 * width / 100).toString();
        image.info.height = parseInt(image.info.height * 8325 * width / 100).toString();
        img_filename = uuid.v4();
        opt["zip"].addFile("word/media/" + img_filename + ".jpeg", Buffer.from(image.image_data, "binary"));
        opt["document_rels_obj"]["Relationships"]["Relationship"].push({
            "$": {
                Id: rid,
                Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                Target: "media/" + img_filename + ".jpeg"
            }
        });
        run_tmp["w:drawing"] = [
            {
                "wp:inline": [
                    {
                        "$": {
                            "distT": "0",
                            "distB": "0",
                            "distL": "0",
                            "distR": "0",
                            "wp14:anchorId": "6406A3DD",
                            "wp14:editId": "6DC94A05"
                        },
                        "wp:extent": [
                            {
                                "$": {
                                    "cx": image.info.width,
                                    "cy": image.info.height
                                }
                            }
                        ],
                        "wp:effectExtent": [
                            {
                                "$": {
                                    "l": "0",
                                    "t": "0",
                                    "r": "0",
                                    "b": "0"
                                }
                            }
                        ],
                        "wp:docPr": [
                            {
                                "$": {
                                    "id": opt["img_index"].toString(),
                                    "name": "图片 " + opt["img_index"].toString()
                                }
                            }
                        ],
                        "wp:cNvGraphicFramePr": [
                            {
                                "a:graphicFrameLocks": [
                                    {
                                        "$": {
                                            "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
                                            "noChangeAspect": "1"
                                        }
                                    }
                                ]
                            }
                        ],
                        "a:graphic": [
                            {
                                "$": {
                                    "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main"
                                },
                                "a:graphicData": [
                                    {
                                        "$": {
                                            "uri": "http://schemas.openxmlformats.org/drawingml/2006/picture"
                                        },
                                        "pic:pic": [
                                            {
                                                "$": {
                                                    "xmlns:pic": "http://schemas.openxmlformats.org/drawingml/2006/picture"
                                                },
                                                "pic:nvPicPr": [
                                                    {
                                                        "pic:cNvPr": [
                                                            {
                                                                "$": {
                                                                    "id": opt["img_index"].toString(),
                                                                    "name": ""
                                                                }
                                                            }
                                                        ],
                                                        "pic:cNvPicPr": [
                                                            ""
                                                        ]
                                                    }
                                                ],
                                                "pic:blipFill": [
                                                    {
                                                        "a:blip": [
                                                            {
                                                                "$": {
                                                                    "r:embed": rid
                                                                }
                                                            }
                                                        ],
                                                        "a:stretch": [
                                                            {
                                                                "a:fillRect": [
                                                                    ""
                                                                ]
                                                            }
                                                        ]
                                                    }
                                                ],
                                                "pic:spPr": [
                                                    {
                                                        "a:xfrm": [
                                                            {
                                                                "a:off": [
                                                                    {
                                                                        "$": {
                                                                            "x": "0",
                                                                            "y": "0"
                                                                        }
                                                                    }
                                                                ],
                                                                "a:ext": [
                                                                    {
                                                                        "$": {
                                                                            "cx": image.info.width,
                                                                            "cy": image.info.height
                                                                        }
                                                                    }
                                                                ]
                                                            }
                                                        ],
                                                        "a:prstGeom": [
                                                            {
                                                                "$": {
                                                                    "prst": "rect"
                                                                },
                                                                "a:avLst": [
                                                                    ""
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

        new_run_arr.push(run_tmp);
    }


    // for right
    let right_run_arr = [];
    if (right !== "") {
        run["w:t"][0]["_"] = right;
        opt["img_index"] += 1;
        right_run_arr = await replace_html_img(run, data, opt);
        new_run_arr = new_run_arr.concat(right_run_arr);
    }
    return new_run_arr;
}

let run_arr_process_for_html_img = async function (run_arr, data, opt) {
    console.log(run_arr);
    let new_run_arr = [];
    for (let index = 0; index < run_arr.length; ++ index) {
        opt["img_index"] = 1;
        let new_run_arr_tmp = await replace_html_img(run_arr[index], data, opt);
        new_run_arr = new_run_arr.concat(new_run_arr_tmp);
    }
    return new_run_arr;
}

let replace_html_u = async function (run, data, opt) {
    if (!run.hasOwnProperty("w:t") || run["w:t"].length === 0 || !run["w:t"][0].hasOwnProperty("_")) {
        return [run];
    }
    let new_run_arr = [];
    let matched_arr = run["w:t"][0]["_"].match(/(?<=<u>).*?(?=<\/u>)/g);
    if (!matched_arr) {
        new_run_arr.push(run);
        return [run];
    }

    let pos = run["w:t"][0]["_"].indexOf("<u>");
    let left = run["w:t"][0]["_"].substring(0, pos);
    pos = run["w:t"][0]["_"].indexOf("<\/u>");
    let right = run["w:t"][0]["_"].substring(pos + "<\/u>".length);

    // for left
    if (left !== "") {
        let run_tmp = {};
        if (run.hasOwnProperty("w:rPr")) {
            run_tmp["w:rPr"] = JSON.parse(JSON.stringify(run["w:rPr"]));
        }
        run_tmp["w:t"] = [{"_": left}];
        new_run_arr.push(run_tmp);
    }

    // for middle matched
    let run_tmp = {};
    if (run.hasOwnProperty("w:rPr")) {
        run_tmp["w:rPr"] = JSON.parse(JSON.stringify(run["w:rPr"]));
    }
    else {
        run_tmp["w:rPr"] = [{}];
    }

    let w_u = [{
        "$": {
            "w:val": "single"
        }
    }];
    if (typeof run_tmp["w:rPr"][0] === "object") {
        run_tmp["w:rPr"][0]["w:u"] = w_u
    } else {
        run_tmp["w:rPr"][0] = {
            "w:u": w_u
        }
    }
    run_tmp["w:t"] = [{"_": matched_arr[0]}];
    new_run_arr.push(run_tmp);

    // for right
    let right_run_arr = [];
    if (right !== "") {
        run["w:t"][0]["_"] = right;
        right_run_arr = await replace_html_u(run, data, opt);
        new_run_arr = new_run_arr.concat(right_run_arr);
    }
    return new_run_arr;
}

let run_arr_process_for_html_tag = async function (run_arr, data, opt) {
    let new_run_arr = [];
    for (let index = 0; index < run_arr.length; ++ index) {
        let new_run_arr_tmp = await replace_html_u(run_arr[index], data, opt);
        new_run_arr = new_run_arr.concat(new_run_arr_tmp);
    }
    return new_run_arr;
}
// let run_process = function (run_arr, data, opt) {
//     for (let run_obj of run_arr) {
//         if (run_obj.hasOwnProperty("w:t") && typeof run_obj["w:t"][0] !== "object") {
//             run_obj["w:t"][0] = text_process(run_obj["w:t"][0], data, opt, func);
//         }
//     }
//     return run_arr;
// }

// 检查每一个run，合并属性一样的run，产生新的数组
let run_arr_process_merge_runs_by_rPr = async function (run_arr, data, opt) {
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

let rPr_is_empty = function (rPr) {
    if (rPr === undefined || rPr.length == 0 || Object.keys(rPr[0]).length == 0) {
        return true;
    }
    return false;
}

let clear_rPr = function (rPr) {
    try {
        delete rPr[0]["w:rFonts"][0]["$"]["w:hint"];
    }
    catch (e) {}
    try {
        if (Object.keys(rPr[0]["w:rFonts"][0]["$"]).length == 0) {
            delete rPr[0]["w:rFonts"];
        }
    } catch (e) {}
    try {
        delete rPr[0]["w:lang"];
    }
    catch (e) {}
}

let cmp_rPr = function (rPr_a_arr, rPr_b_arr) {
    clear_rPr(rPr_a_arr);
    clear_rPr(rPr_b_arr);
    if (rPr_is_empty(rPr_a_arr) && (rPr_is_empty(rPr_b_arr))) {
        return true;
    }
    if (rPr_is_empty(rPr_a_arr) + rPr_is_empty(rPr_b_arr) == 1) {
        return false;
    }

    let rPr_a = rPr_a_arr[0];
    let rPr_b = rPr_b_arr[0];
    // if (rPr_a.hasOwnProperty("w:rFonts") + rPr_b.hasOwnProperty("w:rFonts") == 1) {
    //     return false;
    // }

    if (rPr_a.hasOwnProperty("w:rFonts") && rPr_a["w:rFonts"][0]["$"].hasOwnProperty("w:hint")) {
        delete rPr_a["w:rFonts"][0]["$"]["w:hint"];
        if (Object.keys(rPr_a["w:rFonts"][0]["$"]).length === 0) {
            delete rPr_a["w:rFonts"];
        }
    }
    if (rPr_b.hasOwnProperty("w:rFonts") && rPr_b["w:rFonts"][0]["$"].hasOwnProperty("w:hint")) {
        delete rPr_b["w:rFonts"][0]["$"]["w:hint"];
        if (Object.keys(rPr_b["w:rFonts"][0]["$"]).length === 0) {
            delete rPr_b["w:rFonts"];
        }
    }
    // w:rPr里可能没有w:rFonts。只针对两者都有的话，忽略hint值使之相同。
    // if (rPr_a.hasOwnProperty("w:rFonts") + rPr_b.hasOwnProperty("w:rFonts") == 2) {
    //     if (!rPr_a["w:rFonts"][0]["$"].hasOwnProperty("w:hint") && rPr_b["w:rFonts"][0]["$"].hasOwnProperty("w:hint")) {
    //         rPr_a["w:rFonts"][0]["$"]["w:hint"] = rPr_b["w:rFonts"][0]["$"]["w:hint"];
    //     }
    //
    //     if (!rPr_b["w:rFonts"][0]["$"].hasOwnProperty("w:hint") && rPr_a["w:rFonts"][0]["$"].hasOwnProperty("w:hint")) {
    //         rPr_b["w:rFonts"][0]["$"]["w:hint"] = rPr_a["w:rFonts"][0]["$"]["w:hint"];
    //     }
    // }

    if (rPr_a.hasOwnProperty("w:lang")) {
        delete rPr_a["w:lang"];
    }
    if (rPr_b.hasOwnProperty("w:lang")) {
        delete rPr_b["w:lang"];
    }

    let result = util.isDeepStrictEqual(rPr_a_arr, rPr_b_arr);
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
        template = template.replaceAll(text_matched, escapeXml(data[variable_name]));
    }
    return template;
}

let replace_with_extend = function (template, data, opt) {
    if (template === undefined || template === "") {
        return template;
    }
    let reg_str = escapeRegExp(opt.delimiter[0]) + "\#.*?" + escapeRegExp(opt.delimiter[1]);
    let reg = new RegExp(reg_str, "g");
    let loop_matched_arr = template.match(reg);
    if (!loop_matched_arr) {
        return replace_hold(template, data, opt);
    }
    let loop_matched = loop_matched_arr[0];
    let left_str = template.slice(0, template.indexOf(loop_matched));
    let right_str = template.substring(template.indexOf(loop_matched) + loop_matched.length);
    let loop_variable_name = loop_matched.slice(opt.delimiter[0].length + 1, 0 - opt.delimiter[1].length);
    let delimiter_tmp = opt.delimiter[0] + "\/" + loop_variable_name + opt.delimiter[1]
    let pos = right_str.indexOf(delimiter_tmp);
    let middle_str = right_str.slice(0, pos);
    right_str = right_str.slice(pos + delimiter_tmp.length);

    left_str = replace_hold(left_str, data, opt);
    right_str = replace_with_extend(right_str, data, opt);
    let middle_result = "";
    if (data.hasOwnProperty(loop_variable_name)) {
        let count = data[loop_variable_name].length
        for (let index = 0; index < count; ++ index) {
            middle_result += replace_with_extend(middle_str, data[loop_variable_name][index], opt);
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
    await body_arr_process(document_obj["w:document"]["w:body"], data, opt, {
        run_arr_process: run_arr_process_merge_runs_by_rPr
    });

    // 再还原出xml
    let xml2js_builder = new xml2js.Builder({
        renderOpts: {
            pretty: true
        }
    });
    document_xml = xml2js_builder.buildObject(document_obj);

    // 利用正则表达式来展开循环，替换placeholder
    document_xml = replace_with_extend(document_xml, data, opt);

    // 再次解析，
    document_obj = await parser.parseStringPromise(document_xml).then(function (result) {
        return result;
    });
    await body_arr_process(document_obj["w:document"]["w:body"], data, opt, {
        paragraph_arr_process: paragraph_arr_process_for_html_paragraph_split,
        run_arr_process: run_arr_process_for_html_paragraph_split
    });

    // 处理内容中的html<u>字符
    await body_arr_process(document_obj["w:document"]["w:body"], data, opt, {
        run_arr_process: run_arr_process_for_html_tag
    });

    // 处理内容中的html<img>图片
    let document_rels_xml = zip.readAsText("word/_rels/document.xml.rels");
    let document_rels_obj = await parser.parseStringPromise(document_rels_xml).then(function (result) {
        return result;
    });
    let content_types_xml = zip.readAsText("[Content_Types].xml");
    let content_types_obj = await parser.parseStringPromise(content_types_xml).then(function (result) {
        return result;
    });
    opt["document_rels_obj"] = document_rels_obj;
    opt["zip"] = zip;
    await body_arr_process(document_obj["w:document"]["w:body"], data, opt, {
        run_arr_process: run_arr_process_for_html_img
    });
    content_types_obj["Types"]["Default"].push({
        "$":{Extension: "png", ContentType: "image/png"}
    });
    content_types_obj["Types"]["Default"].push({
        "$":{Extension: "jpeg", ContentType: "image/jpeg"}
    });
    document_xml = xml2js_builder.buildObject(document_obj);
    document_rels_xml = xml2js_builder.buildObject(opt["document_rels_obj"]);
    content_types_xml = xml2js_builder.buildObject(content_types_obj);

    // update document_xml
    zip.updateFile("word/_rels/document.xml.rels", document_rels_xml);
    zip.updateFile("word/document.xml", document_xml);
    zip.updateFile("[Content_Types].xml", content_types_xml);
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