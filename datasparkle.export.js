// ==UserScript==
// @name         DataSparkle导出Excel
// @namespace    http://tampermonkey.net/
// @version      0.1.1
// @description  Final Update
// @author       zhujin
// @match        https://www.datasparkle.net/*
// @require      https://lf9-cdn-tos.bytecdntp.com/cdn/expire-1-M/jquery/3.6.0/jquery.min.js
// @require      http://cdn.staticfile.org/xlsx/0.16.1/xlsx.mini.min.js
// @grant        GM_xmlhttpRequest
// ==/UserScript==

var xmlExportContent = {traffic:[], source:[]};
const EXCEL_TYPE_STORE_DETAIL = 0;

var LAST_LOCATION = "";
var yhpl_acct_id = "";
(function() {
    'use strict';
    document.body.onclick = function(){
        console.log(window.location.href);
        if(LAST_LOCATION!=window.location.href){
            LAST_LOCATION = window.location.hre
            if(window.location.href.indexOf("useLeaderboards")>0){
                addButton();
                console.log('enter')
            }
        }
    };
})();

function addButton(){
if($("#export_ext").length==0){
    console.log('add btn')

    var obtn = $("._btn-download_a2o3c_43 button");
    //obtn[0].onclick = ()=>{getStore(EXCEL_TYPE_STORE_DETAIL)};
    var obtn2 = obtn.clone();
    obtn.parent().hide();
    obtn2.attr("id","export_ext");
    obtn2.removeAttr("disabled");
    obtn2.css("cursor","pointer");
    obtn2.css("pointer-events","");
    obtn2.click(()=>{getStore(EXCEL_TYPE_STORE_DETAIL)});
    $("._switch-content_12887_71").append(obtn2[0]);
}
}
function appendButton(parent, type, text){
    var bt = document.createElement('BUTTON');
    bt.innerText = text;
    bt.onclick = ()=>{getStore(type)};
    bt.id = getButtonIdByType(type);
    parent.appendChild(bt);
}
function getButtonIdByType(type){
    switch (type){
        case EXCEL_TYPE_STORE_DETAIL:
            return 'yhplExport';
        default:
            return 'yhpl';
    }
}

async function yhplSleep() {
    await sleep(200)
    //  console.log('yhplSleep end!')
}

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms))
}

function onStoreError(type, e){
    console.log(e);
    console.log('onError');
}
function getStore(type){
    console.log('getStore enter type: '+type);
    var ths = $(".ant-table-thead th");
    xmlExportContent = {traffic:[], source:[]};

    var thname = [];
    ths.each(function(idx){
        console.log(idx+":"+$(this).text());
        if(idx>0&&idx<ths.length){
            thname.push($(this).text());
        }
    });
    xmlExportContent.traffic.push(thname);

    var tb = $(".ant-table-tbody tr");
    tb.each(function(idx){
        if(idx>0){
            var tds = $("td",$(this));
            var row = [];
            tds.each(function(tdidx){
                if(tdidx>0&&tdidx<ths.length){
                    if(tdidx==2){
                         row.push($("p:eq(0)",$(this)).text());
                    }else if(tdidx==4){
                          row.push($(this).text().replace(/\,/g,""));
                    }else{
                        row.push($(this).text());
                    }
                }
            });
            xmlExportContent.traffic.push(row);
        }
    });
    showProgress(type,1,10);
}

function getRateString(rate){
    if (!rate){
        return "";
    }
    if(rate == 'null'){
        return "";
    }
    return rate+"%";
}


function appendDateSuffix(index, val){
    if (val == null){
        return "";
    }
    if (index.indexOf('RATE')>=0){
        val = val + "%";
    }
    return val;
}

function parseJson(text){
    var node = null;
    try {
        node = JSON.parse(text);
    } catch(e){
        console.log('parseJson error:'+e);
    }
    return node;
}

function getDateString(d){
    var beginTime = []
    var month = d.getMonth() + 1;
    var day = d.getDate()
    beginTime[0] = d.getFullYear();
    beginTime[1] = month < 10 ? '0' + month : month;
    beginTime[2] = day < 10 ? '0' + day : day;
    return beginTime.join('');
}
function getFieldString(field){
    return field ? field:'';
}


function showProgress(type, step, total){
    exportAsXLS(type, xmlExportContent);
}

//segment excel
function exportAsXLS(type, table){
    switch (type){
        case EXCEL_TYPE_STORE_DETAIL:{
            var sheet = XLSX.utils.aoa_to_sheet(table.traffic);
            openDownloadDialog(sheet2blob([{sheet: sheet,name:'排行榜'}]), '排行榜.xlsx');
            break;
        } 
    }
}
function sheet2blob(sheets) {
    var sheetsSize = sheets.length;
    var SheetNames = [];
    var Sheets = {};

    for (var index = 0; index<sheets.length; index++){
        var child = sheets[index];
        var sheetName = child.name || 'sheet'+(index+1);
        SheetNames.push(sheetName);
        var sheet = child.sheet;
        Sheets[sheetName] = sheet;
    }

    var workbook = {
        SheetNames: SheetNames,
        //Sheets: {}
        Sheets: Sheets
    };
    //workbook.Sheets[sheetName] = sheet; // 生成excel的配置项

    var wopts = {
        bookType: 'xlsx', // 要生成的文件类型
        bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
        type: 'binary'
    };
    var wbout = XLSX.write(workbook, wopts);
    var blob = new Blob([s2ab(wbout)], {
        type: "application/octet-stream"
    }); // 字符串转ArrayBuffer
    function s2ab(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }
    return blob;
}
function openDownloadDialog(url, saveName) {
    if (typeof url == 'object' && url instanceof Blob) {
        url = URL.createObjectURL(url); // 创建blob地址
    }
    var aLink = document.createElement('a');
    aLink.href = url;
    aLink.download = saveName || ''; // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，file:///模式下不会生效
    var event;
    if (window.MouseEvent) event = new MouseEvent('click');
    else {
        event = document.createEvent('MouseEvents');
        event.initMouseEvent('click', true, false, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
    }
    aLink.dispatchEvent(event);
}
//eng segment excel
