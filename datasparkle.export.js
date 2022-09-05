
// ==UserScript==
// @name         DataSpark导出Excel
// @namespace    http://tampermonkey.net/
// @version      0.1.17
// @description  Final Update
// @author       zhujin
// @match        https://www.datasparkle.net/trackInsight/*
// @require     https://lf9-cdn-tos.bytecdntp.com/cdn/expire-1-M/jquery/3.6.0/jquery.min.js
// @require      http://cdn.staticfile.org/xlsx/0.16.1/xlsx.mini.min.js
// @grant        GM_xmlhttpRequest
// ==/UserScript==

var xmlExportContent = {traffic:[], source:[]};
var yhplHTTPCount = 0;
const EXCEL_TYPE_STORE_DETAIL = 0;
const EXCEL_TYPE_STORE_COMMENT = 1;
const EXCEL_TYPE_STORE_SEARCH_WORD_ZONE = 2;
const EXCEL_TYPE_STORE_PROMOTION = 3;
const EXCEL_TYPE_STORE_GOODS = 4;
const EXCEL_TYPE_STORE_SEARCH_WORD_STORE = 5;

const SEARCH_WORD_ROW = 30;
const SEARCH_WORD_ZONE_API_COUNT = 2;
const SEARCH_WORD_STORE_API_COUNT = 3;
var yhpl_acct_id = "";
(function() {
    'use strict';
    yhplHTTPCount = 0;
    addButton();
    console.log('enter')
})();

function addButton(){
    var searchBtn = $(".ant-card-body button")[0];
    var header = searchBtn.parentElement;
    var button = document.createElement('BUTTON');
    button.innerText = "导出";
    button.onclick = ()=>{getStore(EXCEL_TYPE_STORE_DETAIL)};
    button.id = getButtonIdByType(EXCEL_TYPE_STORE_DETAIL);
    button.className = searchBtn.className;
    $(button).css("margin-left","10px");
    //header.appendChild(button);

    var obtn = $("._btn-download_a2o3c_43 button");
    //obtn[0].onclick = ()=>{getStore(EXCEL_TYPE_STORE_DETAIL)};
    var obtn2 = obtn.clone();
    obtn.parent().hide();
    obtn2.removeAttr("disabled");
    obtn2.css("cursor","pointer");
    obtn2.css("pointer-events","");
    obtn2.click(()=>{getStore(EXCEL_TYPE_STORE_DETAIL)});
    $("._switch-content_12887_71").append(obtn2[0]);
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
        case EXCEL_TYPE_STORE_COMMENT:
            return 'yhplComment';
        case EXCEL_TYPE_STORE_SEARCH_WORD_ZONE:
            return 'yhplSearchZone';
        case EXCEL_TYPE_STORE_SEARCH_WORD_STORE:
            return 'yhplSearchStore';
        case EXCEL_TYPE_STORE_PROMOTION:
            return 'yhplPromotion';
        case EXCEL_TYPE_STORE_GOODS:
            return 'yhplGoods';
        default:
            return 'yhpl';
    }
}


function onStoreSuccessTypeDetail(stores){
    var length = stores.length;
    var totalRequestSize = length * 2;
    for(var poi = 0; poi < length; poi++){
        var child = stores[poi];
        console.log(child.poiName+"," +child.id);
        getStoreDetail(poi * 2, totalRequestSize, child.id, child.poiName);
        getStoreTrafficSource(poi * 2 + 1, totalRequestSize, child.id, child.poiName);
        yhplSleep();
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
                         row.push($("p:eq(1)",$(this)).text());
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

function getStoreName(storeName, storeID){
    if (storeID > 0){
        return (storeName+"("+storeID+")");
    } else {
        return storeName;
    }
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


//endsegment 搜索热词
//segment 营销分析
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
