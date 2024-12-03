/* eslint-disable */
jQuery.noConflict();
(function($, PLUGIN_ID) {
    "use strict";

    const APPP_ID = kintone.app.getId();
    const config = kintone.plugin.app.getConfig(PLUGIN_ID);
    const EXCEL_FILE_APPP_ID = JSON.parse(config.excelAppId);
    const APPP_ID_RECORD_NAME = 'appId';
    const ORDER_NO_RECORD_NAME = 'orderNo';
    const VISIBLE_RECORD_NAME = 'visible';
    var initFlg = true;
    
    kintone.events.on("app.record.detail.show", async function() {
        if (initFlg) {
            await init();
            initFlg = false;
        }
    });
    
    $(document).on('click', '.excel-output', async function() {
        var id = $(this).data('id');
        output(id);
    });

    async function init() {
        try {
            var query = APPP_ID_RECORD_NAME + ' = ' + APPP_ID + ' and ' + VISIBLE_RECORD_NAME + ' in ("表示") order by ' + ORDER_NO_RECORD_NAME + ' asc';
            const body = {
                app: EXCEL_FILE_APPP_ID,
                query: query
            };

            const resp = await kintone.api(kintone.api.url('/k/v1/records.json', true), 'GET', body);
            if(resp.records.length > 0) {
                var div = $('<div class="gaia-argoui-app-toolbar-statusmenu"><div class="gaia-app-statusbar" style=""><div class="gaia-app-statusbar-actions"><div class="gaia-app-statusbar-actionlist-wrapper"></div><div class="gaia-app-statusbar-actionmenu-wrapper excel-output-action"><div class="gaia-app-statusbar-actionmenu"></div></div></div></div></div>');
                $.each(resp.records, function(i, r){
                    if(r.btnName != '') {
                        div.find('.gaia-app-statusbar-actionmenu').append($('<span class="gaia-app-statusbar-action"><span class="gaia-app-statusbar-action-label excel-output" data-id="' + r.$id.value + '">' + r.btnName.value + '</span></span>'));
                    }
                });
                $('.gaia-argoui-app-toolbar-statusmenu').append(div);
            }
        } catch (e) {
            console.error(e);
        }
    }
    
    async function output(configRecordId) {
        try {
            const recordId = kintone.app.record.getId();

            var body = {
                app: kintone.app.getId()
            };
            
            const fieldsResponse = await kintone.api(kintone.api.url('/k/v1/app/form/fields.json', true), 'GET', body);
            const fields = fieldsResponse.properties;
            
            body = {
                app: APPP_ID,
                id: recordId
            };

            const detail = await kintone.api(kintone.api.url('/k/v1/record.json', true), 'GET', body);
            const detailRecord = detail.record;
            const sortedKey = Object.keys(detailRecord).sort();
            var array = [];

            $.each(sortedKey, function(keyIndex, keyValue) {
                var label = "";
                var code = "";
                var value = "";
                
                if (fields.hasOwnProperty(keyValue)) {
                    label = fields[keyValue].label;
                    code = fields[keyValue].code;
                } else {
                    code = keyValue;
                }

                var record = detailRecord[keyValue];
                
                // SUBTABLEの場合
                if (record.type === "SUBTABLE") {
                    // レコードがある場合
                    if (record.value.length > 0) {
                        var sortedSubKey = Object.keys(record.value[0].value).sort();
                        var tmp = [];

                        // サブテーブル内のフィールド情報
                        var subTableFields = fields[keyValue]["fields"];

                        $.each(sortedSubKey, function(subKeyIndex, subKeyValue) {
                            var row = [];
                            $.each(record.value, function(subIndex, subRecord) {
                                value = getData(subRecord.value[subKeyValue], subTableFields[subKeyValue]);
                                const subItemLabel = subTableFields[subKeyValue].label;
                                const subItemCode = subKeyValue;
                                row.push([label, code, subItemLabel, subItemCode, value]);
                            });
                            tmp.push(row);
                        });
                        array.push(tmp);
                    } else {
                        array.push([label, code, '']);
                    }
                } else {
                    value = getData(record, fields[keyValue]);
                    array.push([label, code, value]);
                }
            });
            
            var appId = EXCEL_FILE_APPP_ID;
            var fieldCode = 'appId';
            var value = kintone.app.getId();
            var query = fieldCode + ' = ' + value + ' and ' + '$id = ' + configRecordId;
        
            // レコード取得
            const resp = await kintone.api('/k/v1/records', 'GET', { app: appId, query: query });

            const record = resp.records;

            const recordFileValue = record[0]['templeteFile'].value[0];

            const url =location.origin + "/k/v1/file.json?fileKey=" + recordFileValue.fileKey;

            const req = new XMLHttpRequest();
            req.open("GET", url);
            req.setRequestHeader("X-Requested-With", "XMLHttpRequest");
            req.responseType = "arraybuffer";

            // テンプレート読み込み
            req.onload = async (e) => {
                const wb = await XlsxPopulate.fromDataAsync(req.response);
                const ws = wb.sheet('kintoneレコード情報');

                if(ws) {
                    var n = 0;
                    for (var i = 0; i < array.length; i ++) {
                        var alphabet = getExcelColumnLabel(n);
                        if(Array.isArray(array[i][0])) {
                            var colIdx = n;
                            $.each(array[i], function(idx, row){
                                for(var j = 0; j < row.length; j++) {
                                    alphabet = getExcelColumnLabel(colIdx);
                                    if(j == 0) {
                                        // // フィールド名(サブテーブル)
                                        // ws.cell(alphabet + 1).value(row[j][0]);

                                        // // フィールドコード(サブテーブル)
                                        // ws.cell(alphabet + 2).value(row[j][1]);

                                        // フィールド名
                                        ws.cell(alphabet + 1).value(row[j][2]);

                                        // フィールドコード
                                        ws.cell(alphabet + 2).value(row[j][3]);
                                    }
                                    // 値
                                    ws.cell(alphabet + (3 + j)).value(row[j][4]);
                                }
                                colIdx++;
                            });
                            n = colIdx;
                        } else {
                            ws.cell(alphabet + 1).value(array[i][0].split('￥')[0]);
                            ws.cell(alphabet + 2).value(array[i][1]);
                            ws.cell(alphabet + 3).value(array[i][2]);
                            n++;
                        }
                    }
                }

                const wbout = await wb.outputAsync();
                const blob = new Blob([wbout], {
                    type: 'application/octet-stream'
                });

                const fileNameSplit = recordFileValue.name.split("/").reverse()[0].split('.');
                const file_name = fileNameSplit[0];
                const extend  = fileNameSplit[1];

                saveAs(blob, file_name + '_' + recordId + '.' + extend);
            };

            req.send();
        } catch (error) {
            alert("エラーが発生しました。\n" + error.name + " : " + error.message);
        }
    }
    
    function getData(data, property) {
        try {
            // value値の取得
            if (data.value) {
                // name属性
                if (data.value.hasOwnProperty('name')) {
                    return data.value.name;
                    
                } else if (Array.isArray(data.value)) {
                    // 配列
                    var tmp = "";
                    for (let i = 0; i < data.value.length; i ++) {
                        // 配列がobjectか
                        if (typeof data.value[i] === 'object') {
                            return getData(data.value[i]);
                        } else {
                            if (tmp) {
                                tmp = tmp + "\n" + data.value[i];
                            } else {
                                tmp = data.value[i];
                            }
                        }
                    }

                    return tmp;
                } else if (typeof data.value === 'object') {
                    // object
                    for (var key1 in data.value) {
                        return getData(data.value[key1]);
                    }
                } else {
                    // 数値桁区切り
                    if (data.type === "NUMBER" && property.digit) {
                        return parseFloat(data.value).toLocaleString();
                    }

                    // // 日付 = ""DATE, 時刻 = "TIME", 日時 = "DATETIME", 作成日時 = "CREATED_TIME", 更新日時 = "UPDATED_TIME"
                    // // TODO 現在のフォーマット = 2024/04/01 (時刻はフォーマット変更なし)
                    // if (data.type === "DATE") {
                    //     return new Date(data.value).toLocaleDateString("ja-JP", {year: "numeric",month: "2-digit", day: "2-digit"});
                        
                    //     // 元の日付と時間
                    //     // var date = new Date('2024/04/09 10:15');

                    //     // // 日付と時間をまとめてフォーマット
                    //     // var formattedDateTime = date.toLocaleString('ja-JP', {
                    //     // year: 'numeric',
                    //     // month: 'long',
                    //     // day: 'numeric',
                    //     // hour: 'numeric',
                    //     // minute: 'numeric'
                    //     // });

                    //     // console.log(formattedDateTime); // Output: 2024年4月9日 10:15
                    // }

                    // if (data.type === "DATETIME" || data.type === "CREATED_TIME" || data.type === "UPDATED_TIME") {
                    //     return new Date(data.value).toLocaleDateString("ja-JP", {year: "numeric",month: "2-digit", day: "2-digit", hour: "numeric", minute: "numeric"});
                    // }

                    // 計算
                    if (data.type === "CALC") {
                        var calc = data.value;

                        // 桁区切り
                        // if (property.format === "NUMBER_DIGIT") {
                        //     calc = parseFloat(calc).toLocaleString();
                        // }
                        // 日時
                        if (property.format === "DATETIME") {
                            calc =  new Date(calc).toLocaleDateString("ja-JP", {year: "numeric",month: "2-digit", day: "2-digit", hour: "numeric", minute: "numeric"});
                        }
                        // 日付
                        if (property.format === "DATE") {
                            calc = new Date(calc).toLocaleDateString("ja-JP", {year: "numeric",month: "2-digit", day: "2-digit"});
                        }
                        
                        // // 単位記号
                        // if (property.unit) {
                        //     // 前につける
                        //     if (property.unitPosition === "BEFORE") {
                        //         calc = property.unit + calc;
                        //     } else {
                        //         calc = calc + property.unit;
                        //     }
                        // }
                        
                        return calc;
                    }

                    // value
                    return data.value;
                }
            } else if (data.hasOwnProperty('name')) {
                // code: , name: の状態
                return data.name;
            } else {
                return null;
            }
        } catch (error) {
            alert("エラーが発生しました。", data, error);
        }
    }
    
    // 1から26までの数字をExcelの列ラベルに変換する
    function getExcelColumnLabel(index) {
        let label = "";
        do {
            label = String.fromCharCode((index % 26) + 65) + label;
            index = Math.floor(index / 26) - 1;
        } while (index >= 0);
        return label;
    }

})(jQuery, kintone.$PLUGIN_ID);