//@ts-nocheck
// 将一个sheet转换为一个kv表的gulp插件
import through2 from 'through2';
import xlsx from 'node-xlsx';
import Vinyl from 'vinyl';
import path from 'path';
import { pinyin, customPinyin } from 'pinyin-pro';
import { pushNewLinesToCSVFile } from './kvToLocalization';

const cli = require('cli-color');

const PLUGIN_NAME = 'gulp-dotax:sheetToKV';

export interface SheetToKVOptions {
    /** 需要略过的表格的正则表达式 */
    sheetsIgnore?: string;
    /** 是否启用啰嗦模式 */
    verbose?: boolean;
    /** 是否将汉语转换为拼音 */
    chineseToPinyin?: boolean;
    /** 自定义的拼音 */
    customPinyins?: Record<string, string>;
    /** KV的缩进方式，默认为四个空格 */
    indent?: string;
    /** 是否将只有两列的表输出为简单键值对 */
    autoSimpleKV?: boolean;
    /** Key行行号，默认为2 */
    keyRowNumber?: number;
    /** KV文件的扩展名，默认为 .txt */
    kvFileExt?: string;
    /** 强制输出空格的单元格内容（如果单元格内容为此字符串，输出为 "key" "" */
    forceEmptyToken?: string;
    /** 中文转换为英文的映射列表，这些中文将会被转换为对应的英文而非拼音 */
    aliasList?: Record<string, string>;
    /**
     * 输出本地化文本到 addon.csv 文件，如果要启动，需要配置 addon.csv所在路径
     * 使用方法：
     *   将 sheet 的第二行写特定的key，例如 `#Loc{}_Lore`，{} 的内容将会被替换为第一列的主键
     **/
    addonCSVPath?: string;
    /** addon.csv输出的默认语言，默认为SChinese */
    addonCSVDefaultLang?: string;
}

function isEmptyOrNullOrUndefined(value: any) {
    return value === null || value === undefined || value === ``;
}

export function sheetToKV(options: SheetToKVOptions) {
    const {
        customPinyins = {},
        sheetsIgnore = /^\s*$/,
        verbose = false,
        forceEmptyToken = `__empty__`,
        autoSimpleKV = true,
        kvFileExt = '.txt',
        chineseToPinyin = false,
        keyRowNumber = 2,
        indent = '    ',
        aliasList = {},
        addonCSVPath = null,
        addonCSVDefaultLang = `SChinese`,
    } = options;

    customPinyin(customPinyins);

    const aliasKeys = Object.keys(aliasList)
        // 按从长到短排序，这样可以保证别名的替换不会出现问题
        .sort((a, b) => b.length - a.length);

    // 本地化token列表
    let locTokens: { [key: string]: string }[] = [];

    function convert_chinese_to_pinyin(da: string) {
        if (da === null || da.match === null) return da;

        // 如果da中包含别名，那么先将别名替换掉（可能是中文替换中文，或者中文替换成英文等等）
        aliasKeys.forEach((aliasKey) => {
            da = da.replace(aliasKey, aliasList[aliasKey]);
        });

        let s = da;
        let reg = /[\u4e00-\u9fa5]+/g;
        let match = s.match(reg);
        if (match != null) {
            match.forEach((m) => {
                s = s
                    .replace(m, pinyin(m, { toneType: 'none', type: 'array' }).join('_'))
                    .replace('ü', 'v');
            });
        }
        return s;
    }

    function deal_with_kv_value(value: string): string {
        if (/^[0-9]+.?[0-9]*$/.test(value)) {
            let number = parseFloat(value);
            // if this is not an integer, max 4 digits after dot
            if (number % 1 !== 0) {
                value = number.toFixed(4);
            }
        }

        if (value === undefined) return '';

        if (forceEmptyToken === value) return '';

        return value;
    }


    function convert_row_to_kv(row: string[], row_index: integer, note_row: string[], key_row: string[], is_array: boolean, array_index: integer): string {
        // 第一列为主键
        let main_key = row[0];
        let _desc = key_row[1]

        let new_line = '\n'

        // ignore_newline 忽略换行
        function deal_cell_note(note_str: string, indentStr: string, cell_key: string, ignore_newline: boolean): string {
            let res_str = null
            if (row_index == 0 && null != note_str) {
                res_str = `${indentStr}// ${note_str} ${new_line}${indentStr}${cell_key}`;
            } else {
                res_str = `${indentStr}${cell_key}`;
            }
            if (!ignore_newline) {
                res_str = res_str + `${new_line}`;
            }
            return res_str
        }

        function checkSpace(str: string) {
            if (typeof str == 'string' && str.trim != null && str != str.trim()) {
                console.warn(
                    cli.red(
                        `${main_key} 中的 ${str} 前后有空格，请检查！`
                    )
                );
            }
        }

        checkSpace(main_key);

        let attachWearablesBlock = false;
        let abilityValuesBlock = false;
        let varIndex = 0;
        let indentLevel = 1 + (is_array ? 1 : 0);

        let locAbilitySpecial = null;


        let listValuesBlock = new Array();
        let mapValuesBlock = new Array();

        //var length = arr.unshift("water"); 

        let end_tail = is_array ? `${indent}}` : `}`
        let row_cells =key_row.map((key, i) => {
                // 判断key前后是否有空格，如果有，那么输出一个警告
                checkSpace(key);

                // 跳过没有的key
                if (isEmptyOrNullOrUndefined(key)) return;
                let indentStr = (indent || `\t`).repeat(indentLevel);
                // 第一列为主键
                if (i === 0) {
                    indentLevel++;

                    let m_key = is_array ? array_index : main_key
                    if (_desc === '_name' || _desc === '_desc') {
                        let note_str = row[1]
                        return `${indentStr}// ${note_str}${new_line}${indentStr}"${m_key}" {${new_line}`;
                    } else {
                        return `${indentStr}"${m_key}" {${new_line}`;
                    }
                }

                // 忽略第二行备注
                if (key_row[i] === '_name' || key_row[i] === '_desc') {
                    return
                }

                let note_str = note_row[i]
                if (null == note_str || note_str == "" || note_str.trim == null || note_str.trim() == "") {
                    note_str = null
                }

                // 处理饰品的特殊键值对
                if (key === `AttachWearables[`) {
                    attachWearablesBlock = true;
                    indentLevel++;
                    indentStr = (indent || `\t`).repeat(indentLevel);
                    // return `${indentStr}"${key.replace(`[`, ``)}" {${new_line}`;
                    return deal_cell_note(note_str, indentStr, `"${key.replace(`[`, ``)}" {`, false)
                }
                // 处理饰品的特殊键值对结束
                if (attachWearablesBlock && key == ']') {
                    attachWearablesBlock = false;
                    indentLevel--;
                    indentStr = (indent || `\t`).repeat(indentLevel + 1);
                    return `${indentStr}}${new_line}`;
                }

                // 获取该单元格的值
                let cell: string = row[i];
                checkSpace(cell);

                if (
                    attachWearablesBlock &&
                    cell != `` &&
                    cell != undefined
                ) {

                    let indentStr = (indent || `\t`).repeat(indentLevel + 1);
                    // 如果输出中包含 { } 等，那么直接输出value，不加双引号
                    if (cell !== null && cell.toString().trimStart().startsWith('{')) {
                        return `${indentStr}"${key}" ${cell}${new_line}`;
                    }

                    let indentStr2 = (indent || `\t`).repeat(indentLevel + 2);
                    return `${indentStr}"Wearable${key}" ${new_line}${indentStr}{${new_line}${indentStr2}"ItemDef" "${cell}"${new_line}${indentStr}}${new_line}`;
                }

                // 处理写excel文件中的本地化文本
                if (key.includes(`#Loc`)) {
                    if (isEmptyOrNullOrUndefined(cell)) return;
                    if (cell.trim && cell.trim() === ``) return;
                    let locKey = key.replace(`#Loc`, ``).replace(`{}`, main_key);
                    // 保存对应的本地化tokens
                    locTokens.push({
                        //TODO, 将Tokens修改为 addon.csv 第一行的第一个元素？
                        KeyName: locKey,
                        [addonCSVDefaultLang]: cell,
                    });
                    return; // 不输出到kv文件
                }

                // 如果key是 #LocValues，那么则作为本地化文本暂存
                if (key == `#ValuesLoc`) {
                    if (isEmptyOrNullOrUndefined(cell)) return;
                    let values_key = '';
                    // 如果key不是数字，那么则作为key
                    if (isNaN(Number(key))) {
                        values_key = key;
                    }
                    let datas = cell.toString().split(' ');
                    if (isNaN(Number(datas[0]))) {
                        values_key = datas[0];
                        cell = cell.replace(`${datas[0]} `, '');
                    }
                    if (values_key == '') {
                        values_key = `unknown_var_${varIndex}`;
                        varIndex++;
                    }

                    if (isEmptyOrNullOrUndefined(cell)) return;
                    if (cell.trim && cell.trim() === ``) return;
                    // 暂存键值的本地化文本
                    locAbilitySpecial = cell;

                    // 如果有暂存的本地化文本，那么作为下一个遇到的键值对的本地化文本输出到 addon.csv
                    if (locAbilitySpecial != null) {
                        let locKey = `DOTA_Tooltip_ability_${main_key}_${values_key}`;
                        // 保存对应的本地化tokens
                        locTokens.push({
                            KeyName: locKey,
                            [addonCSVDefaultLang]: locAbilitySpecial,
                        });
                        locAbilitySpecial = null; // 重置本地化文本状态
                    }
                    return; // 不输出到kv文件中去
                }


                // 如果是数组形式结构
                if (key.endsWith("[")) {
                    listValuesBlock.unshift(key);
                    indentStr = (indent || `\t`).repeat(listValuesBlock.length + indentLevel - 1);
                    let tkey = key.replace(`[`, ``)
                    return deal_cell_note(note_str, indentStr, `"${tkey}" {`, false)
                }
                // 数组结束
                if (key == ']') {
                    listValuesBlock.shift()
                    indentStr = (indent || `\t`).repeat(listValuesBlock.length + indentLevel);
                    return `${indentStr}}` + `${new_line}`;
                }

                // 如果map形式结构
                if (key.endsWith("{")) {
                    listValuesBlock.unshift(key);
                    indentStr = (indent || `\t`).repeat(listValuesBlock.length + indentLevel - 1);
                    let tkey = key.replace(`{`, ``)
                    return deal_cell_note(note_str, indentStr, `"${tkey}" {`, false)
                }

                // map结束
                if (key.endsWith('}')) {
                    listValuesBlock.shift()
                    indentStr = (indent || `\t`).repeat(listValuesBlock.length + indentLevel);
                    return `${indentStr}}${new_line}`;
                }

                if ((isEmptyOrNullOrUndefined(cell)) && !/^Ability[0-9]{1,2}/.test(key)) {
                    return;
                }
                indentStr = (indent || `\t`).repeat(listValuesBlock.length + indentLevel);

                // 缩进
                if (key != "value" && key != "v") {
                    indentStr = (indent || `\t`).repeat(listValuesBlock.length + indentLevel);
                } else {
                    let idlv = Math.max(0, listValuesBlock.length + indentLevel - 4)
                    indentStr = (indent || `\t`).repeat(idlv);
                }
                const output_value = deal_with_kv_value(cell);

                // 如果输出中包含 { } 等，那么直接输出value，不加双引号
                if (
                    output_value != null &&
                    output_value.toString().trimStart().startsWith('{')
                ) {
                    return `${indentStr}"${key}" ${output_value}${new_line}`;
                }

                if (key == "key" || key == "k") {
                    if (row_index == 0 && null != note_str) {
                        return deal_cell_note(note_str, indentStr, `"${output_value}"`, true)
                    } else {
                        return `${indentStr}"${output_value}"`;
                    }
                } else if (key == "value" || key == "v") {
                    return `${indentStr} "${output_value}"${new_line}`;
                }
                return deal_cell_note(note_str, indentStr, `"${key}" "${output_value}"`, false)
            })
            .filter((row) => row != null)
            .map((s) => (chineseToPinyin ? convert_chinese_to_pinyin(s) : s))

        // 删除空的数组
        for (let i = 0; i < row_cells.length; i++) {
            let cell_str = row_cells[i]
            let next_cell_str = row_cells[i  +  1]
            if(null != cell_str  && null != next_cell_str){
                cell_str = cell_str.trim()
                next_cell_str = next_cell_str.trim()
                if (cell_str.endsWith('{') && next_cell_str == '}') {
                    row_cells.splice(i, 1);
                    row_cells.splice(i, 1);
                    i--
                    i--
                }
            }
        }
        return (
            row_cells.join('') + `${indent}${end_tail}${new_line}`
        );
    }

    let genratedFiles: string[] = [];
    function convert(this: any, file: Vinyl, enc: any, next: Function) {
        if (file.isNull()) return next(null, file);
        if (file.isStream()) return next(new Error(`${PLUGIN_NAME} Streaming not supported`));
        if (file.basename.startsWith(`~$`)) {
            console.log(`${PLUGIN_NAME} Ignore temp xlsx file ${file.basename}`);
            return next();
        }
        // ignore files that are not xlsx,xls
        if (!file.basename.endsWith(`.xlsx`) && !file.basename.endsWith(`.xls`)) {
            console.log(cli.green(`${PLUGIN_NAME} ignore non-xlsx file ${file.basename}`));
            return next();
        }

        if (file.isBuffer()) {
            console.log(`${PLUGIN_NAME} Converting ${file.path} to kv`);
            const workbook = xlsx.parse(file.contents);
            workbook.forEach((sheet) => {
                let sheet_name = sheet.name;

                if (new RegExp(sheetsIgnore).test(sheet_name)) {
                    console.log(
                        cli.red(
                            `${PLUGIN_NAME} Ignoring sheet ${sheet_name} in workbook ${file.path} 【已忽略表${sheet_name}】`
                        )
                    );
                    return;
                }

                // 如果名称中包含中文，那么弹出一个提示，说可以把中文名称的表格忽略
                if (sheet_name.match(/[\u4e00-\u9fa5]+/g)) {
                    console.log(
                        cli.yellow(
                            `${PLUGIN_NAME} Warning: ${sheet_name} 包含中文，将其转换为英文输出`
                        )
                    );
                    console.log(cli.yellow(`如果你不想输出这个表，请将其名称加入sheetsIgnore中`));
                    sheet_name = convert_chinese_to_pinyin(sheet_name);
                }

                const sheet_data = sheet.data as string[][];
                const sheet_data_length = sheet_data.length;
                if (sheet_data_length === 0) {
                    if (verbose) {
                        console.log(
                            cli.red(
                                `${PLUGIN_NAME} Ignoring empty sheet ${sheet_name} in workbook ${file.path}`
                            )
                        );
                    }
                    return;
                }

                let note_row = sheet_data[keyRowNumber - 2].map((i) => i.toString()); // 第一行为备注行
                let key_row = sheet_data[keyRowNumber - 1].map((i) => i.toString()); // 第二行为key行
                const kv_data = sheet_data.slice(keyRowNumber);
                const kv_data_length = kv_data.length;
                if (kv_data_length === 0) {
                    if (verbose) {
                        console.log(
                            cli.red(
                                `${PLUGIN_NAME} Ignoring no data sheet ${sheet_name} in workbook ${file.path}`
                            )
                        );
                    }
                    return;
                }

                let kv_data_str = '';

                let file_name = "XLSXContent"
                if (key_row.length == 2 && autoSimpleKV) {
                    const kv_data_simple = kv_data.map((row) => {
                        return `\t"${row[0]}" "${row[1]}"`;
                    });
                    kv_data_str = `${kv_data_simple.join('\n')}`;
                    file_name = sheet_name
                } else {

                    let mapMainKey = new Map();
                    kv_data.map((row, idx) => {
                        if (isEmptyOrNullOrUndefined(row[0])) return;
                        let main_key = row[0];
                        let count = mapMainKey.get(main_key)
                        if (count == null) {
                            count = 0
                        }
                        mapMainKey.set(main_key, count + 1);
                    });

                    let index = 1;
                    let current_mainkey;
                    let indentStr = (indent || `\t`).repeat(1);
                    let new_line = '\n'
                    const kv_data_complex = kv_data.map((row, row_index) => {
                        if (isEmptyOrNullOrUndefined(row[0])) return;
                        let res = ``
                        let main_key = row[0];

                        // 第一列相同key的数据轮询完毕 
                        if (current_mainkey != null && current_mainkey != main_key) {
                            res = `${indentStr}}${new_line}`;
                            current_mainkey = null;
                            index = 1
                        }

                        let count = mapMainKey.get(main_key)
                        if (count > 1) {
                            if (current_mainkey == null) {
                                res = res + `${indentStr}"${main_key}" {${new_line}`;
                            }
                            current_mainkey = main_key
                            index = index + 1
                        } else {
                            current_mainkey = null;
                            index = 1
                        }

                        res = res + convert_row_to_kv(row, row_index, note_row, key_row, count > 1, index - 1);
                        return res;
                    });
                    if (null != current_mainkey) {
                        kv_data_complex.push(`${indentStr}}${new_line}`);
                    }

                    kv_data_str = `${kv_data_complex.join('')}`;

                    if (kv_data_str.includes(`override_hero`)) {
                        file_name = "DOTAHeroes"
                    } else if (kv_data_str.includes(`AttackCapabilities`)) {
                        file_name = "DOTAUnits"
                    } else if (kv_data_str.includes(`BaseClass`) && (kv_data_str.includes(`item_lua`) || kv_data_str.includes(`item_datadriven`))) {
                        file_name = "DOTAItems"
                    } else if (kv_data_str.includes(`BaseClass`) && (kv_data_str.includes(`ability_lua`) || kv_data_str.includes(`ability_datadriven`))) {
                        file_name = "DOTAAbilities"
                    } else {
                        file_name = sheet_name
                    }
                }


                let out_put = `
// this file is auto-generated by Xavier's sheet_to_kv from
// ${file.basename} ${sheet_name}
// SourceCode: https://github.com/XavierCHN/gulp-dotax/blob/master/src/sheetToKV.ts
// Template: https://github.com/XavierCHN/x-template
"${file_name}"
{
${kv_data_str}
}
`;

                const kvBaseName = `${sheet_name}${kvFileExt}`;

                console.log(`${PLUGIN_NAME} Writing sheet content to ${kvBaseName}`);

                // if file already generated, throw an error
                let fileDirectory = file.dirname;
                let generaetdFileFullname = path.join(fileDirectory, kvBaseName);
                if (genratedFiles.includes(generaetdFileFullname)) {
                    throw new Error(`[ERROR] KVFile ${generaetdFileFullname} is duplicated!`);
                }
                genratedFiles.push(generaetdFileFullname);

                // convert all line ending from CRLF TO LF
                out_put.replace(/\r\n/g, '\n');

                const kv_file = new Vinyl({
                    base: file.base,
                    path: file.path,
                    basename: kvBaseName,
                    contents: Buffer.from(out_put),
                });

                this.push(kv_file);
            });
        }
        next();
    }

    function endStream() {
        if (addonCSVPath != null) {
            pushNewLinesToCSVFile(addonCSVPath, locTokens);
        }
        this.emit('end');
    }

    return through2.obj(convert, endStream);
}
