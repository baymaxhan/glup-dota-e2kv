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

    let new_line = '\n'
    let indent_str = indent || '\t'

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

    function checkSpace(str: string) {
        if (typeof str == 'string' && str.trim != null && str != str.trim()) {
            console.warn(
                cli.red(
                    `${main_key} 中的 ${str} 前后有空格，请检查！`
                )
            );
        }
    }

    function convert_row_to_kv(title_row: string[], note_row: string[], kv_data: string[], row_index: integer, primary_col: integer, indent: integer): string {
        // 国际化参数
        let indentLevel = indent + 2;
        // 数组阻塞标记
        let listValuesBlock = new Array();
        let mapValuesBlock = new Array();

        let verticalListBlock = new Array();

        let blockStatus = new Array();

        // const kv_data_length = kv_data.length;
        let end_row_idx = 0
        let cur_mainkey = kv_data[row_index][primary_col]
        function get_end_row() {
            if (end_row_idx > 0) { return end_row_idx }
            end_row_idx = kv_data.length;
            for (let _row = row_index + 1; _row < kv_data.length; _row++) {
                // 主键 列不为空 
                if (!isEmptyOrNullOrUndefined(kv_data[_row][primary_col])) {
                    end_row_idx = _row - 1
                }
            }
            return end_row_idx
        }

        let row_cells = new Array();

        // 根据列进行数据处理 知道处理完最后一列
        function _handle(_row, _col, next) {
            if (_col >= title_row.length) { return }
            let title = title_row[_col]
            // 空title 直接跳过
            // 以_或#开头标识要忽略的列
            if (_col == primary_col || isEmptyOrNullOrUndefined(title) || title.startsWith('_') || title.startsWith('#') || title.includes('#Loc')) {
                _handle(_row, _col + 1, true)
                return
            };

            let indentStr = (`${indent_str}`).repeat(indentLevel);
            // 集合结尾标识
            if (title == "↓]" || title == "↓}" || title == "]" || title == "}") {
                indentLevel = indentLevel - 1
                console.log(`indentLevel = ${indentLevel}`)
                indentStr = (`${indent_str}`).repeat(indentLevel);
                row_cells.push(`${indentStr}}${new_line}`);
                // 处理下一列
                _handle(_row, _col + 1, true)
                return
            }

            // 自定义注释 描述
            if (title != "value" && title != "v") {
                let note_str = note_row[_col]
                if (null == note_str || note_str == "" || note_str.trim == null || note_str.trim() == "") {
                    note_str = null
                }
                if (row_index == 0 && null != note_str) {
                    if (title.includes("[↓") || title.includes("{↓") || title.includes("[") || title.includes("{") || !isEmptyOrNullOrUndefined(kv_data[row_index][_col])) {
                        indentStr = (`${indent_str}`).repeat(indentLevel);
                        row_cells.push(`${indentStr}// ${note_str} ${new_line}`);
                    }
                }
            }

            // 竖向数组
            if (title.endsWith("[↓")) {
                indentStr = (`${indent_str}`).repeat(indentLevel);
                indentLevel = indentLevel + 1
                let s_title = title.replace(`[↓`, ``)
                row_cells.push(`${indentStr}"${s_title}"${indent_str}{${new_line}`);

                let _endrow = get_end_row()
                let _endcol = title_row.length
                for (let col_idx = _col; col_idx < title_row.length; col_idx++) {
                    let s_title = title_row[col_idx]
                    if (!s_title) continue;
                    if (s_title == "↓]") {
                        _endcol = col_idx
                        break
                    }
                }

                indentStr = (`${indent_str}`).repeat(indentLevel);
                for (let index = row_index; index <= _endrow; index++) {
                    row_cells.push(`${indentStr}"${index - row_index + 1}"${indent_str}{${new_line}`);
                    indentLevel = indentLevel + 1
                    for (let col_idx = _col + 1; col_idx < _endcol; col_idx++) {
                        indentStr = (`${indent_str}`).repeat(indentLevel);
                        if (isEmptyOrNullOrUndefined(title_row[col_idx])) {
                            continue
                        }
                        if (isEmptyOrNullOrUndefined(kv_data[index]) || isEmptyOrNullOrUndefined(kv_data[index][col_idx])) {
                            continue
                        }
                        let next_flag = title_row[col_idx].endsWith("[") || title_row[col_idx].endsWith("[↓")
                            || title_row[col_idx].endsWith("{") || title_row[col_idx].endsWith("{↓")
                        // row_cells.push(`${indentStr}"${title_row[col_idx]}" "${kv_data[index][col_idx]}"${new_line}`);
                        _handle(index, col_idx, next_flag)
                    }
                    indentLevel = indentLevel - 1
                    indentStr = (`${indent_str}`).repeat(indentLevel);
                    row_cells.push(`${indentStr}}${new_line}`);
                }

                indentLevel = indentLevel - 1
                indentStr = (`${indent_str}`).repeat(indentLevel);
                row_cells.push(`${indentStr}}${new_line}`);
                // 处理下一列
                _handle(_row, _endcol + 1, true)
                return
            }
            // 竖向map
            if (title.endsWith("{↓")) {
                indentStr = (`${indent_str}`).repeat(indentLevel);
                indentLevel = indentLevel + 1
                let s_title = title.replace(`{↓`, ``)
                // 开始 {
                row_cells.push(`${indentStr}"${s_title}"${indent_str}{${new_line}`);

                let _endrow = get_end_row()
                let _endcol = title_row.length
                for (let col_idx = _col + 1; col_idx < title_row.length; col_idx++) {
                    let s_title = title_row[col_idx]
                    if (!s_title) continue;
                    if (s_title.endsWith("[↓") || s_title.endsWith("{↓") || s_title.endsWith("[") || s_title.endsWith("{") || s_title == "↓}") {
                        _endcol = col_idx
                        break
                    }
                }
                console.log(` ${title} end ${_endcol}`)

                indentStr = (`${indent_str}`).repeat(indentLevel);
                for (let index = row_index; index <= _endrow; index++) {
                    for (let col_idx = _col + 1; col_idx < _endcol; col_idx++) {
                        indentStr = (`${indent_str}`).repeat(indentLevel);
                        if (isEmptyOrNullOrUndefined(title_row[col_idx])) {
                            continue
                        }
                        if (isEmptyOrNullOrUndefined(kv_data[index]) || isEmptyOrNullOrUndefined(kv_data[index][col_idx])) {
                            continue
                        }
                        let next_flag = title_row[col_idx].endsWith("[") || title_row[col_idx].endsWith("[↓")
                            || title_row[col_idx].endsWith("{") || title_row[col_idx].endsWith("{↓")
                        // row_cells.push(`${indentStr}"${title_row[col_idx]}" "${kv_data[index][col_idx]}"${new_line}`);
                        _handle(index, col_idx, next_flag)
                    }
                }

                indentLevel = indentLevel - 1
                indentStr = (`${indent_str}`).repeat(indentLevel);
                // 结束 }
                row_cells.push(`${indentStr}}${new_line}`);

                // 重新计算下次处理第几列
                let t_block = new Array()
                for (let col_idx = _col + 1; col_idx < title_row.length; col_idx++) {
                    let s_title = title_row[col_idx]
                    if (!s_title) continue;
                    if (s_title.endsWith("{↓")) {
                        t_block.unshift(s_title)
                    }
                    if (s_title == "↓}") {
                        if (t_block.length == 0) {
                            _endcol = col_idx
                            break;
                        }
                    }
                }

                // 处理下一列
                console.log(` next ${_endcol + 1}`)
                _handle(_row, _endcol, true)
                return
            }

            // 掉落物品集合	掉落物品		掉落数量	掉落权重	
            // items[↓		item	count	weight	↓]


            // 横向数组
            if (title.endsWith("[")) {
                indentStr = (`${indent_str}`).repeat(indentLevel);
                indentLevel = indentLevel + 1
                let s_title = title.replace(`[`, ``)
                row_cells.push(`${indentStr}"${s_title}"${indent_str}{${new_line}`);
                // 处理下一列
                _handle(_row, _col + 1, true)
                return
            }

            // 横向Map
            if (title.endsWith("{")) {
                indentStr = (`${indent_str}`).repeat(indentLevel);
                indentLevel = indentLevel + 1
                let s_title = title.replace(`{`, ``)
                row_cells.push(`${indentStr}"${s_title}"${indent_str}{${new_line}`);
                // 处理下一列
                _handle(_row, _col + 1, true)
                return
            }

            // 普通列处理
            let cell = kv_data[_row][_col]
            if ((isEmptyOrNullOrUndefined(cell)) && !/^Ability[0-9]{1,2}/.test(title)) {
                _handle(_row, _col + 1, true)
                return;
            }
            // 缩进
            if (title != "value" && title != "v") {
                indentStr = (`${indent_str}`).repeat(indentLevel);
            } else {
                let idlv = Math.max(0, indentLevel - 4)
                indentStr = (`${indent_str}`).repeat(idlv);
            }
            const output_value = deal_with_kv_value(cell);
            // 如果输出中包含 { } 等，那么直接输出value，不加双引号
            if (output_value != null && output_value.toString().trimStart().startsWith('{')) {
                row_cells.push(`${indentStr}"${title}"${indent_str}${output_value}${new_line}`);
            } else if (title == "key" || title == "k") {
                row_cells.push(`${indentStr}"${output_value}"`);
            } else if (title == "value" || title == "v") {
                row_cells.push(`${indentStr}${indent_str}"${output_value}"${new_line}`);
            } else {
                row_cells.push(`${indentStr}"${title}"${indent_str}"${output_value}"${new_line}`);
            }
            if (next) {
                _handle(_row, _col + 1, true)
            }
        }
        _handle(row_index, primary_col, true)
        // row_cells = row_cells.filter((row) => row != null)
        //     .map((s) => (chineseToPinyin ? convert_chinese_to_pinyin(s) : s))

        // 删除空的数组
        for (let i = 0; i < row_cells.length; i++) {
            let cell_str = row_cells[i]
            let next_cell_str = row_cells[i + 1]
            if (null != cell_str && null != next_cell_str) {
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
            row_cells.join('')
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
                let title_row = sheet_data[keyRowNumber - 1].map((i) => i.toString()); // 第二行为key行
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

                // 统计 有效的列
                let columns: string[] = [];
                let col_length = 0
                let row_cells = title_row.map((key, i) => {
                    // 跳过没有的key
                    if (isEmptyOrNullOrUndefined(key)) return;
                    if (key.startsWith(`#`) || key.startsWith(`_`)) return;
                    col_length = col_length + 1
                    columns.push(i)
                })

                let primary_col = columns[0]

                let file_name = "XLSXContent"
                // 只有两列的数组形式 处理成 weapon主键列
                // "weapon" {
                //     "1" "item_10001"
                //     "2" "item_10002"
                //     "3" "item_10003"
                //     "4" "item_10004"
                //     "5" "item_10005"
                //     "6" "item_10006"
                //     "7" "item_10007"
                // }
                if (col_length == 2 && title_row[primary_col].endsWith("[↓")) {
                    // "SkillDropPkt1" {
                    let value_col = columns[1]

                    let indentStr = (`${indent_str}`).repeat(1);
                    let str_columns: string[] = [];
                    let arr_index = 1
                    for (let index = 0; index < kv_data_length; index++) {
                        let row = kv_data[index]
                        // 主键 列不为空 
                        let primary_empty = isEmptyOrNullOrUndefined(row[primary_col])
                        if (!primary_empty && index != 0) {
                            str_columns.push(`${indent_str}}`)
                            arr_index = 1
                        }
                        if (!primary_empty) {
                            str_columns.push(`${indent_str}"${row[primary_col]}" {`)
                        }
                        if (!isEmptyOrNullOrUndefined(row[value_col])) {
                            str_columns.push(`${indent_str}${indentStr}"${arr_index}"${indent_str}"${row[value_col]}"`)
                            arr_index = arr_index + 1
                        }
                    }
                    str_columns.push(`${indent_str}}`)
                    kv_data_str = `${str_columns.join('\n')}`;
                    file_name = sheet_name

                } else if (col_length == 2 && autoSimpleKV) {
                    let value_col = columns[1]
                    const kv_data_simple = kv_data.map((row) => {
                        return `\t"${row[primary_col]}" "${row[value_col]}"`;
                    });
                    kv_data_str = `${kv_data_simple.join('\n')}`;
                    file_name = sheet_name
                } else {
                    // 国际化文本处理
                    title_row.forEach((title, _col) => {
                        // 处理写excel文件中的本地化文本
                        let is_loc = title.includes(`#Loc`)
                        let is_locvalues = title == `#LocValues`
                        if (is_loc || is_locvalues) {
                            let ability_name
                            kv_data.forEach((row, _row) => {
                                let p_cell = row[primary_col]
                                if (!isEmptyOrNullOrUndefined(p_cell)) {
                                    ability_name = p_cell
                                };

                                let loc_text = row[_col]
                                if (isEmptyOrNullOrUndefined(loc_text)) return;
                                if (loc_text.trim && loc_text.trim() === ``) return;
                                if (!is_locvalues) {
                                    let locKey = title.replace(`#Loc`, ``).replace(`{}`, ability_name);
                                    // 保存对应的本地化tokens
                                    locTokens.push({
                                        //TODO, 将Tokens修改为 addon.csv 第一行的第一个元素？
                                        KeyName: locKey,
                                        [addonCSVDefaultLang]: loc_text,
                                    });
                                } else {
                                    let loc_text = row[_col]
                                    if (isEmptyOrNullOrUndefined(loc_text)) return;
                                    if (isEmptyOrNullOrUndefined(title)) return;
                                    let s_title = title_row[_col - 1]
                                    let ability_value
                                    if (s_title == "key" || s_title == "k") {
                                        ability_value = row[_col - 1]
                                        if (isEmptyOrNullOrUndefined(ability_value)) return;
                                        if (ability_value.trim && ability_value.trim() === ``) return;
                                    } else {
                                        ability_value = title_row[_col - 1]
                                    }

                                    let locKey = `DOTA_Tooltip_ability_${ability_name}_${ability_value}`;
                                    locTokens.push({
                                        KeyName: locKey,
                                        [addonCSVDefaultLang]: loc_text,
                                    });
                                }
                            })
                        }
                    })


                    // 如果主键 是一个数组
                    let indentStr = (`${indent_str}`).repeat(1);
                    let arrayIndex = 0
                    let is_array = title_row[primary_col].endsWith("[↓")
                    let primary_key
                    const kv_data_complex = kv_data.map((row, row_index) => {
                        let pk = row[primary_col]
                        if (isEmptyOrNullOrUndefined(pk)) return;
                        let res = ``
                        if (pk != primary_key) {
                            primary_key = pk
                            arrayIndex = 0

                            let note_col = -1
                            title_row.forEach(function (title, idx) {
                                if (title == "#name" || title == "_name") { note_col = idx }
                            });

                            // 结束}
                            if (row_index != 0) {
                                res = res + `${indentStr}}${new_line}`;
                            }
                            // title描述
                            if (note_col != -1) {
                                res = res + `${indentStr}// ${row[note_col]}${new_line}`;
                            }
                            res = res + `${indentStr}"${primary_key}" {${new_line}`;
                        }

                        if (is_array) {
                            arrayIndex = arrayIndex + 1
                            let indentStr2 = (`${indent_str}`).repeat(2);
                            res = res + `${indentStr2}"${arrayIndex}" {${new_line}`;
                            res = res + convert_row_to_kv(title_row, note_row, kv_data, row_index, primary_col, 1);
                            res = res + `${indentStr2}}${new_line}`;
                        } else {
                            res = res + convert_row_to_kv(title_row, note_row, kv_data, row_index, primary_col, 0);
                        }
                        return res;
                    });
                    kv_data_complex.push(`${indentStr}}${new_line}`)
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
