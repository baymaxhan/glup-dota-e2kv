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

function isBlank(value: any) {
    return value === null || value === undefined || value === ``;
}

function isNotBlank(value: any) {
    return !isBlank(value);
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
                let name_col
                let row_cells = title_row.map((key, i) => {
                    // 跳过没有的key
                    if (isEmptyOrNullOrUndefined(key)) return;
                    if (key == "#name") {
                        name_col = i
                    }
                    if (key.startsWith(`#`) || key.startsWith(`_`)) return;
                    col_length = col_length + 1
                    columns.push(i)
                })

                let primary_col = columns[0]

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
                // if (col_length == 2 && title_row[primary_col].endsWith("[↓")) {
                //     // "SkillDropPkt1" {
                //     let value_col = columns[1]

                //     let indentStr = (`${indent_str}`).repeat(1);
                //     let str_columns: string[] = [];
                //     let arr_index = 1
                //     for (let index = 0; index < kv_data_length; index++) {
                //         let row = kv_data[index]
                //         // 主键 列不为空 
                //         let primary_empty = isEmptyOrNullOrUndefined(row[primary_col])
                //         if (!primary_empty && index != 0) {
                //             str_columns.push(`${indent_str}}`)
                //             arr_index = 1
                //         }
                //         if (!primary_empty) {
                //             str_columns.push(`${indent_str}"${row[primary_col]}" {`)
                //         }
                //         if (!isEmptyOrNullOrUndefined(row[value_col])) {
                //             str_columns.push(`${indent_str}${indentStr}"${arr_index}"${indent_str}"${row[value_col]}"`)
                //             arr_index = arr_index + 1
                //         }
                //     }
                //     str_columns.push(`${indent_str}}`)
                //     kv_data_str = `${str_columns.join('\n')}`;

                // } else 
                if (col_length == 2 && autoSimpleKV) {
                    let value_col = columns[1]
                    const kv_data_simple = kv_data.map((row) => {
                        return `\t"${row[primary_col]}" "${row[value_col]}"`;
                    });
                    kv_data_str = `"${sheet_name}"\n{\n${kv_data_simple.join('\n')}\n}`;
                } else {
                    let kv_data_complex = new Array()
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


                    function find_end_col(start_col, end_col, start_mark, end_mark) {
                        let blockList = new Array()
                        for (let _col = start_col + 1; _col <= end_col; _col++) {
                            let title = title_row[_col]
                            if (isBlank(title)) continue;
                            title = title.trim()
                            if (title.endsWith(start_mark)) {
                                blockList.unshift(title)
                            }
                            if (title == end_mark && blockList.length == 0) {
                                return _col
                            }
                            if (title == end_mark) {
                                blockList.shift()
                            }
                        }
                        return end_col
                    }

                    let indentLevel = 0
                    let indentStr = (`${indent_str}`).repeat(indentLevel);
                    function convert_row_to_kv2(start_row: integer, end_row, start_col: integer, end_col): string {
                        // title_row: string[], note_row: string[],
                        for (let _col = start_col; _col <= end_col; _col++) {
                            let title = title_row[_col]
                            if (isBlank(title)) continue;
                            title = title.trim()

                            if (title.startsWith("#")) continue;
                            if (title.endsWith("↓]")) continue;
                            if (title.endsWith("↓}")) continue;
                            if (title.endsWith("]")) continue;
                            if (title.endsWith("}")) continue;

                            // 向下遍历数组
                            if (title.endsWith("[↓")) {
                                indentStr = (`${indent_str}`).repeat(indentLevel);
                                let s_title = title.replace(`[↓`, ``)

                                let isBlankTitle = isBlank(s_title)
                                if(!isBlankTitle){
                                    if(_col == 0){
                                        kv_data_complex.push(`${indentStr}"${s_title}"\n`);
                                        kv_data_complex.push(`${indentStr}{\n`);
                                    }else{
                                        kv_data_complex.push(`${indentStr}"${s_title}" {\n`);
                                    }
                                    // 子元素缩进 + 1
                                    indentLevel = indentLevel + 1
                                }

                                // 子元素下标
                                // 寻找当前数组的结束列 
                                let cur_end_col = find_end_col(_col, end_col, "[↓", "↓]")

                                let hasIndex = false
                                for (let _row = start_row; _row <= end_row; _row++) {
                                    let arrayIndex = kv_data[_row][_col]
                                    // 如果不为空 变数一组数据开始
                                    if (isNotBlank(arrayIndex)) {
                                        hasIndex = true
                                        break
                                    }
                                }

                                // 数组有没有索引列， 没有则每一行作为数组的一组数据
                                if(hasIndex){
                                    let child_indent = (`${indent_str}`).repeat(indentLevel);
                                    let cur_start_row = start_row
                                    for (let _row = start_row; _row <= end_row; _row++) {
                                        let arrayIndex = kv_data[_row][_col]
                                        // 如果不为空 变数一组数据开始
                                        if (isNotBlank(arrayIndex)) {
                                            if (_row != cur_start_row) {
                                                // 开始 {
                                                if (primary_col == _col && name_col) {
                                                    // #name
                                                    let note_name = kv_data[cur_start_row][name_col]
                                                    if (isNotBlank(note_name)) {
                                                        kv_data_complex.push(`${indent_str.repeat(1)}// ${note_name}\n`);
                                                    }
                                                }
                                                kv_data_complex.push(`${child_indent}"${kv_data[cur_start_row][_col]}" {\n`);
                                                indentLevel = indentLevel + 1
                                                convert_row_to_kv2(cur_start_row, _row - 1, _col + 1, cur_end_col - 1)
                                                indentLevel = indentLevel - 1
                                                kv_data_complex.push(`${child_indent}}\n`);
                                            }
                                            cur_start_row = _row
                                        }
                                    }
                                    // 最后一行处理
                                    // 开始 {
                                    if (primary_col == _col && name_col) {
                                        // #name
                                        let note_name = kv_data[cur_start_row][name_col]
                                        if (isNotBlank(note_name)) {
                                            kv_data_complex.push(`${indent_str.repeat(1)}// ${note_name}\n`);
                                        }
                                    }
                                    kv_data_complex.push(`${child_indent}"${kv_data[cur_start_row][_col]}" {\n`);
                                    indentLevel = indentLevel + 1
                                    convert_row_to_kv2(cur_start_row, end_row, _col + 1, cur_end_col - 1)
                                    indentLevel = indentLevel - 1
                                    kv_data_complex.push(`${child_indent}}\n`);
                                }else{
                                    let child_indent = (`${indent_str}`).repeat(indentLevel);
                                    let arr_index = 1
                                    for (let _row = start_row; _row <= end_row; _row++) {
                                        kv_data_complex.push(`${child_indent}"${arr_index}" {\n`);
                                        indentLevel = indentLevel + 1
                                        // 处理中间数据
                                        convert_row_to_kv2(_row, _row, _col + 1, cur_end_col - 1)
                                        indentLevel = indentLevel - 1
                                        kv_data_complex.push(`${child_indent}}\n`);
                                        arr_index = arr_index + 1
                                    }
                                }

                                // 将当前列赋值为结束列，下次遍历结束列的下一列
                                _col = cur_end_col
                                if(!isBlankTitle){
                                    indentLevel = indentLevel - 1
                                    indentStr = (`${indent_str}`).repeat(indentLevel);
                                    kv_data_complex.push(`${indentStr}}\n`);
                                }
                            } else if (title.endsWith("{↓")) {
                                // 向下的map
                                indentStr = (`${indent_str}`).repeat(indentLevel);
                                let s_title = title.replace(`{↓`, ``)
                                let custom_title = kv_data[start_row][_col]
                                if (!isBlank(custom_title)) {
                                    s_title = custom_title
                                }
                                // 开始 {
                                if(isBlank(s_title)){
                                    kv_data_complex.push(`${indentStr} {\n`);
                                }else{
                                    kv_data_complex.push(`${indentStr}"${s_title}" {\n`);
                                }
                                // 子元素缩进 + 1
                                indentLevel = indentLevel + 1

                                // 寻找当前map的结束列 
                                let cur_end_col = find_end_col(_col, end_col, "{↓", "↓}")
                                // 将当前列赋值为结束列，下次遍历结束列的下一列
                                // 处理map中间数据
                                convert_row_to_kv2(start_row, end_row, _col + 1, cur_end_col - 1)
                                // map结束
                                _col = cur_end_col
                                indentLevel = indentLevel - 1
                                indentStr = (`${indent_str}`).repeat(indentLevel);
                                kv_data_complex.push(`${indentStr}}\n`);

                            } else if (title.endsWith("[")) {
                                // 横向的数组
                                indentStr = (`${indent_str}`).repeat(indentLevel);
                                let s_title = title.replace(`[`, ``)
                                // 开始 {
                                kv_data_complex.push(`${indentStr}"${s_title}" {\n`);
                                // 子元素缩进 + 1
                                indentLevel = indentLevel + 1

                                // 子元素下标
                                // 寻找当前数组的结束列 
                                let cur_end_col = find_end_col(_col, end_col, "[", "]")

                                // console.log(`${title} find ${_col}--> ${end_col} end col = ${cur_end_col}`)
                                // 处理横向数组的每一项
                                convert_row_to_kv2(start_row, end_row, _col + 1, cur_end_col - 1)
                                // 数组结束
                                _col = cur_end_col
                                indentLevel = indentLevel - 1
                                indentStr = (`${indent_str}`).repeat(indentLevel);
                                kv_data_complex.push(`${indentStr}}\n`);

                            } else if (title.endsWith("{")) {
                                // 横向的Map
                                indentStr = (`${indent_str}`).repeat(indentLevel);
                                let s_title = title.replace(`{`, ``)
                                let custom_title = kv_data[start_row][_col]
                                if (!isBlank(custom_title)) {
                                    s_title = custom_title
                                }
                                // 开始 {
                                if(isBlank(s_title)){
                                    kv_data_complex.push(`${indentStr} {\n`);
                                }else{
                                    kv_data_complex.push(`${indentStr}"${s_title}" {\n`);
                                }
                                indentLevel = indentLevel + 1

                                // 子元素下标
                                // 寻找当前数组的结束列 
                                let cur_end_col = find_end_col(_col, end_col, "{", "}")

                                // 处理横向数组的每一项
                                convert_row_to_kv2(start_row, end_row, _col + 1, cur_end_col - 1)
                                // 数组结束
                                indentLevel = indentLevel - 1
                                indentStr = (`${indent_str}`).repeat(indentLevel);
                                kv_data_complex.push(`${indentStr}}\n`);
                                _col = cur_end_col
                            } else {
                                if (title_row[_col] == "v" || title_row[_col] == "value") continue;

                                let isKey= title == "k" || title == "key"
                                let value_col
                                if (isKey) {
                                    for (let _c = _col + 1; _c <= end_col; _c++) {
                                        if (isBlank(title_row[_c])) continue;
                                        if (title_row[_c].trim() == "v" || title_row[_c].trim() == "value") {
                                            value_col = _c
                                            break
                                        }
                                    }
                                }

                                indentStr = (`${indent_str}`).repeat(indentLevel);
                                for (let _row = start_row; _row <= end_row; _row++) {
                                    let cell = kv_data[_row][_col]

                                    if (/^Ability[0-9]{1,2}/.test(title)) {
                                        if(isBlank(cell)){
                                            cell = ""
                                        }
                                        kv_data_complex.push(`${indentStr}"${title}" "${cell}"\n`);
                                        continue;
                                    }
                                    if (isBlank(cell)) continue;
                                    if (cell.toString().trimStart().startsWith('{')) {
                                        kv_data_complex.push(`${indentStr}"${title}" ${cell}\n`);
                                    } else if (isKey) {
                                        if(isBlank(value_col)) {
                                            console.warn(
                                                cli.red(
                                                    `sheet ${sheet_name} 第 ${_row} 行 ${_col}列 ${cell} 对应的数值列没有找到，请检查！`
                                                )
                                            );
                                            continue
                                        }
                                        let v_cell = kv_data[_row][value_col]
                                        if (isBlank(v_cell)) continue;
                                        kv_data_complex.push(`${indentStr}"${cell}" "${v_cell}"\n`);
                                    } else {
                                        kv_data_complex.push(`${indentStr}"${title}" "${cell}"\n`);

                                    }
                                }

                                if (isKey) {
                                    if(!isBlank(value_col)) {
                                        _col = value_col
                                    }
                                }
                            }

                        }
                    }

                    convert_row_to_kv2(0, kv_data.length - 1, 0, title_row.length - 1)
                    // 删除空的数组
                    for (let i = 0; i < kv_data_complex.length; i++) {
                        let cell_str = kv_data_complex[i]
                        let next_cell_str = kv_data_complex[i + 1]
                        if (null != cell_str && null != next_cell_str) {
                            cell_str = cell_str.trim()
                            next_cell_str = next_cell_str.trim()
                            if (cell_str.endsWith('{') && next_cell_str == '}') {
                                kv_data_complex.splice(i, 1);
                                kv_data_complex.splice(i, 1);
                                i--
                                i--
                            }
                        }
                    }
                    kv_data_str = `${kv_data_complex.join('')}`;
                }

                let out_put = `
// this file is auto-generated by Xavier's sheet_to_kv from
// ${file.basename} ${sheet_name}
// SourceCode: https://github.com/XavierCHN/gulp-dotax/blob/master/src/sheetToKV.ts
// Template: https://github.com/XavierCHN/x-template
${kv_data_str}
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
