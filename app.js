const fs = require('fs');
const _ = require('lodash');
const uuidv4 = require('uuid/v4');
if(typeof require !== 'undefined') XLSX = require('xlsx');
const domain_uuid = 'a9b1257e-759a-443c-a24e-6c2692396e23';
const form_uuid = '69fbbda1-cf65-453d-bdc6-2c217328499d';
const update_flag = false; // set to true if you are just updating existing fields
let workbook = XLSX.readFile('/Users/robreeves/Documents/excel_docs/SABP_new_fields.xlsx'); // change target file name here

/**
*  MySQL parser for xlsx file
*  Take xlsx file and loop through the file to find the correct column
*  populate variables with values from the fields that you wish to loop through
*  for example tag = row[column3]
*  with the vaiables build a MySQL statement that can be exported and used in the
*  test database for both the Main_Node_Test and Test 2 databases.
*/

/**
 * Find a sheet in the workbook by name, and return an object with keys
 * `sheet` and `range`, where `range` is an object describing the valid cells
 * of the sheet, like `{min: {r: 1, c: 1}, max: {r: 5, c:5}}`.
 */

let dc = XLSX.utils.decode_cell,
      ec = (r, c) => { return XLSX.utils.encode_cell({r: r, c: c}); };

// Define table structure below:

let table = (workbook) => {

    try {
        let sheet = workbook.Sheets.Sheet1,
            range = {min: {r: 1, c: 'A'}, max: {r: 77, c: 'K'}}; // Define range of workbook here (min & max rows and columns in spreadsheet)

        if(!sheet) {
            return { sheet: null, range: null };
        }

        // find size of the sheet
        let ref = sheet['!ref'];

        if(!ref && ref.indexOf(':') === -1) {
            throw new Error("Malformed workbook - no !ref property");
        }

        range.min = dc(ref.split(':')[0]);
        range.max = dc(ref.split(':')[1]);

        return { sheet, range };
    }
    catch (err) {
        console.log(err)
    };

}

/**
 * Find the start position of a table in the given sheet. `colMap` describes
 * the table columns as an object with key prop -> column title. Returns an
 * object with keys `columns` (maps prop -> 0-indexed column number) and
 * `firstRow`, the number of the first row of the table (will be `null`) if the
 * table was not found.
 */

let headers = (params, colMap) => {

    try {
        let range = params.args.range;
        let sheet = params.args.sheet;
        let firstRow = null,
            colsToFind = _.keys(colMap).length,

            // colmap lowercase title -> prop
            colLookup = _.reduce(colMap, (m, v, k) => { m[_.isString(v)? v.toLowerCase() : v] = k; return m; }, {}),

            // colmap props -> 0-indexed column
            columns = _.reduce(colMap, (m, v, k) => { m[k] = null; return m; }, {});

        // Look for header row and extract columns
        for(let r = range.min.r; r <= range.max.r - 1; ++r) {
            let colsFound = 0;

            for(let c = range.min.c; c <= range.max.c; ++c) {
                let cell = sheet[ec(r, c)];

                if(cell && cell.v !== undefined) {
                    let prop = colLookup[cell.t === 's'? cell.v.toLowerCase() : cell.v];
                    if(prop) {
                        columns[prop] = c;
                        ++colsFound;
                    }
                }
            }

            if(colsFound === colsToFind) {
                firstRow = r + 1;
                break;
            }
        }
        return { columns, firstRow };
    }
    catch(err) {
        console.log(err)
    }
}

/**
 * Given the `cols` and `firstRow` as returned by `findTable()`, return a list
 * of objects of all table values. Continues to the end of the sheet unless
 * passed a function `stop` that takes a mapped row object as an argument and
 * returns `true` for that row.
 */

let tableLayout = (params) => {

    try {

        let data = [];
        let firstRow = params.headers.firstRow;
        let columns = params.headers.columns;
        let sheet = params.args.sheet;
        let range = params.args.range;
        let stop = params.args.range.max;

        for(let r = firstRow; r <= range.max.r; ++r) {
            let row = _.reduce(columns, (m, c, k) => {
                let cell = sheet[ec(r, c)];
                m[k] = cell? cell.v : null;
                return m;
            }, {});

            if(row == stop) {
                break;
            }

            data.push(row);
        }

        return data;
    }
    catch (err) {
        console.log(err)
    };
};

let params = {};

params.args = table(workbook);

let colMap = {};

    colMap.colA = params.args.sheet.A1.v;
    colMap.colB = params.args.sheet.B1.v;
    colMap.colC = params.args.sheet.C1.v;
    colMap.colD = params.args.sheet.D1.v;
    colMap.colE = params.args.sheet.E1.v;
    colMap.colF = params.args.sheet.F1.v;
    colMap.colG = params.args.sheet.G1.v;
    colMap.colH = params.args.sheet.H1.v;
    colMap.colI = params.args.sheet.I1.v;
    colMap.colJ = params.args.sheet.J1.v;
    colMap.colK = params.args.sheet.K1.v;

params.colMap = colMap;
params.headers = headers(params, colMap);
params.tableLayout = tableLayout(params);

function convertParamsToJson(params) {

    try {

        for (let i = 0; i < params.tableLayout.length; i++) {
            if (params.tableLayout[i].colI !== null) {
                params.tableLayout[i].colI = JSON.parse(params.tableLayout[i].colI);
            };
            
        };
        return params;
    }
    catch (err) {
        console.log(err)
    };
};

convertParamsToJson(params);

let fieldname_placeholders = {
    "tag": params.colMap.colB,
    "type": params.colMap.colC,
    "title": params.colMap.colD,
    "validators": params.colMap.colE,
    "parent": params.colMap.colF,
    "parent_value": params.colMap.colG,
    "switch": params.colMap.colH,
    "table_options": params.colMap.colI,
    "help": params.colMap.colJ,
    "parent_uuid": params.colK,
    "order" : 0
};

let db_values = {
    "db": "`recsec_patient_surrey`",
    "table": "`tbl_dictionary2`"
};

let field_names = {
    "uuid": "`dictionary_uuid`",
    "tag": "`tag`",
    "title": "`dictionary_title`",
    "type": "`type`",
    "order": "`order`",
    "lookup_key": "`lookup_key`",
    "lookup_code": "`lookup_code`",
    "lookup_value": "`lookup_value`",
    "parent_id": "`parent_id`",
    "status": "`status`",
    "parent_value": "`parent_value`",
    "val": "`val`",
    "label": "`label`",
    "index": "`index`",
    "formtitle": "`title`",
    "parent": "`parent`",
    "switch_field": "`switch_field`"
};

let patient_mysql = (params, field_names, fieldname_placeholders, db_values, update_flag) => {

    // Updates patient table

    try {
        for (let i = 0; i < params.tableLayout.length; i++) {

            const dictionary_uuid = uuidv4();

            let tag_string = params.tableLayout[i].colB;
            let title_string = params.tableLayout[i].colD;
            let type_string = params.tableLayout[i].colC;

            if (update_flag !== true) {

                let sql = `INSERT INTO ${db_values.db}.${db_values.table} (${field_names.uuid}, ${field_names.tag}, ${field_names.title}, ${field_names.type}, ${field_names.order}) VALUES ("${dictionary_uuid}", "${tag_string}", "${title_string}", "${type_string}", "${fieldname_placeholders.order}");` + "\n";

                fs.appendFileSync("/tmp/patient_new_fields_sql", sql, function(err) {
                    if(err) {
                        return console.log(err);
                    }
                    console.log("The file was saved!");
                });

            } else if (update_flag !== false) {

                let sql = `UPDATE ${db_values.db}.${db_values.table} SET ${field_names.title} = "${title_string}" WHERE (${field_names.tag} = "${tag_string}");` + "\n";

                fs.appendFileSync("/tmp/patient_updated_fields_sql", sql, function(err) {
                    if(err) {
                        return console.log(err);
                    }
                    console.log("The file was saved!");
                });
            };

            if (params.tableLayout[i].colI !== null) {
                db_table = "`tbl_lookup`";
                let optionsArray = params.tableLayout[i].colI;
                let lookup_values = {}
                    lookup_values.order = 0;
                    lookup_values.parent_uuid = params.tableLayout[i].colK;

                let lookup_delete = `DELETE FROM ${db_values.db}.${db_table} WHERE ${field_names.parent_id} = "${lookup_values.parent_uuid}";` +"\n";

                if (update_flag !== false && lookup_values.parent_uuid !== null) {
                                
                    fs.appendFileSync("/tmp/patient_updated_fields_sql", lookup_delete, function(err) {
                        if(err) {
                            return console.log(err);
                        }
                        console.log("The file was saved!");
                    });
                }

                for (let j = 0; j < optionsArray.length; j++) {

                    let optionsObject = optionsArray[j];
                
                    for (let [key, value] of Object.entries(optionsObject)) {

                        lookup_values.lookup_key = uuidv4();
                        lookup_values.lookup_code = key;
                        lookup_values.lookup_value = value;
                        lookup_values.parent_id = dictionary_uuid;
                        lookup_values.status = 1;

                        if (update_flag !== true) {

                            let lookup_sql = `INSERT INTO ${db_values.db}.${db_table} (${field_names.lookup_key}, ${field_names.lookup_code}, ${field_names.lookup_value}, ${field_names.parent_id}, ${field_names.order}, ${field_names.status}) VALUES ("${lookup_values.lookup_key}", "${lookup_values.lookup_code}", "${lookup_values.lookup_value}", "${lookup_values.parent_id}", "${lookup_values.order}", "${lookup_values.status}");` + "\n";

                            fs.appendFileSync("/tmp/patient_new_fields_sql", lookup_sql, function(err) {
                                if(err) {
                                    return console.log(err);
                                }
                                console.log("Option values saved!");
                            });

                            lookup_values.order += 10;

                        } else if (update_flag !== false && lookup_values.parent_uuid !== null) {

                            let lookup_sql = `INSERT INTO ${db_values.db}.${db_table} (${field_names.lookup_key}, ${field_names.lookup_code}, ${field_names.lookup_value}, ${field_names.parent_id}, ${field_names.order}, ${field_names.status}) VALUES ("${lookup_values.lookup_key}", "${lookup_values.lookup_code}", "${lookup_values.lookup_value}", "${lookup_values.parent_uuid}", "${lookup_values.order}", "${lookup_values.status}");` + "\n";

                            fs.appendFileSync("/tmp/patient_updated_fields_sql", lookup_sql, function(err) {
                                if(err) {
                                    return console.log(err);
                                }
                                console.log("The file was saved!");
                            });

                            lookup_values.order += 10;

                        }
                    };
                };
            };

            fieldname_placeholders.order += 10;
        };

    }

    catch (err) {
        console.log(err)
    }
};

let portal_mysql = (params, fieldname_placeholders, db_values, field_names) => {

    // Updates portal database

    try {
        for (let i = 0; i < params.tableLayout.length; i++) {
            fieldname_placeholders.order = params.tableLayout.colK
            // change database and table names here
            db_values.db = "`recsec_portal`"
            db_values.table = "`recsec_portal_formfields`"

            let spreadsheet_values = {
                "tag": params.tableLayout[i].colB,
                "title": params.tableLayout[i].colD,
                "parent": params.tableLayout[i].colF,
                "parent_value": params.tableLayout[i].colG,
                "type": params.tableLayout[i].colC,
                "order": params.tableLayout[i].colK,
                "switch": params.tableLayout[i].colH
            };

            if (update_flag !== true) {

                let portal_sql = `INSERT INTO ${db_values.db}.${db_values.table} (${field_names.tag}, ${field_names.formtitle}, ${field_names.parent}, ${field_names.parent_value}, ${field_names.type}, ${field_names.order}, ${field_names.switch_field}) VALUES ("${spreadsheet_values.tag}", "${spreadsheet_values.title}", "${spreadsheet_values.parent}", "${spreadsheet_values.parent_value}", "${spreadsheet_values.type}", "${spreadsheet_values.order}", "${spreadsheet_values.switch}");` +"\n";

                fs.appendFileSync("/tmp/portal_new_fields_sql", portal_sql, function(err) {
                    if(err) {
                        return console.log(err);
                    }
                    console.log("The file was saved!");
                });

            } else if (update_flag !== false) {

                let portal_sql = `UPDATE ${db_values.db}.${db_values.table} SET ${field_names.formtitle} = "${spreadsheet_values.title}" WHERE (${field_names.tag} = "${spreadsheet_values.tag}" AND ${field_names.parent} = "${spreadsheet_values.parent}" AND ${field_names.switch_field} = "${spreadsheet_values.switch}");` + "\n";

                fs.appendFileSync("/tmp/portal_updated_fields_sql", portal_sql, function(err) {
                    if(err) {
                        return console.log(err);
                    }
                    console.log("The file was saved!");
                });
            };

            let spreadsheet_options = params.tableLayout[i].colI;

            if(spreadsheet_options !== null) {

                db_values.table = "`recsec_portal_formoptions`";
                let index = 0;
                let formOptions = {
                    "tag": spreadsheet_values.tag
                };

                if (update_flag !== false) {

                    let formoptions_delete = `DELETE FROM ${db_values.db}.${db_values.table} WHERE ${field_names.tag} = "${formOptions.tag}";` +"\n";

                    fs.appendFileSync("/tmp/portal_updated_fields_sql", formoptions_delete, function(err) {
                        if(err) {
                            return console.log(err);
                        }
                        console.log("Portal values saved!");
                    });

                };

                for (let h = 0; h < spreadsheet_options.length; h++) {

                    let optionsObject = spreadsheet_options[h];

                    for (let [key, value] of Object.entries(optionsObject)) {

                        formOptions.val = key;
                        formOptions.label = value;

                        if (update_flag !== true) {

                            let formoptions_sql = `INSERT INTO ${db_values.db}.${db_values.table} (${field_names.tag}, ${field_names.val}, ${field_names.label}, ${field_names.index}, ${field_names.parent}) VALUES ("${formOptions.tag}", "${formOptions.val}", "${formOptions.label}", "${index}", "${spreadsheet_values.parent}");` +"\n";

                            fs.appendFileSync("/tmp/portal_new_fields_sql", formoptions_sql, function(err) {
                                if(err) {
                                    return console.log(err);
                                }
                                console.log("Portal values saved!");
                            });

                            index += 10;

                        } else if (update_flag !== false) {

                            let formoptions_sql = `INSERT INTO ${db_values.db}.${db_values.table} (${field_names.tag}, ${field_names.val}, ${field_names.label}, ${field_names.index}, ${field_names.parent}) VALUES ("${formOptions.tag}", "${formOptions.val}", "${formOptions.label}", "${index}", "${spreadsheet_values.parent}");` +"\n";

                            fs.appendFileSync("/tmp/portal_updated_fields_sql", formoptions_sql, function(err) {
                                if(err) {
                                    return console.log(err);
                                }
                                console.log("Portal values saved!");
                            });

                            index += 10;

                        }
                    };
                };
            }
        };
    }
    catch (err) {
    console.log(err)
    };
};

patient_mysql(params, field_names, fieldname_placeholders, db_values, update_flag);
portal_mysql(params, fieldname_placeholders, db_values, field_names, update_flag);

