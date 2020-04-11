"use strict";

// Imports for script
const path = require("path");
const fs = require("fs");
const exceljs = require("exceljs");
const xlsx = require("xlsx");

// Excel headers
let tableHeaders = { category: [], manager: [] };
const fileTitles = {
    category: "Internal Secure Area Manager Report",
    manager: "NRE Monthly Validation Report",
};
const fileBasicDesc = {
    category: [
        { key: "Location", label: "Location: " },
        { key: "Site", label: "Site Name: " },
        { key: "Category", label: "Internal Secure Area Name: " },
        { key: "Manager", label: "Internal Secure Area Manager Name: " }
    ],
    manager: [
        { key: "Manager Name", label: "Manager Name: " },
        { key: "ManagerMailID", label: "Manager email address: " },
    ],
};
const removeColumns = {
    category: ["Location", "Site", "Manager"],
    manager: ["Manager Name"],
};
const sheetNames = {
    manager: "Manager Name",
    category: "Category",
}
const types = {
    manager: 'ManagerMailID',
    category: "Category",
};
const dateColumns = {
    manager: [4, 5],
    category: [4, 5],
}

// Script arguments
let args = require("minimist")(process.argv.slice(2), {
    boolean: "help",
    string: ["inputfile", "outdir", "date", "type"],
    alias: {
        i: "inputFile",
        o: "outdir",
        d: "date",
        t: "type",
    }
})
const basePath = process.env.basePath || __dirname;

// Script help
function help() {
    console.log("Excel splitter usage");
    console.log("   --inputfile or --i : Input file path with filename");
    console.log("   --outdir    or --o : Output directory path to save output files");
    console.log("   --date      or --d : Date of generation");
    console.log("   --type      or --t : Column type for splitting (manager) / (category)");
}

function printError(error) {
    console.log(`${new Date().toLocaleString()} Error: ${error}`);
}

function printStatus(message) {
    console.log(`${new Date().toLocaleString()} Status: ${message}`);
}

function printOutput(total, successful, failed) {
    printStatus("Files successfully generated !!");
    printStatus("Report");
    printStatus("***********************************************************************************");
    printStatus("");
    printStatus(`Total Files: ${total}, Successful: ${successful}, Failed: ${failed}`);
    printStatus("");
    printStatus("***********************************************************************************");
}

function formatWorksheet(ws) {
    for (let i = 1; i <= tableHeaders[args.t].length; i++) {
        ws.getColumn(i).width = 30;

        if (dateColumns[args.t].includes(i)) {
            ws.getColumn(i).style = { numFmt: 'dd-mmm-yy', alignment: { horizontal: 'left' } };
        } else {
            ws.getColumn(i).style = { alignment: { horizontal: 'left' } };
        }
    }

    return ws;
}

function fileBasicData(ws, fileData) {
    if (args.t === 'manager') {
        ws.mergeCells('A1:J1');
    } else {
        ws.mergeCells('A1:H1');
    }

    ws.getRow(1).height = 30;
    const fileHeader = ws.getCell('A1');
    
    fileHeader.value = fileTitles[args.t]
    fileHeader.font = { size: 14, bold: true, color: { argb: 'ffffff' } };
    fileHeader.alignment = { horizontal: 'center', vertical: 'middle' };
    fileHeader.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor: { argb: '4f81bd' },
      };
    
    for (const data in fileData) {
        const { label, value } = fileData[data];

        ws.addRow([label, value]).font = {
            bold: true,
        };
    }

    ws = formatWorksheet(ws);

    return ws;
}

function initialNewWorkbook(sheetName) {
    const wb = new exceljs.Workbook();
    wb.creator = "Gagan";
    wb.created = new Date();
    wb.modified = new Date();
    wb.calcProperties.fullCalcOnLoad = true;
    wb.views = [{ x: 0, y: 0, width: 10000, height: 20000, firstSheet: 0, activeTab: 1, visibility: 'visible' }];

    const ws = wb.addWorksheet(sheetName, { views: [{ showGridLines: false }] });

    return { wb, ws };
}

function addFileData(ws, rows) {
    ws.addTable({
        name: 'employeeTable',
        ref: 'A8',
        headerRow: true,
        style: {
            showRowStripes: true,
        },
        columns: tableHeaders[args.t],
        rows: rows.map(row => Object.values(row)),
    }).commit();
}

async function writeFiles(uniqueCategories, fileDescription) {
    let successful = 0, failed = 0;
    const total = uniqueCategories.size;
    for (const [key, values] of uniqueCategories) {
        let { wb, ws } = initialNewWorkbook(key);

        // Write into excel
        ws = fileBasicData(ws, fileDescription.get(key));
        ws = addFileData(ws, values);

        try {
            await wb.xlsx.writeFile(`${args.o}/${key}.xlsx`);
            successful = successful + 1;
            printStatus(`${key} file generated`);
        } catch (er) {
            failed = failed + 1;
            printError(`${key} file failed to generate\n${er}`);
        };
    }
    printOutput(total, successful, failed);
}

function processFileData(sheet) {
    printStatus("File processing started...");

    const rows = xlsx.utils.sheet_to_json(sheet, { defval: ' ' });
    const uniqueCategories = new Map(), fileDescription = new Map();
    for (const row of rows) {
        const uniqueKey = row[types[args.t]].trim().toLowerCase();

        if (!uniqueCategories.has(uniqueKey)) {
            const fileBasicContent = fileBasicDesc[args.t].reduce((acc, { key, label }) => {
                acc[key] = {
                    label,
                    value: row[key],
                }
                return acc;
            }, {});
            fileBasicContent['date'] = { value: args.date, label: 'Month & Year: ' };
            fileDescription.set(uniqueKey, fileBasicContent);

            for (const col of removeColumns[args.t]) {
                delete row[col];
            }

            uniqueCategories.set(uniqueKey, [row]);
        } else {
            for (const col of removeColumns[args.t]) {
                delete row[col];
            }

            uniqueCategories.set(uniqueKey, uniqueCategories.get(uniqueKey).concat([row]));
        }

    }

    writeFiles(uniqueCategories, fileDescription);
}

function updateExcelHeader(sheet) {
    var range = xlsx.utils.decode_range(sheet['!ref']);
    var C, R = range.s.r;
    for (C = range.s.c; C <= range.e.c; ++C) {
        var cell = sheet[xlsx.utils.encode_cell({ c: C, r: R })] /* find the cell in the first row */

        var name = "UNKNOWN " + C; // <-- replace with your desired default 
        if (cell && cell.t) name = xlsx.utils.format_cell(cell);

        if (!removeColumns[args.t].includes(name)) {
            tableHeaders[args.t].push({ name })
        }
    }
}

function readInputFile() {
    fs.readFile(args.i, (err, content) => {
        if (err) {
            printError(`Failed to read the input file\n${err.toString()}`);
            return;
        } else {
            printStatus("Input file read successfully!!");
            const workbook = xlsx.read(content, { type: "buffer", cellDates: true });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];

            updateExcelHeader(sheet);
            processFileData(sheet);
        }
    });
}

function validateArguments() {
    const inputFilePath = path.resolve(args.i);
    const outDirPath = path.resolve(args.o);

    if (!fs.existsSync(inputFilePath)) {
        printError("Invalid input file path");
        return;
    } else if (!fs.existsSync(outDirPath)) {
        printError("Invalid output directory");
        return;
    } else if (!Object.keys(types).includes(args.t)) {
        printError("Invalid type. Possible types (manager) / (category)");
        return;
    } else {
        args = { ...args, i: inputFilePath, o: outDirPath, t: args.t.toLowerCase() };
        readInputFile();
    }
}

function main() {
    if (args.help) {
        help();
        return;
    } else if (!args.i || !args.i.length) {
        printError("Missing input file path");
        return;
    } else if (!args.o || !args.o.length) {
        printError("Missing output directory path");
        return;
    } else if (!args.t || !args.t.length) {
        printError("Missing type for split");
        return;
    } else {
        validateArguments();
    }
}
main();