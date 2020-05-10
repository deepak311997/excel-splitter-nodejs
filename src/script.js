"use strict";

// Imports for script
const path = require("path");
const fs = require("fs");
const exceljs = require("exceljs");
const xlsx = require("xlsx");
const zip = require('zip-folder');

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
const types = {
    manager: 'ManagerMailID',
    category: "Category",
};
const dateColumns = {
    manager: [4, 5],
    category: [4, 5],
}
const themeTableHeader = {
    'TableStyleLight9': '4f81bd',
    'TableStyleLight10': 'C0504D',
    'TableStyleLight11': '9BBB59',
    'TableStyleLight12': '8064A2',
    'TableStyleLight13': '4BACC6',
    'TableStyleLight14': 'F79646',
}
const defaultTheme = Object.keys(themeTableHeader)[0];
const outputPath = `${__dirname}/output/`;

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

async function formatWorksheet(ws, type) {
    for (let i = 1; i <= tableHeaders[type].length; i++) {
        ws.getColumn(i).width = 30;

        if (dateColumns[type].includes(i)) {
            ws.getColumn(i).style = { numFmt: 'dd-mmm-yy', alignment: { horizontal: 'left' } };
        } else {
            ws.getColumn(i).style = { alignment: { horizontal: 'left' } };
        }
    }

    return ws;
}

async function fileBasicData(ws, fileData, { type, theme = defaultTheme }) {
    if (type === 'manager') {
        ws.mergeCells('A1:J1');
    } else {
        ws.mergeCells('A1:H1');
    }

    ws.getRow(1).height = 30;
    const fileHeader = ws.getCell('A1');
    
    fileHeader.value = fileTitles[type]
    fileHeader.font = { size: 14, bold: true, color: { argb: 'ffffff' } };
    fileHeader.alignment = { horizontal: 'center', vertical: 'middle' };
    fileHeader.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor: { argb: themeTableHeader[theme] },
      };
    
    for (const data in fileData) {
        const { label, value } = fileData[data];

        ws.addRow([label, value]).font = {
            bold: true,
        };
    }

    ws = await formatWorksheet(ws, type);

    return ws;
}

function initialNewWorkbook(sheetName) {
    const wb = new exceljs.Workbook();
    wb.creator = "Gagan";
    wb.created = new Date();
    wb.modified = new Date();
    wb.calcProperties.fullCalcOnLoad = true;
    wb.views = [{ x: 0, y: 0, width: 10000, height: 20000, firstSheet: 0, activeTab: 1, visibility: 'visible' }];

    wb.addWorksheet(sheetName, { views: [{ showGridLines: false }] });

    return wb;
}

function addFileData(ws, rows, { type, theme = defaultTheme }) {
    ws.addTable({
        name: 'employeeTable',
        ref: 'A8',
        headerRow: true,
        style: {
            showRowStripes: true,
            theme,
        },
        columns: tableHeaders[type],
        rows: rows.map(row => Object.values(row)),
    }).commit();
}

function zipFolder() {
    return new Promise((res, rej) => {
        zip(outputPath, `${__dirname}/result.zip`, function(err) {
            if(err) {
               printError(err);
               rej(err);
            } else {
                res();
            }
        });
    });
}

async function writeFiles(uniqueCategories, fileDescription, args) {
    let successful = 0, failed = 0;
    const total = uniqueCategories.size;

    fs.mkdirSync(outputPath);

    for (const [key, values] of uniqueCategories) {
        let wb = initialNewWorkbook(key);

        // Write into excel
        fileBasicData(wb.worksheets[0], fileDescription.get(key), { type: args.type, theme: args.theme });
        addFileData(wb.worksheets[0], values, { type: args.type, theme: args.theme });

        try {
            await wb.xlsx.writeFile(`${outputPath}${key}.xlsx`);
            successful = successful + 1;
            printStatus(`${key} file generated`);
        } catch (er) {
            failed = failed + 1;
            printError(`${key} file failed to generate\n${er}`);
        };
    }
    printOutput(total, successful, failed);
    await zipFolder();
}

async function processFileData(sheet, args) {
    printStatus("File processing started...");

    const rows = xlsx.utils.sheet_to_json(sheet, { defval: ' ' });
    const uniqueCategories = new Map(), fileDescription = new Map();
    for (const row of rows) {
        let uniqueKey = row[types[args.type]].trim();

        if (args.type === 'manager') {
            uniqueKey = uniqueKey.toLowerCase();
        }

        if (!uniqueCategories.has(uniqueKey)) {
            const fileBasicContent = fileBasicDesc[args.type].reduce((acc, { key, label }) => {
                acc[key] = {
                    label,
                    value: row[key],
                }
                return acc;
            }, {});
            fileBasicContent['date'] = { value: args.date, label: 'Month & Year: ' };
            fileDescription.set(uniqueKey, fileBasicContent);

            for (const col of removeColumns[args.type]) {
                delete row[col];
            }

            uniqueCategories.set(uniqueKey, [row]);
        } else {
            for (const col of removeColumns[args.type]) {
                delete row[col];
            }

            uniqueCategories.set(uniqueKey, uniqueCategories.get(uniqueKey).concat([row]));
        }

    }

    return await writeFiles(uniqueCategories, fileDescription, args);
}

async function updateExcelHeader(sheet, type) {
    var range = xlsx.utils.decode_range(sheet['!ref']);
    var C, R = range.s.r;
    for (C = range.s.c; C <= range.e.c; ++C) {
        var cell = sheet[xlsx.utils.encode_cell({ c: C, r: R })] /* find the cell in the first row */

        var name = "UNKNOWN " + C; // <-- replace with your desired default 
        if (cell && cell.t) name = xlsx.utils.format_cell(cell);

        if (!removeColumns[type].includes(name)) {
            tableHeaders[type].push({ name })
        }
    }
}

async function readInputFile(args) {
    printStatus("Input file read successfully!!");
    const workbook = xlsx.read(args.inputfile, { type: "buffer", cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    await updateExcelHeader(sheet, args.type);
    return await processFileData(sheet, args);
}

function main(req, res) {
    const { type, output, theme, date } = req.body;
    const outDirPath = path.resolve(output);

    if (!req.files || !req.files.length) {
        res.status(400).send('Missing input file');
    } else if (!type || !type.length) {
        res.status(400).send("Missing type for split");
    } else {
        readInputFile({ type, outdir: output, inputfile: req.files[0].buffer, theme, date }).then(() => {
            const resultPath =`${__dirname}/result.zip`;
            const stream = fs.createReadStream(resultPath);

            stream.pipe(res);
            res.on('finish', () => {
                if (fs.existsSync(outputPath)){
                    fs.rmdirSync(outputPath, { recursive: true });
                    fs.unlinkSync(resultPath);
                }
            });
        }).catch(err => {
            printError(err);
            res.status(500).send("Failed to send file");
        });
    }
}

module.exports = main;