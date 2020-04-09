"use strict";

// Imports for script
var path = require("path");
var fs = require("fs");
var xlsx = require("xlsx");

// Excel headers
const headersValues = ['First Name', 'Last Name', 'Employee ID', 'BADGE ACTIVE DATE', 'BADGE DEACTIVE DATE', 'BADGE No', 'Location', 'CATEGORY', 'MANAGER', 'Location_1', 'Site'];
const headersLabels = ['First Name', 'Last Name', 'Employee ID', 'Badge Active Date', 'Badge Deactive Date', 'Badge No', 'Location', 'Category', 'Manager', 'Department Code', 'Site'];
const fileTitle = "Interal Secure Area Manager Report";
const fileBasicDesc = [
    { key: "Location_1", label: "Location: ", labelId: 'A2', valueId: 'B2' },
    { key: "Site", label: "Site Name: ",  labelId: 'A3', valueId: 'B3'},
    // { key: "date", label: "Month & Year: "},
    { key: "CATEGORY", label: "Internal Secure Area Name: ",  labelId: 'A5', valueId: 'B5'},
    { key: "MANAGER", label: "Internal Secure Area Manager Name: ",  labelId: 'A6', valueId: 'B6'}
];

// Script arguments
const args = require("minimist")(process.argv.slice(2), {
    boolean: "help",
    string: ["inputfile", "outdir", "date"],
    alias: {
        i: "inputFile",
        o: "outdir",
        d: "date",
    }
})
const basePath = process.env.basePath || __dirname;

// Script help
function help() {
    console.log("Excel splitter usage");
    console.log("   --inputfile or --i : Input file path with filename");
    console.log("   --outdir or --o : Output directory path to save output files");
    console.log("   --date or --d : Date of generation");
}

function printError(error) {
    console.log(`${new Date().toLocaleString()} Error: ${error}`);
}

function printStatus(message) {
    console.log(`${new Date().toLocaleString()} Status: ${message}`);
}

function printOutput(total, successful) {
    printStatus("Files successfully generated !!");
    printStatus("Report");
    printStatus("***********************************************************************************");
    printStatus("");
    printStatus(`Total Files: ${total}, Successful: ${successful}, Failed: ${total - successful}`);
    printStatus("");
    printStatus("***********************************************************************************");
}

function initialNewWorkbook(fileData) {
    let fileDesc = {
        A1: {
            w: 'Interal Secure Area Manager Report ',
            t: 's',
        },
    };
    for (const data in fileData) {
        const { label, labelId, value, valueId } = fileData[data];

        fileDesc[labelId] = {
            w: label,
            t: 's',
        }
        fileDesc[valueId] = {
            w: value,
            t: 's',
        }
    }

    return fileDesc;
}

function writeFiles(uniqueCategories, fileDescription, fileBluePrint, outputDir) {
    let successful = 0;
    const total = uniqueCategories.size;
    for (const [key, value] of uniqueCategories) {
        let wb = xlsx.utils.book_new();
        wb = { ...wb, ...fileBluePrint };
        // wb = initialNewWorkbook(wb, fileDescription.get(key));
        const ws = xlsx.utils.json_to_sheet(value);

        wb.SheetNames = [key];
        wb.Sheets[key] = ws;

        try {
            xlsx.writeFile(wb, `${outputDir}/${key}.xlsx`);
            successful++;
            printStatus(`${key} file generated`);
        } catch (er) {
            printError(`${key} file failed to generate\n${er}`);
        }
    }
    printOutput(total, successful);
}

function processFileData(rows, fileBluePrint, outputDir) {
    printStatus("File processing started...");

    const uniqueCategories = new Map(), fileDescription = new Map();
    for (const row of rows) {
        const uniqueKey = row['CATEGORY'];
        
        if (!uniqueCategories.has(uniqueKey)) {
            uniqueCategories.set(uniqueKey, [row]);
            fileDescription.set(uniqueKey, fileBasicDesc.reduce((acc, { key, label, labelId, valueId }) => {
                acc[key] = {
                    label,
                    labelId,
                    value: row[key],
                    valueId,
                }
                return acc;
            }, {}));
        } else {
            uniqueCategories.set(uniqueKey, uniqueCategories.get(uniqueKey).concat([row]));
        }
    }

    writeFiles(uniqueCategories, fileDescription, fileBluePrint, outputDir);
}

function readInputFile(inputFilePath, outputDir) {
    fs.readFile(inputFilePath, (err, content) => {
        if (err) {
            printError(`Failed to read the input file\n${err.toString()}`);
            return;
        } else {
            printStatus("Input file read successfully!!");
            const workbook = xlsx.read(content, { type:"buffer", cellStyles: true, cellText: false, raw: true });
            const { Sheets: { Sheet1 }, opts, Themes } = workbook;

            processFileData(xlsx.utils.sheet_to_json(Sheet1), {
                opts,
                Themes,
            }, outputDir);
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
    } else {
        readInputFile(inputFilePath, outDirPath);
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
    } else if (!args.d || !args.d.length) {
        printError("Missing date");
        return;
    } else {
        validateArguments();
    }
}
main();