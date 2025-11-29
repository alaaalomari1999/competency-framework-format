const fs = require('fs');
const path = require('path');
const { parse } = require('csv-parse/sync');
const XLSX = require('xlsx');
const readline = require('readline');

// --- Configuration ---

const TARGET_COLUMNS = [
    "Parent ID number", "ID number", "Short name", "Description",
    "Description format", "Scale values", "Scale configuration",
    "Rule type (optional)", "Rule outcome (optional)", "Rule config (optional)",
    "Cross-referenced competency ID numbers", "Exported ID (optional)",
    "Is framework", "Taxonomy"
];

const STANDARD_VALUES = {
    "Description format": 1,
    "Scale values": "Not yet competent,Competent",
    "Scale configuration": "[{\"scaleid\":\"2\"},{\"id\":2,\"scaledefault\":1,\"proficient\":1}]",
    "Taxonomy": "competency,competency,competency,competency,competency"
};

const ROW_STANDARD_VALUES = {
    "Description format": 1,
    "Scale values": "",
    "Scale configuration": "",
    "Taxonomy": "competency,competency,competency,competency,competency",
};

// --- Helper Functions ---

/**
 * Generates a short code from a name.
 * Logic:
 * 1. If name matches pattern like "K1", "S1" (short with digits), keep it.
 * 2. Else, take first letter of each word (uppercase).
 */
function generateCode(name) {
    if (!name) return '';
    name = name.trim();

    // Check if it's a code like K1, S1, C1, CO-K
    if (/^[A-Za-z]+\d+$/.test(name)) {
        return name;
    }

    // Otherwise, acronym
    const words = name.split(/\s+/);
    let code = '';
    for (const w of words) {
        const cleanW = w.replace(/[^a-zA-Z0-9]/g, '');
        if (cleanW) {
            code += cleanW[0].toUpperCase();
        }
    }
    return code || name;
}

/**
 * Parses a CSV file using csv-parse library.
 */
function parseCsvWithLibrary(csvContent) {
    try {
        const records = parse(csvContent, {
            columns: true,
            skip_records: 1,
            skip_empty_lines: true,
            trim: true,
            encoding: 'utf8',
            relax_column_count: true,
        });

        return records.map(record => {
            const keys = Object.keys(record).map(k => k.trim());
            const nameKey = keys.find(k => k.includes('Name'));
            const descKey = keys.find(k => k.includes('Description'));

            return {
                name: (nameKey && record[nameKey]) ? String(record[nameKey]).trim() : '',
                description: (descKey && record[descKey]) ? String(record[descKey]).trim() : ''
            };
        }).filter(item => item.name || item.description);

    } catch (e) {
        console.error("CSV Parsing Error:", e);
        return [];
    }
}

/**
 * Parses an Excel file using xlsx library.
 */
function parseExcelFile(filePath) {
    try {
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        const data = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
            defval: ''
        });

        if (data.length < 2) {
            console.warn("Excel file has insufficient data");
            return [];
        }

        const headers = data[1];
        const rows = data.slice(2);

        const nameIndex = headers.findIndex(h => String(h).includes('Name'));
        const descIndex = headers.findIndex(h => String(h).includes('Description'));

        return rows.map(row => {
            const name = nameIndex >= 0 ? String(row[nameIndex] || '').trim() : '';
            const description = descIndex >= 0 ? String(row[descIndex] || '').trim() : '';

            return { name, description };
        }).filter(item => item.name || item.description);

    } catch (e) {
        console.error("Excel Parsing Error:", e);
        return [];
    }
}

/**
 * Reformats the competency data into the target CSV structure.
 */
function cleanAndReformatCompetencies(data, programName, fileId) {
    const reformattedData = [];

    const mainFramework = data[0];

    // Generate Prefix from the Program Name (Short Name) for children
    // e.g. "Physical Education" -> "PE"
    const prefix = generateCode(mainFramework.name || programName);

    // Root Row uses the user-provided fileId
    const rootRowId = fileId;

    reformattedData.push({
        "Parent ID number": "",
        "ID number": rootRowId,
        "Short name": mainFramework.name,
        "Description": mainFramework.description,
        "Is framework": 1,
        "Exported ID (optional)": "",
        ...STANDARD_VALUES
    });

    const structureMap = {
        'Knowledge': { id: `${prefix}-K`, parent: '' },
        'Skills': { id: `${prefix}-S`, parent: '' },
        'Competence': { id: `${prefix}-C`, parent: '' }
    };

    const subCategoryParents = {
        'Theoretical Understanding': 'Knowledge',
        'Applied Knowledge': 'Knowledge',
        'Practical Application': 'Knowledge',
        'Communication Skills': 'Skills',
        'Generic Problem Solving': 'Skills',
        'Critical Thinking': 'Skills',
        'Autonomy & Responsibility': 'Competence'
    };

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const name = row.name;
        const description = row.description;

        if (!name) continue;

        let parentId = '';
        let idNumber = '';
        let shortName = name;

        const code = generateCode(name);

        if (structureMap[name]) {
            idNumber = structureMap[name].id;
            parentId = structureMap[name].parent;
        } else if (subCategoryParents[name]) {
            const parentName = subCategoryParents[name];
            const parentObj = structureMap[parentName];

            if (parentObj) {
                parentId = parentObj.id;
                idNumber = `${parentId}-${code}`;
                structureMap[name] = { id: idNumber, parent: parentId };
            } else {
                parentId = '';
                idNumber = `${prefix}-${code}`;
            }
        } else {
            let parentName = '';
            if (name.startsWith('K') || name.startsWith('k')) parentName = 'Theoretical Understanding';
            else if (name.startsWith('S') || name.startsWith('s')) parentName = 'Generic Problem Solving';
            else if (name.startsWith('C') || name.startsWith('c')) parentName = 'Autonomy & Responsibility';

            if (parentName && structureMap[parentName]) {
                parentId = structureMap[parentName].id;
                idNumber = `${parentId}-${code}`;
            } else {
                parentId = '';
                idNumber = `${prefix}-${code}`;
            }
        }

        reformattedData.push({
            "Parent ID number": parentId,
            "ID number": idNumber,
            "Short name": shortName,
            "Description": description,
            "Is framework": "",
            ...ROW_STANDARD_VALUES
        });
    }

    return reformattedData;
}

/**
 * Converts the reformatted data array to a CSV string.
 */
function toCsvString(data) {
    if (data.length === 0) return '';

    let csv = TARGET_COLUMNS.join(',') + '\n';

    data.forEach(row => {
        const values = TARGET_COLUMNS.map(col => {
            let value = row[col] !== undefined ? String(row[col]) : '';
            if (value.includes(',') || value.includes('"') || value.includes('\n') || value.includes(' ') || value.includes('[')) {
                value = `"${value.replace(/"/g, '""')}"`;
            }
            return value;
        });
        csv += values.join(',') + '\n';
    });

    return csv;
}

function ensureDirectoryExists(dirPath) {
    if (!fs.existsSync(dirPath)) {
        fs.mkdirSync(dirPath, { recursive: true });
        console.log(`Created directory: ${dirPath}`);
    }
}

function askForFileId(fileName) {
    return new Promise((resolve) => {
        const rl = readline.createInterface({
            input: process.stdin,
            output: process.stdout
        });

        rl.question(`Enter ID number for "${fileName}" (press Enter for default 2299): `, (answer) => {
            rl.close();
            const fileId = answer.trim() || "2299";
            resolve(fileId);
        });
    });
}

async function processBatchFiles() {
    const beforeDir = './before-convert';
    const afterDir = './after-convert';

    ensureDirectoryExists(beforeDir);
    ensureDirectoryExists(afterDir);

    try {
        const files = fs.readdirSync(beforeDir);
        const processableFiles = files.filter(file => {
            const ext = path.extname(file).toLowerCase();
            return ext === '.csv' || ext === '.xls' || ext === '.xlsx';
        });

        if (processableFiles.length === 0) {
            console.log('No CSV or Excel files found in before-convert directory.');
            return;
        }

        console.log(`Found ${processableFiles.length} files to process...`);

        for (const fileName of processableFiles) {
            try {
                const fileId = await askForFileId(fileName);
                const inputPath = path.join(beforeDir, fileName);
                const fileExt = path.extname(fileName).toLowerCase();
                const outputFileName = `Reformatted - ${path.basename(fileName, fileExt)}.csv`;
                const outputPath = path.join(afterDir, outputFileName);
                const programName = path.basename(fileName, fileExt);

                let parsedData;

                if (fileExt === '.csv') {
                    const rawContent = fs.readFileSync(inputPath, { encoding: 'utf8' });
                    parsedData = parseCsvWithLibrary(rawContent);
                } else if (fileExt === '.xls' || fileExt === '.xlsx') {
                    parsedData = parseExcelFile(inputPath);
                }

                if (!parsedData || parsedData.length === 0) {
                    console.warn(`Skipping ${fileName}: Could not parse any meaningful data.`);
                    continue;
                }

                const reformattedData = cleanAndReformatCompetencies(parsedData, programName, fileId);
                const csvOutput = toCsvString(reformattedData);

                fs.writeFileSync(outputPath, '\ufeff' + csvOutput, { encoding: 'utf8' });
                console.log(`✅ Processed: ${fileName} -> ${outputFileName} (ID: ${fileId})`);

            } catch (error) {
                console.error(`❌ Error processing file ${fileName}: ${error.message}`);
            }
        }
        console.log("Batch processing completed!");
    } catch (error) {
        console.error(`Error reading before-convert directory: ${error.message}`);
    }
}

if (require.main === module) {
    processBatchFiles();
}

module.exports = {
    cleanAndReformatCompetencies,
    toCsvString,
    parseCsvWithLibrary,
    parseExcelFile,
    TARGET_COLUMNS,
    STANDARD_VALUES
};