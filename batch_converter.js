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
    "Scale configuration": "[{\"\"scaleid\"\":\"\"2\"\"},{\"\"id\"\":2,\"\"scaledefault\"\":1,\"\"proficient\"\":1}]",
    "Taxonomy": "competency"
};

// Map of major area titles to their short codes/IDs
const MAJOR_AREA_MAP = {
    "Knowledge": "K",
    "Theoretical Understanding": "K-TU",
    "Practical Application": "K-PA",
    "Skills": "S",
    "Generic Problem Solving": "S-GPS",
    "Communication Skills": "S-CS",
    "Competence": "C",
    "Autonomy & Responsibility": "C-ARC"
};

// Map of Arabic program names to English prefixes
const PROGRAM_PREFIX_MAP = {
    "مخرجات برنامج ادارة الموارد البشرية": "HR",
    "مخرجات برنامج ذكاء الاعمال": "BI",
    "مخرجات برنامج التكنولوجيا المالية": "FT",
    "مخرجات برنامج التسويق الرقمي": "DM",
    "مخرجات برنامج المحاسبة": "AC",
    "مخرجات برنامج قسم العلوم الجمركية والضريبية": "CT",
    "مخرجات برنامج تكنولوجيا معلومات الاعمال": "IT",
    "مخرجات برنامج ادارة الاعمال": "BA",
    "Family Reform and Guidance": "FRG"
};

// --- Helper Functions ---

/**
 * Creates a simple ID prefix from the program name (e.g., 'Business Administration' -> 'BA').
 * @param {string} name - The program name.
 * @returns {string} - The ID prefix.
 */
function createFrameworkPrefix(name) {
    // Remove Arabic words like 'برنامج' (program) or 'قسم' (department) for better prefix
    const cleanName = name
        .replace(/مخرجات برنامج /g, '')
        .replace(/قسم /g, '')
        .split(/\s+-\s*/)[0];

    const parts = cleanName.split(/\s+/).filter(p => p.length > 0);
    if (parts.length > 1) {
        // Use English initials (e.g., 'Business Administration' -> 'BA')
        const englishMatch = cleanName.match(/[A-Z]/g);
        if (englishMatch && englishMatch.length >= 2) {
            return englishMatch.join('');
        }
        // Use a short version of the first two words of the Arabic name
        return parts.slice(0, 2).map(p => p.charAt(0)).join('').toUpperCase();
    }
    // Fallback for single-word names (e.g., 'Accounting')
    return cleanName.substring(0, 3).toUpperCase();
}

/**
 * Parses a CSV file using csv-parse library.
 * Assumes the first line is metadata and the second line is the header ('Name,Description').
 * @param {string} csvContent - The raw CSV file content.
 * @returns {Array<Object>} - Array of competency objects {name, description}.
 */
function parseCsvWithLibrary(csvContent) {
    try {
        // Use columns: true to auto-detect header (Name, Description) from the second line
        // Use skip_records: 1 to skip the very first line of the file (metadata)
        const records = parse(csvContent, {
            columns: true,
            skip_records: 1,
            skip_empty_lines: true,
            trim: true,
            // Ensure encoding is handled correctly, utf8-bom is common for Arabic CSVs
            encoding: 'utf8',
            relax_column_count: true, // Allow rows to have fewer columns if needed
        });

        // Normalize the column names to handle potential whitespace differences
        return records.map(record => {
            const keys = Object.keys(record).map(k => k.trim());
            const nameKey = keys.find(k => k.includes('Name'));
            const descKey = keys.find(k => k.includes('Description'));

            return {
                name: (nameKey && record[nameKey]) ? String(record[nameKey]).trim() : '',
                description: (descKey && record[descKey]) ? String(record[descKey]).trim() : ''
            };
        }).filter(item => item.name || item.description); // Filter out rows with no data

    } catch (e) {
        console.error("CSV Parsing Error:", e);
        return [];
    }
}

/**
 * Parses an Excel file using xlsx library.
 * Assumes the first row is metadata and the second row is the header ('Name,Description').
 * @param {string} filePath - Path to the Excel file.
 * @returns {Array<Object>} - Array of competency objects {name, description}.
 */
function parseExcelFile(filePath) {
    try {
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0]; // Use the first sheet
        const worksheet = workbook.Sheets[sheetName];

        // Convert to JSON with header detection
        const data = XLSX.utils.sheet_to_json(worksheet, {
            header: 1, // Get as array of arrays
            defval: ''
        });

        if (data.length < 2) {
            console.warn("Excel file has insufficient data");
            return [];
        }

        // Skip first row (metadata) and use second row as headers
        const headers = data[1];
        const rows = data.slice(2);

        // Find column indices for Name and Description
        const nameIndex = headers.findIndex(h => String(h).includes('Name'));
        const descIndex = headers.findIndex(h => String(h).includes('Description'));

        return rows.map(row => {
            const name = nameIndex >= 0 ? String(row[nameIndex] || '').trim() : '';
            const description = descIndex >= 0 ? String(row[descIndex] || '').trim() : '';

            return { name, description };
        }).filter(item => item.name || item.description); // Filter out empty rows

    } catch (e) {
        console.error("Excel Parsing Error:", e);
        return [];
    }
}

/**
 * Reformats the competency data into the target CSV structure.
 * @param {Array<Object>} data - Parsed competency data.
 * @param {string} programName - The name of the program for ID generation.
 * @param {string} fileId - The custom file ID provided by user.
 * @returns {Array<Object>} - Array of reformatted rows.
 */
function cleanAndReformatCompetencies(data, programName, fileId) {
    const reformattedData = [];

    // Create the unique framework ID
    const frameworkPrefix = createFrameworkPrefix(programName);
    const frameworkId = `${frameworkPrefix}-FRWK`;

    // Track the current parent ID for specific competencies (e.g., K1, S1)
    let currentSubAreaId = '';

    // --- 1. Process the main framework row ---
    // The first record is the main program area
    const mainFramework = data[0];

    reformattedData.push({
        "Parent ID number": "",
        "ID number": fileId, // Use user-provided file ID
        "Short name": mainFramework.name, // Use the original name as short name
        "Description": mainFramework.description, // Preserve all content without removal
        "Is framework": 1,
        ...STANDARD_VALUES
    });

    // --- 2. Process all other competency rows (starting from index 1) ---
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const name = row.name;
        const description = row.description;

        let parentId = '';
        let idNumber = '';
        let shortName = name; // Always use the original name as short name

        // Get English prefix from mapping or use createFrameworkPrefix as fallback
        const englishPrefix = PROGRAM_PREFIX_MAP[programName] || createFrameworkPrefix(programName);

        if (name in MAJOR_AREA_MAP) {
            // This is a Major Area (Knowledge, Skills, Competence) or Sub-Area (TU, GPS, ARC, etc.)
            const areaCode = MAJOR_AREA_MAP[name];

            // Generate ID as englishPrefix-areaCode
            idNumber = `${englishPrefix}-${areaCode}`;

            // Determine Parent ID for Major/Sub-Areas
            if (name === 'Knowledge' || name === 'Skills' || name === 'Competence') {
                // K, S, C have no parent (empty Parent ID)
                parentId = ""; // No parent for Knowledge, Skills, Competence
            } else if (name === 'Theoretical Understanding' || name === 'Practical Application') {
                // TU, PA are children of K
                parentId = `${englishPrefix}-K`;
            } else if (name === 'Generic Problem Solving' || name === 'Communication Skills') {
                // GPS, CS are children of S
                parentId = `${englishPrefix}-S`;
            } else if (name === 'Autonomy & Responsibility') {
                // ARC is a child of C
                parentId = `${englishPrefix}-C`;
            }

            // Set the current sub-area parent for the specific competencies (K1, S1, C1...)
            currentSubAreaId = idNumber;

        } else if (name.match(/^[Kk]\d+$/)) {
            // Specific K competencies (e.g., K1, k6)
            const areaCode = MAJOR_AREA_MAP['Theoretical Understanding'];
            idNumber = `${englishPrefix}-${areaCode}-${name.toUpperCase()}`;
            parentId = `${englishPrefix}-${areaCode}`; // Kx parent is Theoretical Understanding (TU)

        } else if (name.match(/^[Ss]\d+$/)) {
             // Specific S competencies (e.g., S1, S3)
            const areaCode = MAJOR_AREA_MAP['Generic Problem Solving'];
            idNumber = `${englishPrefix}-${areaCode}-${name.toUpperCase()}`;
            parentId = `${englishPrefix}-${areaCode}`; // Sx parent is Generic Problem Solving (GPS)

        } else if (name.match(/^[Cc]\d+$/)) {
             // Specific C competencies (e.g., C1, C2)
            const areaCode = MAJOR_AREA_MAP['Autonomy & Responsibility'];
            idNumber = `${englishPrefix}-${areaCode}-${name.toUpperCase()}`;
            parentId = `${englishPrefix}-${areaCode}`; // Cx parent is Autonomy & Responsibility (ARC)

        } else {
            // Should not happen if data is clean, but assign the last known area as parent
            console.warn(`[${programName}] Fallback: Uncategorized row "${name}"`);
            idNumber = `${englishPrefix}-${name}`;
            parentId = currentSubAreaId;
        }

        // Skip the main framework row which was already added at index 0
        if (i === 0) continue;

        // Add the reformatted row
        reformattedData.push({
            "Parent ID number": parentId,
            "ID number": idNumber,
            "Short name": shortName,
            "Description": description, // Preserve all content without removal
            "Is framework": "",
            ...STANDARD_VALUES
        });
    }

    return reformattedData;
}

/**
 * Converts the reformatted data array to a CSV string.
 */
function toCsvString(data) {
    if (data.length === 0) return '';

    // Header row
    let csv = TARGET_COLUMNS.join(',') + '\n';

    // Data rows
    data.forEach(row => {
        const values = TARGET_COLUMNS.map(col => {
            let value = row[col] !== undefined ? String(row[col]) : '';

            // Escape double quotes and wrap in quotes if value contains comma or quotes
            if (value.includes(',') || value.includes('"') || value.includes('\n') || value.includes(' ') || value.includes('[')) {
                // Replace all internal double-quotes with two double-quotes for CSV escaping
                value = `"${value.replace(/"/g, '""')}"`;
            }
            return value;
        });
        csv += values.join(',') + '\n';
    });

    return csv;
}

/**
 * Ensures a directory exists, creates it if it doesn't
 */
function ensureDirectoryExists(dirPath) {
    if (!fs.existsSync(dirPath)) {
        fs.mkdirSync(dirPath, { recursive: true });
        console.log(`Created directory: ${dirPath}`);
    }
}

/**
 * Ask user for file ID number
 */
function askForFileId(fileName) {
    return new Promise((resolve) => {
        const rl = readline.createInterface({
            input: process.stdin,
            output: process.stdout
        });

        rl.question(`Enter ID number for "${fileName}" (press Enter for default 2299): `, (answer) => {
            rl.close();
            // Use user input if provided, otherwise use default
            const fileId = answer.trim() || "2299";
            resolve(fileId);
        });
    });
}

// --- Main Execution Logic ---

async function processBatchFiles() {
    const beforeDir = './before-convert';
    const afterDir = './after-convert';

    // Ensure directories exist
    ensureDirectoryExists(beforeDir);
    ensureDirectoryExists(afterDir);

    try {
        // Get all files from before-convert directory
        const files = fs.readdirSync(beforeDir);

        // Filter for CSV and Excel files
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
                // Ask user for file ID
                const fileId = await askForFileId(fileName);

                const inputPath = path.join(beforeDir, fileName);
                const fileExt = path.extname(fileName).toLowerCase();
                const outputFileName = `Reformatted - ${path.basename(fileName, fileExt)}.csv`;
                const outputPath = path.join(afterDir, outputFileName);

                // Extract the clean program name (remove file extension)
                const programName = path.basename(fileName, fileExt);

                let parsedData;

                // 1. Parse the input file based on its type
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

                // 2. Reformat the data with user-provided file ID
                const reformattedData = cleanAndReformatCompetencies(parsedData, programName, fileId);

                // 3. Convert to CSV string
                const csvOutput = toCsvString(reformattedData);

                // 4. Save the new CSV file
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

// Run the batch processor
processBatchFiles();