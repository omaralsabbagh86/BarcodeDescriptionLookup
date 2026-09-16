const csvFile = document.getElementById("csvFile");
const fileInfo = document.getElementById("fileInfo");
const sheetWrapper = document.getElementById("sheetWrapper");
const sheetSelect = document.getElementById("sheetSelect");
const barcodeColumn = document.getElementById("barcodeColumn");
const startButton = document.getElementById("startButton");
const downloadButton = document.getElementById("downloadButton");
const outputFormat = document.getElementById("outputFormat");
const wantedButton = document.getElementById("wantedButton");

const rowsElement = document.getElementById("rows");
const uniqueElement = document.getElementById("unique");
const foundElement = document.getElementById("found");
const notFoundElement = document.getElementById("notFound");

const progressBar = document.getElementById("progressBar");
const progressText = document.getElementById("progressText");
const resultsTable = document.getElementById("resultsTable");

const lookupSource = document.getElementById("lookupSource");
const apiFallback = document.getElementById("apiFallback");

const EXCEL_EXTENSIONS = ["xlsx", "xlsm", "xls"];
const EXCEL_MAX_ROWS = 1048576;

// Your public retail database on GitHub (files split by barcode prefix)
const DB_BASE_URL = "https://raw.githubusercontent.com/omaralsabbagh86/retail-ean13-db/main/data/";
const DB_FETCH_WORKERS = 4;

// Open Food Facts API is rate-limited: keep it for small batches only
const API_MAX_BARCODES = 300;
const API_DELAY_MS = 1200;

let sourceRows = [];
let headers = [];
let results = [];
let currentWorkbook = null;
let currentFileName = "";
let sourceFileBaseName = "Online_Item_Descriptions";


/* ---------------- CSV parsing ---------------- */

function parseCSV(text) {

    text = text.replace(/^\uFEFF/, ""); // remove BOM

    const rows = [];
    let row = [];
    let field = "";
    let insideQuotes = false;

    for (let i = 0; i < text.length; i++) {

        const char = text[i];
        const next = text[i + 1];

        if (char === '"') {

            if (insideQuotes && next === '"') {
                field += '"';
                i++;
            } else {
                insideQuotes = !insideQuotes;
            }

        } else if (char === "," && !insideQuotes) {

            row.push(field);
            field = "";

        } else if ((char === "\n" || char === "\r") && !insideQuotes) {

            if (char === "\r" && next === "\n") {
                i++;
            }

            row.push(field);
            field = "";

            if (row.some(value => value.trim() !== "")) {
                rows.push(row);
            }

            row = [];

        } else {

            field += char;

        }
    }

    if (field !== "" || row.length > 0) {

        row.push(field);

        if (row.some(value => value.trim() !== "")) {
            rows.push(row);
        }
    }

    return rows;
}


/* ---------------- Excel parsing ---------------- */

function ensureXLSX() {

    if (typeof XLSX === "undefined") {
        throw new Error(
            "Excel library (SheetJS) is not loaded. Check that lib/xlsx.full.min.js exists or that you are online."
        );
    }
}

function getFileExtension(name) {

    const match = /\.([^.]+)$/.exec(name || "");

    return match ? match[1].toLowerCase() : "";
}

// Convert one Excel cell to text without losing barcode digits
function cellToText(cell) {

    if (!cell) {
        return "";
    }

    if (cell.t === "n" && Number.isFinite(cell.v)) {

        const formatted = cell.w !== undefined ? String(cell.w) : "";

        // Large numbers shown as 6.29112E+12 -> write full digits instead
        if (!formatted || /e[+-]?\d/i.test(formatted)) {

            return Number.isInteger(cell.v)
                ? cell.v.toFixed(0)
                : String(cell.v);
        }

        return formatted;
    }

    if (cell.w !== undefined) {
        return String(cell.w);
    }

    return cell.v === undefined || cell.v === null ? "" : String(cell.v);
}

function sheetToRows(sheet) {

    if (!sheet || !sheet["!ref"]) {
        return [];
    }

    const range = XLSX.utils.decode_range(sheet["!ref"]);

    // Dense mode: SheetJS 0.20 uses sheet["!data"], 0.18 uses the sheet itself as an array
    const dense = sheet["!data"] || (Array.isArray(sheet) ? sheet : null);

    const rows = [];

    for (let r = range.s.r; r <= range.e.r; r++) {

        const denseRow = dense ? dense[r] : null;

        if (dense && !denseRow) {
            continue;
        }

        const row = [];
        let hasValue = false;

        for (let c = range.s.c; c <= range.e.c; c++) {

            const cell = dense
                ? denseRow[c]
                : sheet[XLSX.utils.encode_cell({ r, c })];

            const text = cellToText(cell);

            if (!hasValue && text.trim() !== "") {
                hasValue = true;
            }

            row.push(text);
        }

        if (hasValue) {
            rows.push(row);
        }
    }

    return rows;
}


/* ---------------- Barcode helpers ---------------- */

function normalizeBarcode(value) {

    if (value === null || value === undefined) {
        return "";
    }

    let barcode = String(value).trim();

    // Excel-style numeric value such as 1234567890123.0
    barcode = barcode.replace(/\.0+$/, "");

    // Keep digits only
    barcode = barcode.replace(/\D/g, "");

    return barcode;
}


// Convert to the 13-digit form used by the retail database
function toEan13(barcode) {

    let digits = normalizeBarcode(barcode);

    if (digits.length === 14 && digits.startsWith("0")) {
        digits = digits.slice(1);
    }

    if (!digits || digits.length > 13) {
        return "";
    }

    return digits.padStart(13, "0");
}


function hasValidCheckDigit(ean13) {

    let sum = 0;

    for (let i = 0; i < 12; i++) {
        sum += Number(ean13[i]) * (i % 2 ? 3 : 1);
    }

    return (10 - (sum % 10)) % 10 === Number(ean13[12]);
}


// Tells whether a barcode can exist in online product sources
function classifyBarcode(barcode) {

    const digits = normalizeBarcode(barcode);

    if (!digits) {
        return "EMPTY";
    }

    if (digits.length < 8) {
        return "SHORT / INTERNAL CODE";
    }

    const ean = toEan13(digits);

    if (!ean) {
        return "INVALID LENGTH";
    }

    if (!hasValidCheckDigit(ean)) {
        return "INVALID CHECK DIGIT";
    }

    if (/^(2|02|04|05|98|99)/.test(ean)) {
        return "IN-STORE / RESTRICTED CODE";
    }

    return "VALID";
}


function databasePrefix(ean13) {

    return ean13.startsWith("978") || ean13.startsWith("979")
        ? ean13.slice(0, 5)
        : ean13.slice(0, 3);
}


function detectBarcodeColumn(headerList) {

    const preferred = [
        "ean13 code",
        "ean13",
        "ean 13",
        "barcode",
        "gtin",
        "upc"
    ];

    for (const preferredName of preferred) {

        const index = headerList.findIndex(
            h => h.trim().toLowerCase() === preferredName
        );

        if (index !== -1) {
            return index;
        }
    }

    return -1;
}


function getUniqueBarcodes(columnIndex) {

    const seen = new Set();
    const list = [];

    sourceRows.forEach(row => {

        const barcode = normalizeBarcode(row[columnIndex]);

        if (barcode && !seen.has(barcode)) {
            seen.add(barcode);
            list.push(barcode);
        }
    });

    return list;
}


/* ---------------- Loading data ---------------- */

function resetResults() {

    resultsTable.innerHTML = "";
    results = [];
    downloadButton.disabled = true;
    wantedButton.disabled = true;
    progressBar.style.width = "0%";
    progressText.textContent = "Ready to start";
    foundElement.textContent = "0";
    notFoundElement.textContent = "0";
}


function loadParsedRows(parsed, sheetName) {

    if (parsed.length < 2) {
        throw new Error("File does not contain enough data (header row + at least one data row).");
    }

    headers = parsed[0].map((h, i) => {
        const name = String(h || "").trim();
        return name || `Column ${i + 1}`;
    });

    // Loop instead of Math.max(...array): spreading 800k rows overflows the call stack
    let columnCount = headers.length;

    for (let i = 0; i < parsed.length; i++) {
        if (parsed[i].length > columnCount) {
            columnCount = parsed[i].length;
        }
    }

    while (headers.length < columnCount) {
        headers.push(`Column ${headers.length + 1}`);
    }

    sourceRows = parsed.slice(1);

    sourceRows.forEach(row => {
        while (row.length < columnCount) {
            row.push("");
        }
    });

    barcodeColumn.innerHTML = "";

    headers.forEach((header, index) => {

        const option = document.createElement("option");

        option.value = index;
        option.textContent = header;

        barcodeColumn.appendChild(option);
    });

    const detected = detectBarcodeColumn(headers);
    const selectedColumn = detected !== -1 ? detected : 0;

    barcodeColumn.value = selectedColumn;

    const uniqueCount = getUniqueBarcodes(selectedColumn).length;

    rowsElement.textContent = sourceRows.length.toLocaleString();
    uniqueElement.textContent = uniqueCount.toLocaleString();

    const sheetText = sheetName ? ` | Sheet: ${sheetName}` : "";

    fileInfo.textContent =
        `Loaded: ${currentFileName}${sheetText} | Rows: ${sourceRows.length.toLocaleString()} | Unique barcodes: ${uniqueCount.toLocaleString()}`;

    startButton.disabled = uniqueCount === 0;

    resetResults();
}


csvFile.addEventListener("change", async function () {

    const file = csvFile.files[0];

    if (!file) {
        return;
    }

    currentFileName = file.name;
    sourceFileBaseName = file.name.replace(/\.[^.]+$/, "") || "Online_Item_Descriptions";

    const extension = getFileExtension(file.name);

    fileInfo.textContent = "Reading file...";

    try {

        if (extension === "csv") {

            currentWorkbook = null;
            sheetWrapper.hidden = true;
            sheetSelect.innerHTML = "";

            const text = await file.text();

            outputFormat.value = "csv";

            loadParsedRows(parseCSV(text), "");

        } else if (EXCEL_EXTENSIONS.includes(extension)) {

            ensureXLSX();

            const buffer = await file.arrayBuffer();

            fileInfo.textContent = "Reading Excel file (large files can take a minute)...";

            await new Promise(resolve => setTimeout(resolve, 50));

            currentWorkbook = XLSX.read(buffer, {
                type: "array",
                dense: true,
                cellFormula: false,
                cellHTML: false,
                cellStyles: false
            });

            if (!currentWorkbook.SheetNames.length) {
                throw new Error("Workbook contains no sheets.");
            }

            sheetSelect.innerHTML = "";

            currentWorkbook.SheetNames.forEach(name => {

                const option = document.createElement("option");

                option.value = name;
                option.textContent = name;

                sheetSelect.appendChild(option);
            });

            sheetWrapper.hidden = currentWorkbook.SheetNames.length < 2;

            const firstSheet = currentWorkbook.SheetNames[0];

            outputFormat.value = "xlsx";

            loadParsedRows(
                sheetToRows(currentWorkbook.Sheets[firstSheet]),
                firstSheet
            );

            // Single sheet: free the workbook memory
            if (currentWorkbook.SheetNames.length < 2) {
                currentWorkbook = null;
            }

        } else {

            throw new Error("Unsupported file type. Please select a .csv, .xlsx, .xlsm or .xls file.");
        }

    } catch (error) {

        console.error(error);

        fileInfo.textContent = "Error reading file: " + error.message;

        startButton.disabled = true;
    }
});


sheetSelect.addEventListener("change", function () {

    if (!currentWorkbook) {
        return;
    }

    const name = sheetSelect.value;

    try {

        loadParsedRows(
            sheetToRows(currentWorkbook.Sheets[name]),
            name
        );

    } catch (error) {

        console.error(error);

        fileInfo.textContent = `Error reading sheet "${name}": ` + error.message;

        startButton.disabled = true;
    }
});


barcodeColumn.addEventListener("change", function () {

    if (!sourceRows.length) {
        return;
    }

    const uniqueCount =
        getUniqueBarcodes(Number(barcodeColumn.value)).length;

    uniqueElement.textContent = uniqueCount.toLocaleString();

    startButton.disabled = uniqueCount === 0;
});


/* ---------------- Online lookup ---------------- */

async function lookupBarcode(barcode) {

    const url =
        `https://world.openfoodfacts.org/api/v2/product/${encodeURIComponent(barcode)}.json`;

    try {

        const response = await fetch(url);

        if (!response.ok) {

            return {
                barcode,
                description: "",
                source: "Open Food Facts",
                status: "NOT FOUND"
            };
        }

        const data = await response.json();

        if (data.status === 1 && data.product) {

            const product = data.product;

            const description =
                product.product_name ||
                product.product_name_en ||
                product.generic_name ||
                product.generic_name_en ||
                "";

            if (description) {

                return {
                    barcode,
                    description,
                    source: "Open Food Facts",
                    status: "FOUND"
                };
            }
        }

        return {
            barcode,
            description: "",
            source: "Open Food Facts",
            status: "NOT FOUND"
        };

    } catch (error) {

        return {
            barcode,
            description: "",
            source: "Open Food Facts",
            status: "ERROR"
        };
    }
}


const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));


function makeResult(barcode, description, source, status) {
    return { barcode, description, source, status };
}


async function fetchDatabaseFile(prefix) {

    const url = `${DB_BASE_URL}ean_${prefix}.csv`;

    for (let attempt = 1; attempt <= 3; attempt++) {

        try {

            const response = await fetch(url);

            if (response.status === 404) {
                return "";
            }

            if (response.ok) {
                return await response.text();
            }

        } catch (error) {
            // retry below
        }

        await sleep(1000 * attempt);
    }

    throw new Error(`Could not download ean_${prefix}.csv`);
}


async function lookupInDatabase(barcodes) {

    const resultMap = new Map();

    // prefix -> Map(ean13 -> [original barcodes])
    const groups = new Map();

    barcodes.forEach(barcode => {

        const ean = toEan13(barcode);

        if (!ean) {
            resultMap.set(barcode, makeResult(barcode, "", "Retail DB", "INVALID"));
            return;
        }

        const prefix = databasePrefix(ean);

        if (!groups.has(prefix)) {
            groups.set(prefix, new Map());
        }

        const group = groups.get(prefix);

        if (!group.has(ean)) {
            group.set(ean, []);
        }

        group.get(ean).push(barcode);
    });

    const prefixes = [...groups.keys()].sort();

    let nextIndex = 0;
    let filesDone = 0;
    let found = 0;

    async function worker() {

        while (nextIndex < prefixes.length) {

            const prefix = prefixes[nextIndex++];
            const group = groups.get(prefix);

            let text = "";
            let failed = false;

            try {
                text = await fetchDatabaseFile(prefix);
            } catch (error) {
                console.error(error);
                failed = true;
            }

            if (text) {

                const rows = parseCSV(text);

                for (let i = 1; i < rows.length; i++) {

                    const ean = (rows[i][0] || "").trim();
                    const originals = group.get(ean);

                    if (originals) {

                        const description = rows[i][1] || "";

                        originals.forEach(original => {
                            resultMap.set(original, makeResult(original, description, "Retail DB", description ? "FOUND" : "NOT FOUND"));
                            if (description) found++;
                        });

                        group.delete(ean);
                    }
                }
            }

            // Whatever is left in this prefix was not found
            group.forEach(originals => {
                originals.forEach(original => {
                    resultMap.set(original, makeResult(original, "", "Retail DB", failed ? "ERROR" : "NOT FOUND"));
                });
            });

            groups.delete(prefix);

            filesDone++;

            const percent = Math.round((filesDone / prefixes.length) * 100);

            progressBar.style.width = percent + "%";

            progressText.textContent =
                `Retail database: ${filesDone.toLocaleString()} / ${prefixes.length.toLocaleString()} files checked (${percent}%) | Found so far: ${found.toLocaleString()}`;

            foundElement.textContent = found.toLocaleString();
        }
    }

    const workers = [];

    for (let i = 0; i < Math.min(DB_FETCH_WORKERS, prefixes.length); i++) {
        workers.push(worker());
    }

    await Promise.all(workers);

    return resultMap;
}


async function lookupWithApi(barcodes, resultMap) {

    for (let i = 0; i < barcodes.length; i++) {

        const barcode = barcodes[i];
        const result = await lookupBarcode(barcode);

        if (result.status === "FOUND" || !resultMap.has(barcode)) {
            resultMap.set(barcode, result);
        }

        const percent = Math.round(((i + 1) / barcodes.length) * 100);
        const secondsLeft = Math.round(((barcodes.length - i - 1) * API_DELAY_MS) / 1000);

        progressBar.style.width = percent + "%";

        progressText.textContent =
            `Open Food Facts: ${(i + 1).toLocaleString()} / ${barcodes.length.toLocaleString()} (${percent}%) | about ${secondsLeft}s left`;

        if (i < barcodes.length - 1) {
            await sleep(API_DELAY_MS);
        }
    }
}


function updateStats(data) {

    let found = 0;

    for (let i = 0; i < data.length; i++) {
        if (data[i] && data[i].status === "FOUND") {
            found++;
        }
    }

    foundElement.textContent = found.toLocaleString();
    notFoundElement.textContent = (data.length - found).toLocaleString();
}


function setControlsDisabled(disabled) {

    startButton.disabled = disabled;
    csvFile.disabled = disabled;
    barcodeColumn.disabled = disabled;
    sheetSelect.disabled = disabled;
    lookupSource.disabled = disabled;
    apiFallback.disabled = disabled;
}


startButton.addEventListener("click", async function () {

    if (!sourceRows.length) {
        return;
    }

    const uniqueBarcodes =
        getUniqueBarcodes(Number(barcodeColumn.value));

    if (!uniqueBarcodes.length) {
        return;
    }

    const source = lookupSource.value;

    if (source === "api" && uniqueBarcodes.length > API_MAX_BARCODES) {

        progressText.textContent =
            `Open Food Facts only allows small batches (max ${API_MAX_BARCODES.toLocaleString()} barcodes). You have ${uniqueBarcodes.length.toLocaleString()} - choose "My retail database" instead.`;

        return;
    }

    setControlsDisabled(true);

    downloadButton.disabled = true;
    wantedButton.disabled = true;
    resultsTable.innerHTML = "";
    progressBar.style.width = "0%";
    progressText.textContent = "Starting lookup...";
    foundElement.textContent = "0";
    notFoundElement.textContent = "0";

    try {

        let resultMap;

        if (source === "api") {

            resultMap = new Map();

            await lookupWithApi(uniqueBarcodes, resultMap);

        } else {

            resultMap = await lookupInDatabase(uniqueBarcodes);

            if (apiFallback.checked) {

                const missing = uniqueBarcodes.filter(b => {
                    const r = resultMap.get(b);
                    return r && r.status === "NOT FOUND";
                });

                if (missing.length && missing.length <= API_MAX_BARCODES) {

                    await lookupWithApi(missing, resultMap);

                } else if (missing.length > API_MAX_BARCODES) {

                    console.warn(`Skipped Open Food Facts check: ${missing.length} barcodes not found (limit ${API_MAX_BARCODES}).`);
                }
            }
        }

        results = uniqueBarcodes.map(b =>
            resultMap.get(b) || makeResult(b, "", "", "NOT FOUND")
        );

        updateStats(results);

        resultsTable.innerHTML = "";

        let shown = 0;

        for (const result of results) {

            if (result.status !== "FOUND") {
                continue;
            }

            const tr = document.createElement("tr");

            tr.innerHTML = `
                <td>${escapeHTML(result.barcode)}</td>
                <td>${escapeHTML(result.description)}</td>
                <td>${escapeHTML(result.source)}</td>
                <td>${escapeHTML(result.status)}</td>
            `;

            resultsTable.appendChild(tr);

            if (++shown >= 100) {
                break;
            }
        }

        const foundCount = results.filter(r => r.status === "FOUND").length;
        const searchable = getWantedBarcodes().length;

        progressBar.style.width = "100%";

        progressText.textContent =
            `Completed: ${results.length.toLocaleString()} unique barcodes checked, ${foundCount.toLocaleString()} found, ${searchable.toLocaleString()} missing with a valid barcode`;

        downloadButton.disabled = false;
        wantedButton.disabled = searchable === 0;

    } catch (error) {

        console.error(error);

        progressText.textContent = "Lookup failed: " + error.message;

    } finally {

        setControlsDisabled(false);
    }
});


function escapeHTML(value) {

    return String(value || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}


/* ---------------- Output building ---------------- */

function buildResultRows() {

    const selectedColumn = Number(barcodeColumn.value);

    const resultMap = new Map();

    results.forEach(result => {
        resultMap.set(result.barcode, result);
    });

    const output = [[
        ...headers,
        "Online Item Description",
        "Lookup Source",
        "Lookup Status",
        "Barcode Check"
    ]];

    sourceRows.forEach(row => {

        const result = resultMap.get(normalizeBarcode(row[selectedColumn]));

        output.push([
            ...row,
            result ? result.description : "",
            result ? result.source : "",
            result ? result.status : "",
            classifyBarcode(row[selectedColumn])
        ]);
    });

    return output;
}


function escapeCSV(value) {

    if (value === null || value === undefined) {
        return "";
    }

    const text = String(value);

    if (
        text.includes(",") ||
        text.includes('"') ||
        text.includes("\n") ||
        text.includes("\r")
    ) {
        return '"' + text.replace(/"/g, '""') + '"';
    }

    return text;
}


function createResultCSV() {

    return buildResultRows()
        .map(row => row.map(escapeCSV).join(","))
        .join("\r\n");
}


// Build sheet with every value stored as TEXT so barcodes keep all digits
function buildTextSheet(rows) {

    const sheet = XLSX.utils.aoa_to_sheet(rows, { dense: true });

    const columnCount = rows[0] ? rows[0].length : 0;
    const sampleSize = Math.min(rows.length, 500);

    sheet["!cols"] = [];

    for (let c = 0; c < columnCount; c++) {

        let longest = 8;

        for (let r = 0; r < sampleSize; r++) {
            const length = String(rows[r][c] || "").length;
            if (length > longest) longest = length;
        }

        sheet["!cols"].push({ wch: Math.min(longest + 2, 60) });
    }

    if (sheet["!ref"]) {
        sheet["!autofilter"] = { ref: sheet["!ref"] };
    }

    return sheet;
}


function createResultXLSX(fileName) {

    ensureXLSX();

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
        workbook,
        buildTextSheet(buildResultRows()),
        "Results"
    );

    const notFoundRows = [["Barcode", "EAN-13", "Lookup Status", "Barcode Check"]];

    results.forEach(r => {
        if (r.status !== "FOUND") {
            notFoundRows.push([r.barcode, toEan13(r.barcode), r.status, classifyBarcode(r.barcode)]);
        }
    });

    XLSX.utils.book_append_sheet(
        workbook,
        buildTextSheet(notFoundRows),
        "Not Found Barcodes"
    );

    XLSX.writeFile(workbook, fileName, { compression: true });
}


function downloadBlob(content, fileName, mimeType) {

    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");

    link.href = url;
    link.download = fileName;

    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    URL.revokeObjectURL(url);
}


downloadButton.addEventListener("click", function () {

    if (!results.length) {
        return;
    }

    const baseName = `${sourceFileBaseName}_Online_Descriptions`;

    try {

        if (outputFormat.value === "xlsx" && sourceRows.length + 1 > EXCEL_MAX_ROWS) {

            progressText.textContent =
                `Too many rows for Excel (limit ${EXCEL_MAX_ROWS.toLocaleString()}). Please choose CSV.`;

        } else if (outputFormat.value === "xlsx") {

            progressText.textContent = "Building Excel file (large files can take a minute)...";

            setTimeout(() => {
                try {
                    createResultXLSX(baseName + ".xlsx");
                    progressText.textContent = "Excel file downloaded";
                } catch (error) {
                    console.error(error);
                    progressText.textContent = "Download failed: " + error.message;
                }
            }, 50);

        } else {

            // BOM so Excel opens Arabic / special characters correctly
            downloadBlob(
                "\uFEFF" + createResultCSV(),
                baseName + ".csv",
                "text/csv;charset=utf-8;"
            );
        }

    } catch (error) {

        console.error(error);

        progressText.textContent = "Download failed: " + error.message;
    }
});


/* ---------------- Missing list for the retail database ---------------- */

function getWantedBarcodes() {

    const list = new Set();

    results.forEach(r => {
        if (r.status !== "FOUND" && classifyBarcode(r.barcode) === "VALID") {
            list.add(toEan13(r.barcode));
        }
    });

    return [...list];
}


wantedButton.addEventListener("click", function () {

    const list = getWantedBarcodes();

    if (!list.length) {
        return;
    }

    downloadBlob(
        list.join("\n") + "\n",
        "wanted_barcodes.txt",
        "text/plain;charset=utf-8;"
    );

    progressText.textContent =
        `wanted_barcodes.txt downloaded (${list.length.toLocaleString()} barcodes) - run Setup-FreeSources.ps1 to send it to your database`;
});
