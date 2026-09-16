const csvFile = document.getElementById("csvFile");
const fileInfo = document.getElementById("fileInfo");
const sheetWrapper = document.getElementById("sheetWrapper");
const sheetSelect = document.getElementById("sheetSelect");
const barcodeColumn = document.getElementById("barcodeColumn");
const startButton = document.getElementById("startButton");
const downloadButton = document.getElementById("downloadButton");
const outputFormat = document.getElementById("outputFormat");

const rowsElement = document.getElementById("rows");
const uniqueElement = document.getElementById("unique");
const foundElement = document.getElementById("found");
const notFoundElement = document.getElementById("notFound");

const progressBar = document.getElementById("progressBar");
const progressText = document.getElementById("progressText");
const resultsTable = document.getElementById("resultsTable");

const EXCEL_EXTENSIONS = ["xlsx", "xlsm", "xls"];

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
    const rows = [];

    for (let r = range.s.r; r <= range.e.r; r++) {

        const row = [];

        for (let c = range.s.c; c <= range.e.c; c++) {

            const address = XLSX.utils.encode_cell({ r, c });

            row.push(cellToText(sheet[address]));
        }

        if (row.some(value => value.trim() !== "")) {
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

    const columnCount = Math.max(
        headers.length,
        ...parsed.map(r => r.length)
    );

    while (headers.length < columnCount) {
        headers.push(`Column ${headers.length + 1}`);
    }

    sourceRows = parsed.slice(1).map(row => {

        const copy = row.slice();

        while (copy.length < columnCount) {
            copy.push("");
        }

        return copy;
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

            currentWorkbook = XLSX.read(buffer, { type: "array" });

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


async function processWithConcurrency(items, workerCount = 5) {

    let nextIndex = 0;
    let completed = 0;

    const output = new Array(items.length);

    async function worker() {

        while (true) {

            const index = nextIndex++;

            if (index >= items.length) {
                return;
            }

            output[index] = await lookupBarcode(items[index]);

            completed++;

            const percent =
                Math.round((completed / items.length) * 100);

            progressBar.style.width = percent + "%";

            progressText.textContent =
                `${completed.toLocaleString()} / ${items.length.toLocaleString()} (${percent}%)`;

            updateStats(output);
        }
    }

    const workers = [];

    for (let i = 0; i < Math.min(workerCount, items.length); i++) {
        workers.push(worker());
    }

    await Promise.all(workers);

    return output;
}


function updateStats(data) {

    const valid = data.filter(Boolean);

    const found = valid.filter(x => x.status === "FOUND").length;
    const notFound = valid.filter(x => x.status !== "FOUND").length;

    foundElement.textContent = found.toLocaleString();
    notFoundElement.textContent = notFound.toLocaleString();
}


function setControlsDisabled(disabled) {

    startButton.disabled = disabled;
    csvFile.disabled = disabled;
    barcodeColumn.disabled = disabled;
    sheetSelect.disabled = disabled;
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

    setControlsDisabled(true);

    downloadButton.disabled = true;
    resultsTable.innerHTML = "";
    progressBar.style.width = "0%";
    progressText.textContent = "Starting online lookup...";
    foundElement.textContent = "0";
    notFoundElement.textContent = "0";

    try {

        results = await processWithConcurrency(uniqueBarcodes, 5);

        results = results.filter(Boolean);

        updateStats(results);

        resultsTable.innerHTML = "";

        results
            .filter(x => x.status === "FOUND")
            .slice(0, 100)
            .forEach(result => {

                const tr = document.createElement("tr");

                tr.innerHTML = `
                    <td>${escapeHTML(result.barcode)}</td>
                    <td>${escapeHTML(result.description)}</td>
                    <td>${escapeHTML(result.source)}</td>
                    <td>${escapeHTML(result.status)}</td>
                `;

                resultsTable.appendChild(tr);
            });

        progressText.textContent =
            `Completed ${results.length.toLocaleString()} online lookups`;

        downloadButton.disabled = false;

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
        "Lookup Status"
    ]];

    sourceRows.forEach(row => {

        const result = resultMap.get(normalizeBarcode(row[selectedColumn]));

        output.push([
            ...row,
            result ? result.description : "",
            result ? result.source : "",
            result ? result.status : ""
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

    const sheet = XLSX.utils.aoa_to_sheet(
        rows.map(row => row.map(value =>
            value === null || value === undefined ? "" : String(value)
        ))
    );

    const columnCount = rows[0] ? rows[0].length : 0;
    const sample = rows.slice(0, 500);

    sheet["!cols"] = [];

    for (let c = 0; c < columnCount; c++) {

        let longest = 8;

        sample.forEach(row => {
            longest = Math.max(longest, String(row[c] || "").length);
        });

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

    const summaryRows = [
        ["Barcode", "Online Item Description", "Lookup Source", "Lookup Status"],
        ...results.map(r => [r.barcode, r.description, r.source, r.status])
    ];

    XLSX.utils.book_append_sheet(
        workbook,
        buildTextSheet(summaryRows),
        "Unique Barcodes"
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

        if (outputFormat.value === "xlsx") {

            createResultXLSX(baseName + ".xlsx");

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
