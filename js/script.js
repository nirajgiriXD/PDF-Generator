/**
 * HTML file will have placeholder class `dynamic-value`
 * which will be replaced with excel sheet file values.
 */

// Elements.
const submitMsg = document.getElementById("submit-msg");
const htmlContent = document.getElementById("html-content");
const downloadBtn = document.getElementById("download-btn");
const htmlTemplate = document.getElementById("html-template");
const excelSheet = document.getElementById("excel-sheet");
const bypassFirstRow = document.getElementById("bypass-first-row");
const htmlPlaceholderCount = document.getElementById("html-template-placeholder-count");
const excelRowsPlaceholderCount = document.getElementById("excel-rows-placeholder-count");
const excelColsPlaceholderCount = document.getElementById("excel-cols-placeholder-count");

// Variables.
let htmlFileContent = '';
let excelFileContent = [];
let numOfHtmlPlaceholder = 0;
let numOfExcelRowsPlaceholder = 0;
let numOfExcelColsPlaceholder = 0;

/**
 * Export
 */

// Generate appearance of pdf.
function generateTemplate(columns) {
    let pdfContent = '';
    columns.forEach((column) => {
        pdfContent += `<p>${column}</p>`;
    });
    pdfContent = '<div>' + pdfContent + '</div>';
    return pdfContent;
}

// Convert template into pdf.
function generatePDF(pdfContent) {
    return new Promise((resolve, reject) => {
        html2pdf()
            .from(pdfContent)
            .outputPdf()
            .then((pdf) => {
                resolve(pdf);
            })
            .catch((error) => {
                reject(error);
            });
    });
}

// Generate zip file.
function generateZIP() {
    const pdfPromises = [];
    const zip = new JSZip();
    excelFileContent.forEach(async(columns) => {
        // Add to zip file list.
        pdfPromises.push(generatePDF(generateTemplate(columns)));
    });

    // Wait for all PDF generation promises to resolve
    Promise.all(pdfPromises)
        .then((pdfs) => {
            pdfs.forEach((pdf, index) => {
                const id = excelFileContent[index][0];
                zip.file(`${id}.pdf`, pdf, { binary: true });
            });

            // Generate the zip file
            zip.generateAsync({ type: "blob" }).then((content) => {
                saveAs(content, "pdf.zip");
            });
        })
        .catch((error) => {
            console.error("Error generating PDFs:", error);
        });
}

/**
 * Event Listener
 */

// HTML template file select.
htmlTemplate.addEventListener("change", function(event) {
    const file = event.target.files[0];
    let reader = new FileReader();
    reader.readAsText(file);
    reader.onload = (e) => {
        htmlFileContent = e.target.result;
        htmlContent.innerHTML = htmlFileContent;

        numOfHtmlPlaceholder = document.querySelectorAll(".dynamic-value").length;
        htmlPlaceholderCount.innerText = numOfHtmlPlaceholder;
    };
}, false);

// Excel sheet file select.
excelSheet.addEventListener("change", function(event) {
    const file = event.target.files[0];
    readXlsxFile(file).then((rows) => {
        // Remove heading (first row).
        if (bypassFirstRow.checked) {
            rows.splice(0, 1);
            console.log('true');
        } else {
            console.log('false');
        }
        excelFileContent = rows;
        numOfExcelRowsPlaceholder = rows.length;
        excelRowsPlaceholderCount.innerText = numOfExcelRowsPlaceholder;

        if (0 < numOfExcelRowsPlaceholder) {
            numOfExcelColsPlaceholder = rows[0].length;
            excelColsPlaceholderCount.innerText = numOfExcelColsPlaceholder;
        }
    });
}, false);

// Download Button.
downloadBtn.addEventListener("click", function(event) {
    event.preventDefault();
    if (numOfHtmlPlaceholder === numOfExcelColsPlaceholder) {
        if (0 === numOfHtmlPlaceholder) {
            submitMsg.innerText = 'Upload required files';
        } else {
            generateZIP();
        }
    } else {
        submitMsg.innerText = 'Number of placeholder values did not match';
    }
});