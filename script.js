document.addEventListener('DOMContentLoaded', function() {
    // Get DOM elements
    const vendorUpload = document.getElementById('vendor-upload');
    const invoiceUpload = document.getElementById('invoice-upload');
    const vendorFile = document.getElementById('vendor-file');
    const invoiceFile = document.getElementById('invoice-file');
    const checkMonth = document.getElementById('check-month');
    const analyzeBtn = document.getElementById('analyze-btn');
    const results = document.getElementById('results');
    const resultsBody = document.getElementById('results-body');
    const analysisMonth = document.getElementById('analysis-month');
    const missingCount = document.getElementById('missing-count');
    const exportBtn = document.getElementById('export-btn');
    const loading = document.getElementById('loading');
    const vendorInfo = document.getElementById('vendor-info');
    const invoiceInfo = document.getElementById('invoice-info');

    let vendorData = null;
    let invoiceData = null;

    // Handle drag and drop for vendor file
    setupDragAndDrop(vendorUpload, vendorFile, handleVendorFile);
    setupDragAndDrop(invoiceUpload, invoiceFile, handleInvoiceFile);

    // Handle file input changes
    vendorFile.addEventListener('change', (e) => handleVendorFile(e.target.files[0]));
    invoiceFile.addEventListener('change', (e) => handleInvoiceFile(e.target.files[0]));

    // Handle analyze button click
    analyzeBtn.addEventListener('click', analyzeFiles);

    // Handle export button click
    exportBtn.addEventListener('click', exportResults);

    function setupDragAndDrop(dropZone, fileInput, handleFile) {
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, preventDefaults, false);
        });

        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }

        dropZone.addEventListener('dragenter', () => dropZone.classList.add('dragover'));
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
        dropZone.addEventListener('drop', (e) => {
            dropZone.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            handleFile(file);
        });
    }

    function handleVendorFile(file) {
        if (validateExcelFile(file)) {
            readExcelFile(file, (data) => {
                vendorData = data;
                updateFileInfo(vendorInfo, file.name, true);
            });
        } else {
            updateFileInfo(vendorInfo, 'Invalid file format. Please use Excel files (.xlsx, .xls)', false);
        }
    }

    function handleInvoiceFile(file) {
        if (validateExcelFile(file)) {
            readExcelFile(file, (data) => {
                invoiceData = data;
                updateFileInfo(invoiceInfo, file.name, true);
            });
        } else {
            updateFileInfo(invoiceInfo, 'Invalid file format. Please use Excel files (.xlsx, .xls)', false);
        }
    }

    function validateExcelFile(file) {
        const validTypes = [
            'application/vnd.ms-excel',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        ];
        return validTypes.includes(file.type);
    }

    function readExcelFile(file, callback) {
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            callback(jsonData);
        };
        reader.readAsArrayBuffer(file);
    }

    function updateFileInfo(infoElement, message, success) {
        infoElement.textContent = message;
        infoElement.className = 'file-info ' + (success ? 'success' : 'error');
    }

    function analyzeFiles() {
        if (!vendorData || !invoiceData || !checkMonth.value) {
            alert('Please upload both files and select a month for analysis');
            return;
        }

        loading.classList.remove('hidden');

        setTimeout(() => {
            const selectedDate = new Date(checkMonth.value);
            const missingAccounts = findMissingAccounts(vendorData, invoiceData, selectedDate);
            displayResults(missingAccounts, selectedDate);
            loading.classList.add('hidden');
        }, 1000);
    }

    function findMissingAccounts(vendors, invoices, selectedDate) {
        const vendorAccounts = new Set(vendors.map(v => v['Account Number']?.toString()));
        const missingAccounts = [];
        const processedAccounts = new Set();

        invoices.forEach(invoice => {
            const invoiceDate = new Date(invoice['Document Date']);
            const accountNumber = invoice['Account Number']?.toString();

            if (invoiceDate.getMonth() === selectedDate.getMonth() && 
                invoiceDate.getFullYear() === selectedDate.getFullYear()) {
                if (!vendorAccounts.has(accountNumber) && !processedAccounts.has(accountNumber)) {
                    missingAccounts.push({
                        accountNumber: accountNumber,
                        amount: invoice['Amount'],
                        lastTransaction: invoice['Document Date']
                    });
                    processedAccounts.add(accountNumber);
                }
            }
        });

        return missingAccounts;
    }

    function displayResults(missingAccounts, selectedDate) {
        resultsBody.innerHTML = '';
        missingCount.textContent = missingAccounts.length;
        analysisMonth.textContent = selectedDate.toLocaleString('default', { month: 'long', year: 'numeric' });

        missingAccounts.forEach(account => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${account.accountNumber}</td>
                <td>N/A</td>
                <td>${new Date(account.lastTransaction).toLocaleDateString()}</td>
            `;
            resultsBody.appendChild(row);
        });

        results.classList.remove('hidden');
    }

    function exportResults() {
        const rows = [['Account Number', 'Last Transaction']];
        
        Array.from(resultsBody.children).forEach(row => {
            rows.push([
                row.children[0].textContent,
                row.children[2].textContent
            ]);
        });

        const ws = XLSX.utils.aoa_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Missing Accounts');
        
        const fileName = `missing_accounts_${checkMonth.value}.xlsx`;
        XLSX.writeFile(wb, fileName);
    }
});
