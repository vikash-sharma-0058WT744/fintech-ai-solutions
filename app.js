// Global variables
let extractedData = null;
let charts = {};

// Initialize app
document.addEventListener('DOMContentLoaded', function() {
    console.log('App initialized');
    initializeUpload();
});

// Initialize upload functionality
function initializeUpload() {
    const fileInput = document.getElementById('fileInput');
    const uploadBtn = document.getElementById('uploadBtn');
    const uploadArea = document.getElementById('uploadArea');

    uploadBtn.addEventListener('click', () => fileInput.click());
    uploadArea.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', handleFileSelect);

    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('dragover');
    });

    uploadArea.addEventListener('dragleave', () => {
        uploadArea.classList.remove('dragover');
    });

    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            fileInput.files = files;
            handleFileSelect({ target: { files: files } });
        }
    });

    document.getElementById('downloadBtn').addEventListener('click', generateDocx);
}

// Handle file selection
function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;

    console.log('File selected:', file.name);
    document.getElementById('fileName').textContent = `Selected: ${file.name}`;
    document.getElementById('loadingIndicator').style.display = 'block';
    document.getElementById('dataSection').style.display = 'none';

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            console.log('Workbook loaded:', workbook.SheetNames);
            processExcelData(workbook);
        } catch (error) {
            console.error('Error reading file:', error);
            alert('Error reading Excel file: ' + error.message);
            document.getElementById('loadingIndicator').style.display = 'none';
        }
    };
    reader.readAsArrayBuffer(file);
}

// Process Excel data
function processExcelData(workbook) {
    try {
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        console.log('Total rows:', jsonData.length);
        console.log('Raw Excel Data:', jsonData);

        extractedData = {
            companyName: '',
            cin: '',
            pan: '',
            lei: '335800LYXG6JYBGK1K19',
            incorporationDate: '',
            officeAddress: '',
            activity: '',
            rating: [],
            directors: [],
            shareholding: [],
            keyStrengths: '',
            aboutCompany: '',
            directorProfiles: '',
            financials: { headers: [], data: {} },
            debtProfile: getDefaultDebtData()
        };

        // Parse each row - data is in column B (index 1)
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.length < 2) continue;

            const cellB = row[1] ? String(row[1]).trim() : '';
            const cellC = row[2] ? String(row[2]).trim() : '';

            console.log(`Row ${i}: CellB="${cellB.substring(0, 50)}..."`);

            // Row 1: Company Name (skip, will get from About)
            if (i === 1) continue;
            
            // Row 2: CIN
            if (i === 2) {
                extractedData.cin = cellB;
                console.log('CIN:', cellB);
            }
            
            // Row 3: PAN
            if (i === 3) {
                extractedData.pan = cellB;
                console.log('PAN:', cellB);
            }
            
            // Row 4: Date (Excel serial number)
            if (i === 4) {
                const dateNum = parseFloat(cellB);
                if (!isNaN(dateNum)) {
                    const date = new Date((dateNum - 25569) * 86400 * 1000);
                    extractedData.incorporationDate = date.toISOString().split('T')[0];
                } else {
                    extractedData.incorporationDate = cellB;
                }
                console.log('Date:', extractedData.incorporationDate);
            }
            
            // Row 5: Office Address
            if (i === 5) {
                extractedData.officeAddress = cellB;
                console.log('Office:', cellB.substring(0, 50));
            }
            
            // Row 6: Activity
            if (i === 6) {
                extractedData.activity = cellB;
                console.log('Activity:', cellB);
            }
            
            // Row 7: Rating (contains table with tabs and newlines)
            if (i === 7) {
                const lines = cellB.split('\n');
                for (let j = 1; j < lines.length; j++) {
                    const parts = lines[j].split('\t');
                    if (parts.length >= 4) {
                        extractedData.rating.push({
                            instrument: parts[0].trim(),
                            amount: parts[1].trim(),
                            rating: parts[2].trim(),
                            action: parts[3].trim()
                        });
                    }
                }
                console.log('Rating:', extractedData.rating);
            }
            
            // Row 8: Board of Directors (contains table)
            if (i === 8) {
                const lines = cellB.split('\n');
                for (let j = 1; j < lines.length; j++) {
                    const parts = lines[j].split('\t');
                    if (parts.length >= 3) {
                        extractedData.directors.push({
                            name: parts[0].trim(),
                            din: parts[1].trim(),
                            designation: parts[2].trim()
                        });
                    }
                }
                console.log('Directors:', extractedData.directors.length);
            }
            
            // Row 9: Shareholding Pattern (contains table)
            if (i === 9) {
                const lines = cellB.split('\n');
                for (let j = 1; j < lines.length; j++) {
                    const parts = lines[j].split('\t');
                    if (parts.length >= 3 && !lines[j].includes('TOTAL')) {
                        extractedData.shareholding.push({
                            category: parts[0].trim(),
                            shares: parts[1].trim(),
                            percentage: parts[2].trim()
                        });
                    }
                }
                console.log('Shareholding:', extractedData.shareholding.length);
            }
            
            // Row 10: Key Strengths
            if (i === 10) {
                extractedData.keyStrengths = cellB;
                console.log('Key Strengths found');
            }
            
            // Row 11: About the Company
            if (i === 11) {
                extractedData.aboutCompany = cellB;
                // Extract company name
                if (cellB.includes('(')) {
                    extractedData.companyName = cellB.split('(')[0].trim();
                }
                console.log('Company Name:', extractedData.companyName);
            }
            
            // Row 12: Director Profiles
            if (i === 12) {
                extractedData.directorProfiles = cellB;
                console.log('Director Profiles found');
            }
            
            // Row 13: Financials (contains table)
            if (i === 13) {
                const lines = cellB.split('\n');
                if (lines.length > 0) {
                    // First line has headers
                    const headerParts = lines[0].split('\t');
                    for (let j = 1; j < headerParts.length; j++) {
                        if (headerParts[j]) {
                            extractedData.financials.headers.push(headerParts[j].trim());
                        }
                    }
                    
                    // Remaining lines have data
                    for (let j = 1; j < lines.length; j++) {
                        const parts = lines[j].split('\t');
                        if (parts.length > 1) {
                            const particular = parts[0].trim();
                            if (particular) {
                                extractedData.financials.data[particular] = {};
                                for (let k = 0; k < extractedData.financials.headers.length; k++) {
                                    extractedData.financials.data[particular][extractedData.financials.headers[k]] = parts[k + 1] || '';
                                }
                            }
                        }
                    }
                }
                console.log('Financials:', Object.keys(extractedData.financials.data).length, 'rows');
            }
        }

        console.log('Extracted Data:', extractedData);
        displayData();
    } catch (error) {
        console.error('Error processing Excel:', error);
        alert('Error processing data: ' + error.message);
        document.getElementById('loadingIndicator').style.display = 'none';
    }
}

// Get default debt data
function getDefaultDebtData() {
    return {
        termLoan: [
            { lender: 'STATE BANK OF INDIA', sanctioned: '880.00', outstanding: '388.58' },
            { lender: 'BANK OF MAHARASHTRA', sanctioned: '450.00', outstanding: '337.33' },
            { lender: 'HDFC BANK LIMITED', sanctioned: '660.00', outstanding: '312.02' },
            { lender: 'BANDHAN BANK', sanctioned: '355.00', outstanding: '292.00' },
            { lender: 'THE FEDERAL BANK LIMITED', sanctioned: '425.00', outstanding: '281.99' }
        ],
        ptc: [
            { lender: 'Indusind Bank Limited', sanctioned: '265.96', outstanding: '234.87' },
            { lender: 'Axis Bank Limited', sanctioned: '114.88', outstanding: '96.14' }
        ],
        da: [
            { lender: 'State Bank of India', sanctioned: '889.80', outstanding: '631.91' },
            { lender: 'DBS Bank', sanctioned: '278.75', outstanding: '220.56' }
        ],
        nhb: [
            { lender: 'National Housing Bank', sanctioned: '1540.00', outstanding: '900.13' }
        ],
        ecb: [
            { lender: 'US International Development Finance Corporation', sanctioned: '243.51', outstanding: '243.51' }
        ],
        ncd: [
            { lender: 'Kotak Bank Limited', sanctioned: '50.00', outstanding: '50.00' }
        ]
    };
}

// Display data
function displayData() {
    console.log('Displaying data...');
    document.getElementById('loadingIndicator').style.display = 'none';
    document.getElementById('dataSection').style.display = 'block';

    displayCompanyInfo();
    displayMetrics();
    displayCharts();
    displayTables();
    displayKeyStrengths();
    displayDebtProfile();
}

// Display company info
function displayCompanyInfo() {
    const companyInfo = document.getElementById('companyInfo');
    companyInfo.innerHTML = `
        <div class="info-item">
            <div class="info-label">Company Name</div>
            <div class="info-value">${extractedData.companyName || 'N/A'}</div>
        </div>
        <div class="info-item">
            <div class="info-label">CIN Number</div>
            <div class="info-value">${extractedData.cin || 'N/A'}</div>
        </div>
        <div class="info-item">
            <div class="info-label">PAN Number</div>
            <div class="info-value">${extractedData.pan || 'N/A'}</div>
        </div>
        <div class="info-item">
            <div class="info-label">LEI Number</div>
            <div class="info-value">${extractedData.lei}</div>
        </div>
        <div class="info-item">
            <div class="info-label">Date of Incorporation</div>
            <div class="info-value">${extractedData.incorporationDate || 'N/A'}</div>
        </div>
        <div class="info-item">
            <div class="info-label">Line of Activity</div>
            <div class="info-value">${extractedData.activity || 'N/A'}</div>
        </div>
    `;
    console.log('Company info displayed');
}

// Display metrics
function displayMetrics() {
    const ratingValue = extractedData.rating.length > 0 ? extractedData.rating[0].rating : 'N/A';
    document.getElementById('ratingValue').textContent = ratingValue;

    document.getElementById('directorsCount').textContent = extractedData.directors.length;

    const tnw = extractedData.financials.data['TNW'];
    const netWorth = tnw && extractedData.financials.headers.length > 0 ? 
        tnw[extractedData.financials.headers[extractedData.financials.headers.length - 1]] : 'N/A';
    document.getElementById('netWorth').textContent = netWorth ? `₹${netWorth} Cr` : 'N/A';

    const totalDebt = calculateTotalDebt();
    document.getElementById('totalDebt').textContent = `₹${totalDebt.toFixed(2)} Cr`;
    
    console.log('Metrics displayed');
}

// Calculate total debt
function calculateTotalDebt() {
    let total = 0;
    const debt = extractedData.debtProfile;
    
    ['termLoan', 'ptc', 'da', 'nhb', 'ecb', 'ncd'].forEach(category => {
        if (debt[category]) {
            debt[category].forEach(item => {
                const outstanding = parseFloat(item.outstanding.replace(/,/g, ''));
                if (!isNaN(outstanding)) total += outstanding;
            });
        }
    });
    
    return total;
}

// Display charts
function displayCharts() {
    displayFinancialChart();
    displayShareholdingChart();
    displayDebtChart();
    console.log('Charts displayed');
}

// Display financial chart
function displayFinancialChart() {
    const ctx = document.getElementById('financialChart');
    if (!ctx) return;
    
    if (charts.financial) charts.financial.destroy();

    const headers = extractedData.financials.headers;
    const tnwData = [];
    const patData = [];

    headers.forEach(year => {
        const tnw = extractedData.financials.data['TNW'];
        const pat = extractedData.financials.data['Profit After Tax'];
        
        if (tnw) tnwData.push(parseFloat(tnw[year].replace(/,/g, '')) || 0);
        if (pat) patData.push(parseFloat(pat[year].replace(/,/g, '')) || 0);
    });

    charts.financial = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: headers,
            datasets: [
                {
                    label: 'Net Worth (TNW)',
                    data: tnwData,
                    backgroundColor: 'rgba(37, 99, 235, 0.7)',
                    borderColor: 'rgba(37, 99, 235, 1)',
                    borderWidth: 2
                },
                {
                    label: 'Profit After Tax',
                    data: patData,
                    backgroundColor: 'rgba(16, 185, 129, 0.7)',
                    borderColor: 'rgba(16, 185, 129, 1)',
                    borderWidth: 2
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Amount (in Crores)'
                    }
                }
            }
        }
    });
}

// Display shareholding chart
function displayShareholdingChart() {
    const ctx = document.getElementById('shareholdingChart');
    if (!ctx) return;
    
    if (charts.shareholding) charts.shareholding.destroy();

    const labels = extractedData.shareholding.map(s => s.category);
    const data = extractedData.shareholding.map(s => {
        const percentage = String(s.percentage).replace('%', '');
        return parseFloat(percentage) || 0;
    });

    const colors = [
        'rgba(37, 99, 235, 0.8)',
        'rgba(16, 185, 129, 0.8)',
        'rgba(245, 158, 11, 0.8)',
        'rgba(239, 68, 68, 0.8)',
        'rgba(139, 92, 246, 0.8)',
        'rgba(236, 72, 153, 0.8)',
        'rgba(14, 165, 233, 0.8)'
    ];

    charts.shareholding = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: labels,
            datasets: [{
                data: data,
                backgroundColor: colors,
                borderWidth: 2,
                borderColor: '#fff'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    position: 'right',
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return context.label + ': ' + context.parsed + '%';
                        }
                    }
                }
            }
        }
    });
}

// Display debt chart
function displayDebtChart() {
    const ctx = document.getElementById('debtChart');
    if (!ctx) return;
    
    if (charts.debt) charts.debt.destroy();

    const debt = extractedData.debtProfile;
    const categories = ['Term Loan', 'PTC', 'DA', 'NHB Refinance', 'ECB', 'NCD'];
    const categoryKeys = ['termLoan', 'ptc', 'da', 'nhb', 'ecb', 'ncd'];
    
    const sanctionedData = [];
    const outstandingData = [];

    categoryKeys.forEach(key => {
        let sanctioned = 0;
        let outstanding = 0;
        
        if (debt[key]) {
            debt[key].forEach(item => {
                sanctioned += parseFloat(item.sanctioned.replace(/,/g, '')) || 0;
                outstanding += parseFloat(item.outstanding.replace(/,/g, '')) || 0;
            });
        }
        
        sanctionedData.push(sanctioned);
        outstandingData.push(outstanding);
    });

    charts.debt = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: categories,
            datasets: [
                {
                    label: 'Sanctioned Amount',
                    data: sanctionedData,
                    backgroundColor: 'rgba(37, 99, 235, 0.7)',
                    borderColor: 'rgba(37, 99, 235, 1)',
                    borderWidth: 2
                },
                {
                    label: 'Outstanding Amount',
                    data: outstandingData,
                    backgroundColor: 'rgba(245, 158, 11, 0.7)',
                    borderColor: 'rgba(245, 158, 11, 1)',
                    borderWidth: 2
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Amount (in Crores)'
                    }
                }
            }
        }
    });
}

// Display tables
function displayTables() {
    const directorsTable = document.getElementById('directorsTable');
    let directorsHTML = `
        <thead>
            <tr>
                <th>Name</th>
                <th>DIN/PAN No</th>
                <th>Designation</th>
            </tr>
        </thead>
        <tbody>
    `;
    
    if (extractedData.directors.length === 0) {
        directorsHTML += `<tr><td colspan="3" style="text-align: center;">No director data available</td></tr>`;
    } else {
        extractedData.directors.forEach(director => {
            directorsHTML += `
                <tr>
                    <td>${director.name}</td>
                    <td>${director.din}</td>
                    <td>${director.designation}</td>
                </tr>
            `;
        });
    }
    
    directorsHTML += '</tbody>';
    directorsTable.innerHTML = directorsHTML;

    const shareholdingTable = document.getElementById('shareholdingTable');
    let shareholdingHTML = `
        <thead>
            <tr>
                <th>Category of Shareholder</th>
                <th>No. of Shares held</th>
                <th>Percentage</th>
            </tr>
        </thead>
        <tbody>
    `;
    
    if (extractedData.shareholding.length === 0) {
        shareholdingHTML += `<tr><td colspan="3" style="text-align: center;">No shareholding data available</td></tr>`;
    } else {
        extractedData.shareholding.forEach(sh => {
            shareholdingHTML += `
                <tr>
                    <td>${sh.category}</td>
                    <td>${sh.shares}</td>
                    <td>${sh.percentage}</td>
                </tr>
            `;
        });
    }
    
    shareholdingHTML += '</tbody>';
    shareholdingTable.innerHTML = shareholdingHTML;
    
    console.log('Tables displayed');
}

// Display key strengths
function displayKeyStrengths() {
    const strengthsDiv = document.getElementById('keyStrengths');
    
    if (!extractedData.keyStrengths) {
        strengthsDiv.innerHTML = '<p>No key strengths data available</p>';
        return;
    }
    
    const strengths = extractedData.keyStrengths.split('•').filter(s => s.trim());
    
    let html = '';
    strengths.forEach(strength => {
        if (strength.trim()) {
            html += `
                <div class="strength-item">
                    <div class="strength-icon">✓</div>
                    <div class="strength-text">${strength.trim()}</div>
                </div>
            `;
        }
    });
    
    strengthsDiv.innerHTML = html || '<p>No key strengths data available</p>';
    console.log('Key strengths displayed');
}

// Display debt profile
function displayDebtProfile() {
    const debtTablesDiv = document.getElementById('debtTables');
    const debt = extractedData.debtProfile;
    
    const categories = [
        { key: 'termLoan', title: 'Term Loan' },
        { key: 'ptc', title: 'Pass Through Certificate (PTC)' },
        { key: 'da', title: 'Direct Assignment (DA)' },
        { key: 'nhb', title: 'NHB Refinance' },
        { key: 'ecb', title: 'External Commercial Borrowing (ECB)' },
        { key: 'ncd', title: 'Non-Convertible Debentures (NCD)' }
    ];
    
    let html = '';
    
    categories.forEach(category => {
        if (debt[category.key] && debt[category.key].length > 0) {
            html += `
                <div class="debt-table-container">
                    <h3>${category.title}</h3>
                    <table class="debt-table">
                        <thead>
                            <tr>
                                <th>Lender</th>
                                <th>Sanctioned (₹ Cr)</th>
                                <th>Outstanding (₹ Cr)</th>
                            </tr>
                        </thead>
                        <tbody>
            `;
            
            debt[category.key].forEach(item => {
                html += `
                    <tr>
                        <td>${item.lender}</td>
                        <td>${item.sanctioned}</td>
                        <td>${item.outstanding}</td>
                    </tr>
                `;
            });
            
            html += `
                        </tbody>
                    </table>
                </div>
            `;
        }
    });
    
    debtTablesDiv.innerHTML = html;
    console.log('Debt profile displayed');
}

// Generate DOCX - Simple text file
function generateDocx() {
    if (!extractedData) {
        alert('Please upload and process an Excel file first.');
        return;
    }

    try {
        console.log('Generating document...');
        
        let content = `TERM LOAN TEASER\n\n`;
        content += `${extractedData.companyName || 'Company Name'}\n\n`;
        content += `COMPANY INFORMATION\n`;
        content += `CIN: ${extractedData.cin || 'N/A'}\n`;
        content += `PAN: ${extractedData.pan || 'N/A'}\n`;
        content += `Date of Incorporation: ${extractedData.incorporationDate || 'N/A'}\n`;
        content += `Line of Activity: ${extractedData.activity || 'N/A'}\n`;
        content += `Rating: ${extractedData.rating.length > 0 ? extractedData.rating[0].rating : 'N/A'}\n\n`;
        
        content += `KEY STRENGTHS\n`;
        content += `${extractedData.keyStrengths || 'No key strengths available'}\n\n`;
        
        content += `ABOUT THE COMPANY\n`;
        content += `${extractedData.aboutCompany || 'No company information available'}\n\n`;
        
        content += `BOARD OF DIRECTORS\n`;
        extractedData.directors.forEach(dir => {
            content += `${dir.name} - ${dir.designation}\n`;
        });
        
        const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
        const fileName = `${extractedData.companyName || 'Company'}_Term_Loan_Teaser.txt`;
        saveAs(blob, fileName);
        
        console.log('Document generated successfully');

    } catch (error) {
        console.error('Error generating document:', error);
        alert('Error generating document: ' + error.message);
    }
}

// Made with Bob
