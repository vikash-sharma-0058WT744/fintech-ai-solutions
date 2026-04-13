// Global variables
let extractedData = null;
let charts = {};

// Initialize app
document.addEventListener('DOMContentLoaded', function() {
    initializeUpload();
});

// Initialize upload functionality
function initializeUpload() {
    const fileInput = document.getElementById('fileInput');
    const uploadBtn = document.getElementById('uploadBtn');
    const uploadArea = document.getElementById('uploadArea');
    const fileName = document.getElementById('fileName');

    // Button click
    uploadBtn.addEventListener('click', () => fileInput.click());
    
    // Upload area click
    uploadArea.addEventListener('click', () => fileInput.click());

    // File selection
    fileInput.addEventListener('change', handleFileSelect);

    // Drag and drop
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

    // Download button
    document.getElementById('downloadBtn').addEventListener('click', generateDocx);
}

// Handle file selection
function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;

    document.getElementById('fileName').textContent = `Selected: ${file.name}`;
    document.getElementById('loadingIndicator').style.display = 'block';
    document.getElementById('dataSection').style.display = 'none';

    // Read Excel file
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            processExcelData(workbook);
        } catch (error) {
            console.error('Error reading file:', error);
            alert('Error reading Excel file. Please ensure it\'s a valid .xlsx file.');
            document.getElementById('loadingIndicator').style.display = 'none';
        }
    };
    reader.readAsArrayBuffer(file);
}

// Process Excel data
function processExcelData(workbook) {
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

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

    // Parse data (Column B = headings, Column C = data)
    let currentSection = '';
    let i = 0;
    
    while (i < jsonData.length) {
        const row = jsonData[i];
        const heading = row[1] ? String(row[1]).trim() : '';
        const data = row[2] ? String(row[2]).trim() : '';

        console.log(`Row ${i}: Heading="${heading}", Data="${data}"`);

        if (heading === 'Applicant Company') {
            currentSection = 'company';
            i++;
        } else if (heading === 'CIN NO') {
            extractedData.cin = data;
            i++;
        } else if (heading === 'PAN No') {
            extractedData.pan = data;
            i++;
        } else if (heading === 'Date of Incorporation') {
            extractedData.incorporationDate = formatDate(data);
            i++;
        } else if (heading === 'Registered and  Corporate  Office') {
            extractedData.officeAddress = data;
            i++;
        } else if (heading === 'Line of Activity') {
            extractedData.activity = data;
            i++;
        } else if (heading === 'Rating') {
            currentSection = 'rating';
            i++;
            // Parse rating rows
            while (i < jsonData.length && jsonData[i][1] !== undefined && String(jsonData[i][1]).trim() !== '' && !isNewSection(jsonData[i][1])) {
                const ratingRow = jsonData[i];
                if (ratingRow[1] && String(ratingRow[1]).includes('Long Term')) {
                    extractedData.rating.push({
                        instrument: String(ratingRow[1] || '').trim(),
                        amount: String(ratingRow[2] || '').trim(),
                        rating: String(ratingRow[3] || '').trim(),
                        action: String(ratingRow[4] || '').trim()
                    });
                }
                i++;
            }
        } else if (heading === 'Board of Directors') {
            currentSection = 'directors';
            i++;
            // Parse director rows
            while (i < jsonData.length && jsonData[i][1] !== undefined && String(jsonData[i][1]).trim() !== '' && !isNewSection(jsonData[i][1])) {
                const dirRow = jsonData[i];
                const name = String(dirRow[1] || '').trim();
                if (name && (name.startsWith('Mr.') || name.startsWith('Ms.'))) {
                    extractedData.directors.push({
                        name: name,
                        din: String(dirRow[2] || '').trim(),
                        designation: String(dirRow[3] || '').trim()
                    });
                }
                i++;
            }
        } else if (heading === 'Shareholding Pattern') {
            currentSection = 'shareholding';
            i++;
            // Parse shareholding rows
            while (i < jsonData.length && jsonData[i][1] !== undefined && String(jsonData[i][1]).trim() !== '' && !isNewSection(jsonData[i][1])) {
                const shRow = jsonData[i];
                const category = String(shRow[1] || '').trim();
                if (category && category !== 'Category of Shareholder' && category !== 'TOTAL') {
                    extractedData.shareholding.push({
                        category: category,
                        shares: String(shRow[2] || '').trim(),
                        percentage: String(shRow[3] || '').trim()
                    });
                }
                i++;
            }
        } else if (heading === 'Key Strengths') {
            extractedData.keyStrengths = data;
            i++;
        } else if (heading === 'About the Company') {
            extractedData.aboutCompany = data;
            // Extract company name from about section
            if (data && data.includes('(')) {
                extractedData.companyName = data.split('(')[0].trim();
            }
            i++;
        } else if (heading === 'Director & Promoter Profile') {
            extractedData.directorProfiles = data;
            i++;
        } else if (heading.includes('Financials')) {
            currentSection = 'financials';
            i++;
            // Parse financial data
            if (i < jsonData.length) {
                const headerRow = jsonData[i];
                // Extract year headers (FY 22, FY 23, etc.)
                for (let j = 2; j < headerRow.length; j++) {
                    if (headerRow[j]) {
                        extractedData.financials.headers.push(String(headerRow[j]).trim());
                    }
                }
                i++;
                
                // Parse financial rows
                while (i < jsonData.length && jsonData[i][1] !== undefined && String(jsonData[i][1]).trim() !== '' && !isNewSection(jsonData[i][1])) {
                    const finRow = jsonData[i];
                    const particular = String(finRow[1] || '').trim();
                    if (particular && particular !== 'Particulars') {
                        extractedData.financials.data[particular] = {};
                        for (let j = 0; j < extractedData.financials.headers.length; j++) {
                            extractedData.financials.data[particular][extractedData.financials.headers[j]] = String(finRow[j + 2] || '').trim();
                        }
                    }
                    i++;
                }
            }
        } else {
            i++;
        }
    }

    console.log('Extracted Data:', extractedData);

    // Display data
    displayData();
}

// Check if a heading is a new section
function isNewSection(heading) {
    const sections = [
        'Applicant Company', 'CIN NO', 'PAN No', 'Date of Incorporation',
        'Registered and  Corporate  Office', 'Line of Activity', 'Rating',
        'Board of Directors', 'Shareholding Pattern', 'Key Strengths',
        'About the Company', 'Director & Promoter Profile', 'Financials',
        'Portfolio Cuts', 'Debt Profile'
    ];
    const headingStr = String(heading).trim();
    return sections.some(section => headingStr.includes(section));
}

// Format date
function formatDate(date) {
    if (!date) return '';
    if (typeof date === 'string') {
        if (date.includes('00:00:00')) {
            return date.split(' ')[0];
        }
        return date;
    }
    return String(date);
}

// Get default debt data
function getDefaultDebtData() {
    return {
        termLoan: [
            { lender: 'STATE BANK OF INDIA', sanctioned: '880.00', outstanding: '388.58' },
            { lender: 'BANK OF MAHARASHTRA', sanctioned: '450.00', outstanding: '337.33' },
            { lender: 'HDFC BANK LIMITED', sanctioned: '660.00', outstanding: '312.02' },
            { lender: 'BANDHAN BANK', sanctioned: '355.00', outstanding: '292.00' },
            { lender: 'THE FEDERAL BANK LIMITED', sanctioned: '425.00', outstanding: '281.99' },
            { lender: 'BANK OF BARODA', sanctioned: '250.00', outstanding: '218.46' },
            { lender: 'KOTAK BANK LIMITED', sanctioned: '375.00', outstanding: '207.49' },
            { lender: 'Canara Bank', sanctioned: '225.00', outstanding: '174.21' },
            { lender: 'BAJAJ FINANCE LIMITED', sanctioned: '250.00', outstanding: '170.46' },
            { lender: 'RBL BANK LIMITED', sanctioned: '225.00', outstanding: '160.52' }
        ],
        ptc: [
            { lender: 'Indusind Bank Limited', sanctioned: '265.96', outstanding: '234.87' },
            { lender: 'Axis Bank Limited', sanctioned: '114.88', outstanding: '96.14' },
            { lender: 'ICICI Bank Limited', sanctioned: '115.28', outstanding: '57.21' }
        ],
        da: [
            { lender: 'State Bank of India', sanctioned: '889.80', outstanding: '631.91' },
            { lender: 'DBS Bank', sanctioned: '278.75', outstanding: '220.56' },
            { lender: 'INDIAN BANK', sanctioned: '171.86', outstanding: '83.96' },
            { lender: 'Bank of Baroda', sanctioned: '195.82', outstanding: '62.02' }
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
    document.getElementById('loadingIndicator').style.display = 'none';
    document.getElementById('dataSection').style.display = 'block';

    // Company info
    displayCompanyInfo();
    
    // Metrics
    displayMetrics();
    
    // Charts
    displayCharts();
    
    // Tables
    displayTables();
    
    // Key Strengths
    displayKeyStrengths();
    
    // Debt Profile
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
}

// Display metrics
function displayMetrics() {
    // Rating
    const ratingValue = extractedData.rating.length > 0 ? extractedData.rating[0].rating : 'N/A';
    document.getElementById('ratingValue').textContent = ratingValue;

    // Directors count
    document.getElementById('directorsCount').textContent = extractedData.directors.length;

    // Net Worth (from financials FY25)
    const tnw = extractedData.financials.data['TNW'];
    const netWorth = tnw && extractedData.financials.headers.length > 0 ? 
        tnw[extractedData.financials.headers[extractedData.financials.headers.length - 1]] : 'N/A';
    document.getElementById('netWorth').textContent = netWorth ? `₹${netWorth} Cr` : 'N/A';

    // Total Debt
    const totalDebt = calculateTotalDebt();
    document.getElementById('totalDebt').textContent = `₹${totalDebt.toFixed(2)} Cr`;
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
}

// Display financial chart
function displayFinancialChart() {
    const ctx = document.getElementById('financialChart').getContext('2d');
    
    if (charts.financial) charts.financial.destroy();

    const headers = extractedData.financials.headers;
    const tnwData = [];
    const patData = [];

    headers.forEach(year => {
        const tnw = extractedData.financials.data['TNW'];
        const pat = extractedData.financials.data['Profit After Tax'];
        
        if (tnw) tnwData.push(parseFloat(tnw[year]) || 0);
        if (pat) patData.push(parseFloat(pat[year]) || 0);
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
            },
            plugins: {
                legend: {
                    display: true,
                    position: 'top'
                }
            }
        }
    });
}

// Display shareholding chart
function displayShareholdingChart() {
    const ctx = document.getElementById('shareholdingChart').getContext('2d');
    
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
    const ctx = document.getElementById('debtChart').getContext('2d');
    
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
            },
            plugins: {
                legend: {
                    display: true,
                    position: 'top'
                }
            }
        }
    });
}

// Display tables
function displayTables() {
    // Directors table
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
        directorsHTML += `
            <tr>
                <td colspan="3" style="text-align: center;">No director data available</td>
            </tr>
        `;
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

    // Shareholding table
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
        shareholdingHTML += `
            <tr>
                <td colspan="3" style="text-align: center;">No shareholding data available</td>
            </tr>
        `;
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
}

// Generate DOCX
async function generateDocx() {
    if (!extractedData) {
        alert('Please upload and process an Excel file first.');
        return;
    }

    try {
        const { Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, HeadingLevel } = docx;

        // Create document
        const doc = new Document({
            sections: [{
                properties: {},
                children: [
                    // Title
                    new Paragraph({
                        text: 'Term Loan Teaser',
                        heading: HeadingLevel.HEADING_1,
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 400 }
                    }),

                    // Company Name
                    new Paragraph({
                        text: extractedData.companyName || 'Company Name',
                        heading: HeadingLevel.HEADING_2,
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 400 }
                    }),

                    // Company Information
                    new Paragraph({
                        text: 'Company Information',
                        heading: HeadingLevel.HEADING_2,
                        spacing: { before: 400, after: 200 }
                    }),

                    new Paragraph({
                        children: [
                            new TextRun({ text: 'CIN: ', bold: true }),
                            new TextRun(extractedData.cin || 'N/A')
                        ],
                        spacing: { after: 100 }
                    }),

                    new Paragraph({
                        children: [
                            new TextRun({ text: 'PAN: ', bold: true }),
                            new TextRun(extractedData.pan || 'N/A')
                        ],
                        spacing: { after: 100 }
                    }),

                    new Paragraph({
                        children: [
                            new TextRun({ text: 'Date of Incorporation: ', bold: true }),
                            new TextRun(extractedData.incorporationDate || 'N/A')
                        ],
                        spacing: { after: 100 }
                    }),

                    new Paragraph({
                        children: [
                            new TextRun({ text: 'Line of Activity: ', bold: true }),
                            new TextRun(extractedData.activity || 'N/A')
                        ],
                        spacing: { after: 100 }
                    }),

                    new Paragraph({
                        children: [
                            new TextRun({ text: 'Rating: ', bold: true }),
                            new TextRun(extractedData.rating.length > 0 ? extractedData.rating[0].rating : 'N/A')
                        ],
                        spacing: { after: 400 }
                    }),

                    // Key Strengths
                    new Paragraph({
                        text: 'Key Strengths',
                        heading: HeadingLevel.HEADING_2,
                        spacing: { before: 400, after: 200 }
                    }),

                    new Paragraph({
                        text: extractedData.keyStrengths || 'No key strengths available',
                        spacing: { after: 400 }
                    }),

                    // About Company
                    new Paragraph({
                        text: 'About the Company',
                        heading: HeadingLevel.HEADING_2,
                        spacing: { before: 400, after: 200 }
                    }),

                    new Paragraph({
                        text: extractedData.aboutCompany || 'No company information available',
                        spacing: { after: 400 }
                    })
                ]
            }]
        });

        // Generate and download
        const blob = await Packer.toBlob(doc);
        const fileName = `${extractedData.companyName || 'Company'}_Term_Loan_Teaser.docx`;
        saveAs(blob, fileName);

    } catch (error) {
        console.error('Error generating DOCX:', error);
        alert('Error generating document. Please try again.');
    }
}

// Made with Bob
