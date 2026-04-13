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
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

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
    for (let i = 2; i < jsonData.length; i++) {
        const heading = jsonData[i][1];
        const data = jsonData[i][2];

        if (!heading) continue;

        const headingStr = String(heading).trim();

        if (headingStr === 'Applicant Company') {
            // Will be extracted from About Company
        } else if (headingStr === 'CIN NO') {
            extractedData.cin = data || '';
        } else if (headingStr === 'PAN No') {
            extractedData.pan = data || '';
        } else if (headingStr === 'Date of Incorporation') {
            extractedData.incorporationDate = formatDate(data);
        } else if (headingStr === 'Registered and  Corporate  Office') {
            extractedData.officeAddress = data || '';
        } else if (headingStr === 'Line of Activity') {
            extractedData.activity = data || '';
        } else if (headingStr === 'Rating') {
            extractedData.rating = parseRating(data);
        } else if (headingStr === 'Board of Directors') {
            extractedData.directors = parseDirectors(data);
        } else if (headingStr === 'Shareholding Pattern') {
            extractedData.shareholding = parseShareholding(data);
        } else if (headingStr === 'Key Strengths') {
            extractedData.keyStrengths = data || '';
        } else if (headingStr === 'About the Company') {
            extractedData.aboutCompany = data || '';
            // Extract company name
            if (data && data.includes('(')) {
                extractedData.companyName = data.split('(')[0].trim();
            }
        } else if (headingStr === 'Director & Promoter Profile') {
            extractedData.directorProfiles = data || '';
        } else if (headingStr.startsWith('Financials')) {
            extractedData.financials = parseFinancials(data);
        }
    }

    // Display data
    displayData();
}

// Parse rating data
function parseRating(data) {
    if (!data) return [];
    const lines = String(data).split('\n');
    const ratings = [];
    for (const line of lines) {
        if (line.includes('Long Term')) {
            const parts = line.split('\t');
            if (parts.length >= 4) {
                ratings.push({
                    instrument: parts[0].trim(),
                    amount: parts[1].trim(),
                    rating: parts[2].trim(),
                    action: parts[3].trim()
                });
            }
        }
    }
    return ratings;
}

// Parse directors data
function parseDirectors(data) {
    if (!data) return [];
    const lines = String(data).split('\n');
    const directors = [];
    for (const line of lines) {
        if (line.startsWith('Mr.') || line.startsWith('Ms.')) {
            const parts = line.split('\t');
            if (parts.length >= 3) {
                directors.push({
                    name: parts[0].trim(),
                    din: parts[1].trim(),
                    designation: parts[2].trim()
                });
            }
        }
    }
    return directors;
}

// Parse shareholding data
function parseShareholding(data) {
    if (!data) return [];
    const lines = String(data).split('\n');
    const shareholding = [];
    for (const line of lines) {
        const parts = line.split('\t');
        if (parts.length >= 3 && !line.includes('Category of Shareholder') && !line.includes('TOTAL')) {
            const category = parts[0].trim();
            if (category) {
                shareholding.push({
                    category: category,
                    shares: parts[1].trim(),
                    percentage: parts[2].trim()
                });
            }
        }
    }
    return shareholding;
}

// Parse financials data
function parseFinancials(data) {
    if (!data) return { headers: [], data: {} };
    const lines = String(data).split('\n');
    const headers = [];
    const financialData = {};

    if (lines.length > 0) {
        const headerParts = lines[0].split('\t');
        for (let i = 1; i < headerParts.length; i++) {
            if (headerParts[i]) headers.push(headerParts[i].trim());
        }
    }

    for (let i = 1; i < lines.length; i++) {
        const parts = lines[i].split('\t');
        if (parts.length > 1) {
            const key = parts[0].trim();
            if (key && key !== 'Particulars') {
                financialData[key] = {};
                for (let j = 0; j < headers.length; j++) {
                    financialData[key][headers[j]] = parts[j + 1] || '';
                }
            }
        }
    }

    return { headers, data: financialData };
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
            plugins: {
                legend: {
                    position: 'top',
                },
                title: {
                    display: false
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Amount (Rs. in Crores)'
                    }
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
            plugins: {
                legend: {
                    position: 'top',
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Amount (Rs. in Crores)'
                    }
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
    
    extractedData.directors.forEach(director => {
        directorsHTML += `
            <tr>
                <td>${director.name}</td>
                <td>${director.din}</td>
                <td>${director.designation}</td>
            </tr>
        `;
    });
    
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
    
    extractedData.shareholding.forEach(sh => {
        shareholdingHTML += `
            <tr>
                <td>${sh.category}</td>
                <td>${sh.shares}</td>
                <td>${sh.percentage}</td>
            </tr>
        `;
    });
    
    shareholdingHTML += '</tbody>';
    shareholdingTable.innerHTML = shareholdingHTML;
}

// Display key strengths
function displayKeyStrengths() {
    const strengthsDiv = document.getElementById('keyStrengths');
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
    
    strengthsDiv.innerHTML = html;
}

// Display debt profile
function displayDebtProfile() {
    const debtTablesDiv = document.getElementById('debtTables');
    const debt = extractedData.debtProfile;
    
    const categories = [
        { key: 'termLoan', title: 'Term Loan' },
        { key: 'ptc', title: 'PTC' },
        { key: 'da', title: 'DA (Direct Assignment)' },
        { key: 'nhb', title: 'NHB Refinance' },
        { key: 'ecb', title: 'ECB (External Commercial Borrowing)' },
        { key: 'ncd', title: 'NCD (Non-Convertible Debentures)' }
    ];
    
    let html = '';
    
    categories.forEach(category => {
        if (debt[category.key] && debt[category.key].length > 0) {
            html += `
                <div class="debt-category">
                    <div class="debt-category-header">${category.title}</div>
                    <table class="debt-category-table">
                        <thead>
                            <tr>
                                <th>Lender Name</th>
                                <th>Sanctioned Amt</th>
                                <th>Outstanding Amt</th>
                            </tr>
                        </thead>
                        <tbody>
            `;
            
            let totalSanctioned = 0;
            let totalOutstanding = 0;
            
            debt[category.key].forEach(item => {
                html += `
                    <tr>
                        <td>${item.lender}</td>
                        <td>${item.sanctioned}</td>
                        <td>${item.outstanding}</td>
                    </tr>
                `;
                totalSanctioned += parseFloat(item.sanctioned.replace(/,/g, '')) || 0;
                totalOutstanding += parseFloat(item.outstanding.replace(/,/g, '')) || 0;
            });
            
            html += `
                    <tr class="debt-total">
                        <td>TOTAL</td>
                        <td>${totalSanctioned.toFixed(2)}</td>
                        <td>${totalOutstanding.toFixed(2)}</td>
                    </tr>
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
        alert('Please upload an Excel file first');
        return;
    }

    try {
        const doc = new docx.Document({
            sections: [{
                properties: {},
                children: [
                    // Title
                    new docx.Paragraph({
                        text: "Term Loan Teaser",
                        heading: docx.HeadingLevel.TITLE,
                        alignment: docx.AlignmentType.CENTER,
                        spacing: { after: 400 }
                    }),
                    
                    // Company Name
                    new docx.Paragraph({
                        text: extractedData.companyName || "Company Name",
                        heading: docx.HeadingLevel.HEADING_1,
                        alignment: docx.AlignmentType.CENTER,
                        spacing: { after: 400 }
                    }),
                    
                    // Executive Summary
                    new docx.Paragraph({
                        text: "Executive Summary",
                        heading: docx.HeadingLevel.HEADING_2,
                        spacing: { before: 400, after: 200 }
                    }),
                    
                    new docx.Paragraph({
                        text: `Company: ${extractedData.companyName}`,
                        spacing: { after: 100 }
                    }),
                    new docx.Paragraph({
                        text: `CIN: ${extractedData.cin}`,
                        spacing: { after: 100 }
                    }),
                    new docx.Paragraph({
                        text: `PAN: ${extractedData.pan}`,
                        spacing: { after: 100 }
                    }),
                    new docx.Paragraph({
                        text: `LEI: ${extractedData.lei}`,
                        spacing: { after: 100 }
                    }),
                    new docx.Paragraph({
                        text: `Date of Incorporation: ${extractedData.incorporationDate}`,
                        spacing: { after: 100 }
                    }),
                    new docx.Paragraph({
                        text: `Office: ${extractedData.officeAddress}`,
                        spacing: { after: 100 }
                    }),
                    new docx.Paragraph({
                        text: `Activity: ${extractedData.activity}`,
                        spacing: { after: 400 }
                    }),
                    
                    // About Company
                    new docx.Paragraph({
                        text: "About the Company",
                        heading: docx.HeadingLevel.HEADING_2,
                        spacing: { before: 400, after: 200 }
                    }),
                    new docx.Paragraph({
                        text: extractedData.aboutCompany || "",
                        spacing: { after: 400 }
                    }),
                    
                    // Key Strengths
                    new docx.Paragraph({
                        text: "Key Strengths",
                        heading: docx.HeadingLevel.HEADING_2,
                        spacing: { before: 400, after: 200 }
                    }),
                    ...extractedData.keyStrengths.split('•').filter(s => s.trim()).map(strength => 
                        new docx.Paragraph({
                            text: strength.trim(),
                            bullet: { level: 0 },
                            spacing: { after: 100 }
                        })
                    ),
                    
                    // Note about full details
                    new docx.Paragraph({
                        text: "\nFor complete details including Directors, Shareholding, Financials, and Debt Profile, please refer to the web interface.",
                        italics: true,
                        spacing: { before: 400 }
                    })
                ]
            }]
        });

        const blob = await docx.Packer.toBlob(doc);
        saveAs(blob, `${extractedData.companyName.replace(/\s+/g, '_')}_Term_Loan_Teaser.docx`);
        
        alert('Document downloaded successfully!');
    } catch (error) {
        console.error('Error generating DOCX:', error);
        alert('Error generating document. Please try again.');
    }
}

// Made with Bob
