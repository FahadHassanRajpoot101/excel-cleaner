// ============================================
// PAYWALL & SUBSCRIPTION SYSTEM
// ============================================
const TRIAL_KEY = 'excel_cleaner_trial_used';
const LICENSE_KEY = 'excel_cleaner_license';

// ⚠️ YAHAN apni Gumroad license keys add karein (jab customer khareede)
const VALID_LICENSES = [
    'DEMO-KEY-1234-5678',  // Test ke liye - baad mein hata dein
];

function isSubscribed() {
    const license = localStorage.getItem(LICENSE_KEY);
    return license && VALID_LICENSES.includes(license);
}

function hasUsedTrial() {
    return localStorage.getItem(TRIAL_KEY) === 'true';
}

function markTrialUsed() {
    localStorage.setItem(TRIAL_KEY, 'true');
}

function showPaywall() {
    document.getElementById('paywallModal')?.classList.add('active');
}

function hidePaywall() {
    document.getElementById('paywallModal')?.classList.remove('active');
}

// Modal event listeners
document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('modalClose')?.addEventListener('click', hidePaywall);
    
    document.getElementById('activateBtn')?.addEventListener('click', function() {
        const input = document.getElementById('licenseInput');
        const error = document.getElementById('licenseError');
        const key = input.value.trim().toUpperCase();
        
        if (!key) {
            error.textContent = 'Please enter a license key';
            return;
        }
        
        if (VALID_LICENSES.includes(key)) {
            localStorage.setItem(LICENSE_KEY, key);
            hidePaywall();
            alert('✅ License activated! Unlimited cleaning unlocked.');
        } else {
            error.textContent = 'Invalid license key. Check your email.';
            input.value = '';
        }
    });
    
    document.getElementById('licenseInput')?.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') document.getElementById('activateBtn')?.click();
    });
});

// ============================================
// GLOBAL VARIABLES
// ============================================
let originalData = [];
let cleanedData = [];

// ============================================
// UI UPDATE FUNCTIONS
// ============================================

function updateStats(originalData, cleanedData, flaggedCount) {
    const rows = document.getElementById('statRows');
    const clean = document.getElementById('statClean');
    const flagged = document.getElementById('statFlagged');
    const dup = document.getElementById('statDuplicates');
    
    if (rows) rows.textContent = cleanedData.length || 0;
    if (clean) {
        const cleanCount = cleanedData.filter(row => {
            return !Object.values(row).some(val => 
                typeof val === 'string' && 
                (val.includes('Invalid') || val.includes('Possible Typo') || val.includes('Duplicate'))
            );
        }).length || 0;
        clean.textContent = cleanCount;
    }
    if (flagged) flagged.textContent = flaggedCount || 0;
    if (dup) dup.textContent = (originalData.length - cleanedData.length) || 0;
}

function setStatus(status) {
    const led = document.getElementById('statusLed');
    if (!led) return;
    led.textContent = status;
    led.className = 'status-led';
    if (status.includes('clean') || status.includes('ready') || status.includes('loaded') || status.includes('downloaded')) {
        led.classList.add('active');
    } else if (status.includes('working') || status.includes('loading')) {
        led.classList.add('working');
    }
}

function updateFileName(name) {
    const el = document.getElementById('fileName');
    if (el) el.textContent = name || 'no file loaded';
}

function updateFooterStatus(text) {
    const el = document.getElementById('footerStatus');
    if (el) el.textContent = text;
}

function updateFooterStats(text) {
    const el = document.getElementById('footerStats');
    if (el) el.textContent = text;
}

// ============================================
// FILE UPLOAD HANDLER
// ============================================
document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    
    if (!file) {
        alert('Please select a file');
        return;
    }
    
    updateFileName(file.name);
    setStatus('⏳ loading');
    updateFooterStatus('⏳ loading file...');
    
    const reader = new FileReader();
    
    reader.onload = function(event) {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            originalData = XLSX.utils.sheet_to_json(firstSheet);
            cleanedData = [...originalData];
            renderTable(cleanedData);
            
            updateStats(originalData, cleanedData, 0);
            setStatus('✓ loaded');
            updateFooterStatus(`✓ ${originalData.length} rows loaded`);
            updateFooterStats(`${file.name} — ${originalData.length} rows`);
            
            console.log('✅ File loaded!', originalData.length, 'rows');
        } catch (error) {
            alert('Error reading file: ' + error.message);
            setStatus('✗ error');
            updateFooterStatus('✗ error loading file');
        }
    };
    
    reader.readAsArrayBuffer(file);
});

// ============================================
// RENDER TABLE FUNCTION
// ============================================
function renderTable(data) {
    const container = document.getElementById('tableContainer');
    
    if (!data || data.length === 0) {
        container.innerHTML = '<p class="table-placeholder">⏎ upload a file to begin</p>';
        return;
    }
    
    const columns = Object.keys(data[0]);
    
    let html = '<table>';
    html += '<thead><tr>';
    columns.forEach(col => {
        html += `<th>${col}</th>`;
    });
    html += '</tr></thead>';
    
    html += '<tbody>';
    data.forEach(row => {
        html += '<tr>';
        columns.forEach(col => {
            const value = row[col] !== undefined && row[col] !== null ? row[col] : '';
            let cellClass = '';
            
            if (typeof value === 'string' && 
                (value.includes('Invalid') || value.includes('Possible Typo') || value.includes('Duplicate'))) {
                cellClass = 'flagged-cell';
            }
            
            html += `<td class="${cellClass}">${value}</td>`;
        });
        html += '</tr>';
    });
    html += '</tbody>';
    html += '</table>';
    
    container.innerHTML = html;
}

// ============================================
// V4.7: COUNTRY NORMALIZATION
// ============================================
function normalizeCountry(country) {
    if (!country || typeof country !== 'string') return country;
    
    const countryMap = {
        'usa': 'United States', 'us': 'United States', 'u.s.': 'United States',
        'u.s.a': 'United States', 'united states': 'United States', 'america': 'United States',
        'american': 'United States', 'uk': 'United Kingdom', 'u.k.': 'United Kingdom',
        'united kingdom': 'United Kingdom', 'great britain': 'United Kingdom',
        'england': 'United Kingdom', 'britain': 'United Kingdom',
        'uae': 'United Arab Emirates', 'u.a.e.': 'United Arab Emirates',
        'united arab emirates': 'United Arab Emirates', 'pak': 'Pakistan',
        'paksitan': 'Pakistan', 'pk': 'Pakistan', 'ind': 'India', 'in': 'India',
        'chn': 'China', 'cn': 'China', 'aus': 'Australia', 'au': 'Australia',
        'can': 'Canada', 'ca': 'Canada', 'germany': 'Germany', 'de': 'Germany',
        'deu': 'Germany', 'france': 'France', 'fr': 'France', 'italy': 'Italy',
        'it': 'Italy', 'spain': 'Spain', 'es': 'Spain', 'japan': 'Japan',
        'jp': 'Japan', 'korea': 'South Korea', 'kr': 'South Korea',
        'south korea': 'South Korea', 'russia': 'Russia', 'ru': 'Russia',
        'brazil': 'Brazil', 'br': 'Brazil', 'mexico': 'Mexico', 'mx': 'Mexico',
        'south africa': 'South Africa', 'za': 'South Africa'
    };
    
    const normalized = country.toLowerCase().trim();
    return countryMap[normalized] || country;
}

// ============================================
// V4.7: DETECT MATH COLUMNS
// ============================================
function detectMathColumns(row) {
    const columns = Object.keys(row);
    const result = { qtyCol: null, priceCol: null, totalCol: null };
    
    const qtyKeywords = ['quantity', 'qty', 'qty.', 'units'];
    for (let col of columns) {
        const lower = col.toLowerCase().trim();
        if (qtyKeywords.some(kw => lower === kw || lower.endsWith(kw))) {
            result.qtyCol = col;
            break;
        }
    }
    
    const priceKeywords = ['unit price', 'unit cost', 'price', 'rate', 'amount'];
    for (let col of columns) {
        const lower = col.toLowerCase().trim();
        if (priceKeywords.some(kw => lower.includes(kw) || lower === kw)) {
            if (col !== result.qtyCol) {
                result.priceCol = col;
                break;
            }
        }
    }
    
    const totalKeywords = ['total', 'net', 'gross', 'bill', 'subtotal', 'grand total'];
    for (let col of columns) {
        const lower = col.toLowerCase().trim();
        if (totalKeywords.some(kw => lower.includes(kw) || lower === kw)) {
            result.totalCol = col;
            break;
        }
    }
    
    return result;
}

// ============================================
// V4.7: AUTO-RECALCULATE TOTAL
// ============================================
function autoRecalculateTotal(row, quantityCol, unitPriceCol, totalCol) {
    let qty = row[quantityCol];
    let price = row[unitPriceCol];
    let existingTotal = row[totalCol];
    
    let qtyNum = null;
    let priceNum = null;
    
    if (qty !== null && qty !== undefined && qty !== '') {
        if (typeof qty === 'number') qtyNum = qty;
        else if (typeof qty === 'string') {
            const clean = cleanNumericValue(qty);
            if (typeof clean === 'number') qtyNum = clean;
        }
    }
    
    if (price !== null && price !== undefined && price !== '') {
        if (typeof price === 'number') priceNum = price;
        else if (typeof price === 'string') {
            const clean = cleanNumericValue(price);
            if (typeof clean === 'number') priceNum = clean;
        }
    }
    
    let totalNum = null;
    if (existingTotal !== null && existingTotal !== undefined && existingTotal !== '') {
        if (typeof existingTotal === 'number') totalNum = existingTotal;
        else if (typeof existingTotal === 'string') {
            const clean = cleanNumericValue(existingTotal);
            if (typeof clean === 'number') totalNum = clean;
        }
    }
    
    let calculatedTotal = null;
    if (qtyNum !== null && priceNum !== null) {
        calculatedTotal = qtyNum * priceNum;
    }
    
    if (calculatedTotal !== null) {
        if (totalNum === null || 
            (typeof totalNum === 'number' && Math.abs(totalNum - calculatedTotal) > 0.001)) {
            return calculatedTotal;
        }
    }
    
    return existingTotal;
}

// ============================================
// V4.7: FLEXIBLE DATE PARSER
// ============================================
function parseDateFlexible(dateStr) {
    if (!dateStr || typeof dateStr !== 'string') return null;
    
    let str = dateStr.trim();
    if (str === '') return null;
    str = str.replace(/\s+/g, ' ');
    
    const monthMap = {
        'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
        'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
        'january': 1, 'february': 2, 'march': 3, 'april': 4, 'june': 6,
        'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12
    };
    
    const textMonthMatch = str.match(/([a-zA-Z]+)\s+(\d{1,2})\s*,?\s*(\d{4})/i) || 
                           str.match(/(\d{1,2})\s+([a-zA-Z]+)\s*,?\s*(\d{4})/i);
    if (textMonthMatch) {
        let monthStr, day, year;
        if (textMonthMatch[1].match(/[a-zA-Z]/)) {
            monthStr = textMonthMatch[1];
            day = parseInt(textMonthMatch[2]);
            year = parseInt(textMonthMatch[3]);
        } else {
            day = parseInt(textMonthMatch[1]);
            monthStr = textMonthMatch[2];
            year = parseInt(textMonthMatch[3]);
        }
        
        const month = monthMap[monthStr.toLowerCase()];
        if (month && day >= 1 && day <= 31 && year >= 1900 && year <= 2100) {
            const date = new Date(year, month - 1, day);
            if (!isNaN(date)) {
                return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
            }
        }
    }
    
    const formats = [
        { regex: /^(\d{4})-(\d{1,2})-(\d{1,2})$/, order: 'ymd' },
        { regex: /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/, order: 'ymd' },
        { regex: /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/, order: 'dmy' },
        { regex: /^(\d{1,2})-(\d{1,2})-(\d{4})$/, order: 'dmy' },
        { regex: /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/, order: 'dmy' },
        { regex: /^(\d{4})\.(\d{1,2})\.(\d{1,2})$/, order: 'ymd' }
    ];
    
    for (let fmt of formats) {
        const match = str.match(fmt.regex);
        if (match) {
            let day, month, year;
            
            if (fmt.order === 'ymd') {
                year = parseInt(match[1]);
                month = parseInt(match[2]);
                day = parseInt(match[3]);
            } else {
                let d = parseInt(match[1]);
                let m = parseInt(match[2]);
                let y = parseInt(match[3]);
                
                if (d > 12 && m <= 12) { day = d; month = m; year = y; }
                else if (m > 12 && d <= 12) { day = m; month = d; year = y; }
                else { day = d; month = m; year = y; }
            }
            
            if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                const date = new Date(year, month - 1, day);
                if (!isNaN(date) && date.getFullYear() === year && date.getMonth() === month - 1) {
                    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                }
            }
        }
    }
    
    const nativeDate = new Date(str);
    if (!isNaN(nativeDate)) {
        const year = nativeDate.getFullYear();
        const month = nativeDate.getMonth() + 1;
        const day = nativeDate.getDate();
        if (year >= 1900 && year <= 2100) {
            return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        }
    }
    
    return null;
}

// ============================================
// V4.7: WORD TO NUMBER CONVERTER
// ============================================
function wordToNumber(word) {
    const map = {
        'zero': 0, 'one': 1, 'two': 2, 'three': 3, 'four': 4,
        'five': 5, 'six': 6, 'seven': 7, 'eight': 8, 'nine': 9,
        'ten': 10, 'eleven': 11, 'twelve': 12, 'thirteen': 13,
        'fourteen': 14, 'fifteen': 15, 'sixteen': 16, 'seventeen': 17,
        'eighteen': 18, 'nineteen': 19, 'twenty': 20, 'thirty': 30,
        'forty': 40, 'fifty': 50, 'sixty': 60, 'seventy': 70,
        'eighty': 80, 'ninety': 90, 'hundred': 100
    };
    return map[word.toLowerCase()] !== undefined ? map[word.toLowerCase()] : null;
}

// ============================================
// V4.7: QUANTITY CLEANING
// ============================================
function cleanQuantityValue(value) {
    if (value === null || value === undefined) return null;
    
    try {
        let num;
        if (typeof value === 'number') {
            num = value;
        } else if (typeof value === 'string') {
            const wordNum = wordToNumber(value.trim());
            if (wordNum !== null) {
                num = wordNum;
            } else {
                let str = value.trim();
                str = str.replace(/[$€£PKR₹¥₨,]/g, '');
                str = str.replace(/[Oo]/g, '0');
                str = str.replace(/[^0-9.-]/g, '');
                num = parseFloat(str);
            }
        } else {
            return 'Invalid Value';
        }
        
        if (isNaN(num)) return 'Invalid Value';
        
        if (num < 0) return 'Invalid (Negative)';
        else if (num === 0) return 0;
        else return Math.floor(num);
        
    } catch (error) {
        return 'Invalid Value';
    }
}

// ============================================
// V4.7: CLEAN NUMERIC VALUE
// ============================================
function cleanNumericValue(value) {
    if (value === null || value === undefined) return null;
    
    if (typeof value === 'number') {
        if (value === 0) return 0;
        if (value < 0) return 'Invalid (Negative)';
        return value;
    }
    
    if (typeof value === 'string') {
        let str = value.trim();
        
        const nullPlaceholders = ['n/a', 'null', 'none', 'nil', 'na', '-', '--', 'nan', 'undefined'];
        if (nullPlaceholders.includes(str.toLowerCase())) return null;
        
        const wordNum = wordToNumber(str);
        if (wordNum !== null) return wordNum;
        
        let cleaned = str;
        cleaned = cleaned.replace(/[$€£PKR₹¥₨,]/g, '');
        cleaned = cleaned.replace(/[Oo]/g, '0');
        cleaned = cleaned.replace(/\.{2,}/g, '.');
        cleaned = cleaned.replace(/[^0-9.-]/g, '');
        cleaned = cleaned.trim();
        
        if (cleaned === '' || cleaned === '.' || cleaned === '-') {
            return 'Invalid Value';
        }
        
        const num = parseFloat(cleaned);
        if (!isNaN(num)) {
            if (num === 0) return 0;
            if (num < 0) return 'Invalid (Negative)';
            return num;
        }
        
        return 'Invalid Value';
    }
    
    return 'Invalid Value';
}

// ============================================
// V4.7: EMAIL DOMAIN TYPO DETECTION
// ============================================
const typoDomains = {
    'gamil.com': 'gmail.com', 'gmial.com': 'gmail.com', 'gmail.con': 'gmail.com',
    'gmail.cm': 'gmail.com', 'gmai.com': 'gmail.com', 'gmal.com': 'gmail.com',
    'yahooo.com': 'yahoo.com', 'yaho.com': 'yahoo.com', 'yahho.com': 'yahoo.com',
    'yahoo.con': 'yahoo.com', 'yhoo.com': 'yahoo.com', 'hotmial.com': 'hotmail.com',
    'hotmail.con': 'hotmail.com', 'hotmal.com': 'hotmail.com', 'hotmil.com': 'hotmail.com',
    'homtail.com': 'hotmail.com', 'outlok.com': 'outlook.com', 'outllok.com': 'outlook.com',
    'outlook.con': 'outlook.com', 'gmial.con': 'gmail.com', 'gamil.con': 'gmail.com',
    'yahoo.cm': 'yahoo.com', 'hotmail.cm': 'hotmail.com', 'gmil.com': 'gmail.com',
    'yhaoo.com': 'yahoo.com', 'yahooo.cm': 'yahoo.com'
};

function validateEmail(email) {
    if (!email) return { valid: false, reason: 'Empty' };
    
    const emailStr = email.toString().trim();
    if (emailStr === '') return { valid: false, reason: 'Empty' };
    
    const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
    if (!emailRegex.test(emailStr)) {
        return { valid: false, reason: 'Invalid Format' };
    }
    
    const invalidChars = /[#!$%^&*()=+{}[\]|\\;:'",<>/?]/;
    if (invalidChars.test(emailStr)) {
        return { valid: false, reason: 'Invalid Characters' };
    }
    
    const domain = emailStr.split('@')[1].toLowerCase();
    if (typoDomains[domain]) {
        return { 
            valid: false, 
            reason: `Possible Typo: Did you mean @${typoDomains[domain]}?`
        };
    }
    
    return { valid: true, reason: 'Valid' };
}

// ============================================
// V4.7: DUPLICATE ID DETECTION
// ============================================
function detectAndFlagDuplicateIds(cleaned) {
    if (!cleaned || cleaned.length === 0) return cleaned;
    
    const columns = Object.keys(cleaned[0]);
    let idColumns = [];
    
    for (let col of columns) {
        const lower = col.toLowerCase();
        if (lower.includes('id') || lower.includes('order') || 
            lower.includes('invoice') || lower.includes('ticket') ||
            lower === 's.no' || lower === 'sr.no' || lower === '#' ||
            lower === 'row id' || lower === 'record id' || lower.includes('serial')) {
            idColumns.push(col);
        }
    }
    
    if (idColumns.length === 0) return cleaned;
    
    const idColumn = idColumns[0];
    const idGroups = new Map();
    
    for (let row of cleaned) {
        const id = row[idColumn];
        if (id !== null && id !== undefined && id !== '') {
            const key = id.toString().trim();
            if (!idGroups.has(key)) idGroups.set(key, []);
            idGroups.get(key).push(row);
        }
    }
    
    const result = [];
    const flaggedRows = [];
    
    for (let [id, rows] of idGroups) {
        if (rows.length === 1) {
            result.push(rows[0]);
        } else {
            const idIndex = columns.indexOf(idColumn);
            const compareColumns = columns.filter((_, idx) => idx !== idIndex);
            
            const signatures = rows.map(row => {
                return compareColumns.map(col => {
                    const val = row[col];
                    return val !== undefined && val !== null ? val.toString().trim() : '';
                }).join('|||');
            });
            
            const firstSignature = signatures[0];
            const allIdentical = signatures.every(sig => sig === firstSignature);
            
            if (allIdentical) {
                result.push(rows[0]);
            } else {
                for (let row of rows) {
                    const flaggedRow = { ...row };
                    flaggedRow._duplicate_status = 'Duplicate ID — Conflicting Data';
                    flaggedRows.push(flaggedRow);
                }
            }
        }
    }
    
    return [...result, ...flaggedRows];
}

// ============================================
// V4.7: GENERATE CLEANING SUMMARY
// ============================================
function generateCleaningSummary(originalData, cleanedData) {
    const summary = {
        originalRows: originalData.length,
        cleanedRows: cleanedData.length,
        removedRows: originalData.length - cleanedData.length,
        invalidEmail: 0, invalidPhone: 0, invalidQuantity: 0, invalidPrice: 0,
        invalidAge: 0, invalidDate: 0, invalidTotal: 0, negativeValues: 0,
        duplicatesFlagged: 0, duplicateIds: 0, invalidValue: 0
    };
    
    cleanedData.forEach(row => {
        Object.entries(row).forEach(([key, val]) => {
            if (typeof val !== 'string') return;
            const lowerKey = key.toLowerCase();
            
            if (val.includes('Invalid Email') || val.includes('Possible Typo')) {
                summary.invalidEmail++;
            } else if (val.includes('Invalid Phone')) {
                summary.invalidPhone++;
            } else if (val.includes('Invalid Age')) {
                summary.invalidAge++;
            } else if (val.includes('Invalid Date')) {
                summary.invalidDate++;
            } else if (val.includes('Invalid (Negative)')) {
                if (lowerKey.includes('quantity') || lowerKey.includes('qty') || lowerKey.includes('unit')) summary.invalidQuantity++;
                else if (lowerKey.includes('price') || lowerKey.includes('amount') || lowerKey.includes('cost') || lowerKey.includes('rate')) summary.invalidPrice++;
                else if (lowerKey.includes('total') || lowerKey.includes('net') || lowerKey.includes('gross') || lowerKey.includes('bill')) summary.invalidTotal++;
                else summary.negativeValues++;
            } else if (val.includes('Invalid Value')) {
                summary.invalidValue++;
                if (lowerKey.includes('quantity') || lowerKey.includes('qty') || lowerKey.includes('unit')) summary.invalidQuantity++;
                else if (lowerKey.includes('price') || lowerKey.includes('amount') || lowerKey.includes('cost') || lowerKey.includes('rate')) summary.invalidPrice++;
                else if (lowerKey.includes('total') || lowerKey.includes('net') || lowerKey.includes('gross') || lowerKey.includes('bill')) summary.invalidTotal++;
            } else if (val.includes('Duplicate ID')) {
                summary.duplicateIds++;
            }
        });
        
        if (row._duplicate_status && row._duplicate_status.includes('Duplicate')) {
            summary.duplicatesFlagged++;
        }
    });
    
    return summary;
}

// ============================================
// V4.7: MAIN CLEANING FUNCTION
// ============================================
function cleanExcelData(data) {
    let cleaned = [...data];
    
    // RULE 1: Text Normalization
    cleaned = cleaned.map(row => {
        const newRow = {};
        for (let key in row) {
            let value = row[key];
            
            if (typeof value === 'string') {
                value = value.trim();
                value = value.replace(/\s+/g, ' ');
                
                if (key.toLowerCase().includes('country') || key.toLowerCase().includes('nation')) {
                    value = normalizeCountry(value);
                }
                
                const categoricalColumns = ['Status', 'Gender', 'Category', 'Type', 'Region', 'Department', 'Grade'];
                if (categoricalColumns.includes(key)) {
                    value = value.toUpperCase();
                } else {
                    const skipColumns = ['Email', 'Phone', 'ID', 'Code', 'URL', 'Website'];
                    if (!skipColumns.includes(key)) {
                        value = value.toLowerCase().split(' ').map(word => 
                            word.charAt(0).toUpperCase() + word.slice(1)
                        ).join(' ');
                    }
                }
            }
            
            newRow[key] = value;
        }
        return newRow;
    });
    
    // RULE 2: Date Parsing
    cleaned = cleaned.map(row => {
        const newRow = { ...row };
        
        for (let key in row) {
            const lowerKey = key.toLowerCase();
            const dateKeywords = ['date', 'joining', 'dob', 'birth', 'created', 'updated', 'start', 'end', 'hire', 'issued', 'expiry'];
            
            if (dateKeywords.some(kw => lowerKey.includes(kw))) {
                const value = row[key];
                if (typeof value === 'string' && value.trim() !== '') {
                    const parsed = parseDateFlexible(value);
                    newRow[key] = parsed || 'Invalid Date';
                } else if (typeof value === 'number') {
                    try {
                        const date = new Date((value - 25569) * 86400 * 1000);
                        if (!isNaN(date)) {
                            const year = date.getFullYear();
                            const month = String(date.getMonth() + 1).padStart(2, '0');
                            const day = String(date.getDate()).padStart(2, '0');
                            newRow[key] = `${year}-${month}-${day}`;
                        }
                    } catch (e) {
                        newRow[key] = 'Invalid Date';
                    }
                }
            }
        }
        return newRow;
    });
    
    // RULE 3: Numeric Cleaning
    cleaned = cleaned.map(row => {
        const newRow = { ...row };
        
        for (let key in row) {
            const lowerKey = key.toLowerCase();
            const quantityKeywords = ['quantity', 'qty', 'qty.', 'units'];
            const isQuantity = quantityKeywords.some(kw => lowerKey === kw || lowerKey.endsWith(kw));
            
            if (isQuantity) {
                const cleanedQty = cleanQuantityValue(row[key]);
                newRow[key] = cleanedQty !== null ? cleanedQty : '';
            } else {
                const numericKeywords = ['salary', 'price', 'amount', 'budget', 'income', 'cost', 'fee', 'payment', 'total', 'balance', 'age', 'score', 'rate', 'tax'];
                if (numericKeywords.some(kw => lowerKey.includes(kw))) {
                    const cleanedNum = cleanNumericValue(row[key]);
                    if (cleanedNum !== null) {
                        if (lowerKey === 'age' && typeof cleanedNum === 'number') {
                            if (cleanedNum < 0 || cleanedNum > 120) {
                                newRow[key] = 'Invalid Age';
                                continue;
                            }
                        }
                        newRow[key] = cleanedNum;
                    } else {
                        newRow[key] = '';
                    }
                }
            }
        }
        return newRow;
    });
    
    // RULE 4: Email & Phone Validation
    cleaned = cleaned.map(row => {
        const newRow = { ...row };
        
        if (row.Email) {
            const result = validateEmail(row.Email);
            if (!result.valid) newRow.Email = result.reason;
        }
        
        if (row.Phone) {
            let phone = row.Phone.toString().replace(/[^0-9+]/g, '');
            phone = phone.replace(/\+{2,}/g, '+');
            
            if (phone.length >= 10) {
                if (phone.startsWith('92') && phone.length === 12) newRow.Phone = `+${phone}`;
                else if (phone.startsWith('0') && phone.length === 11) newRow.Phone = `+92${phone.substring(1)}`;
                else if (phone.length === 10) newRow.Phone = `+92${phone}`;
                else newRow.Phone = phone;
            } else {
                newRow.Phone = 'Invalid Phone';
            }
        }
        
        return newRow;
    });
    
    // AUTO-RECALCULATE TOTAL
    if (cleaned.length > 0) {
        const mathCols = detectMathColumns(cleaned[0]);
        
        if (mathCols.qtyCol && mathCols.priceCol && mathCols.totalCol) {
            let correctedCount = 0;
            
            cleaned = cleaned.map(row => {
                const newRow = { ...row };
                const correctedTotal = autoRecalculateTotal(row, mathCols.qtyCol, mathCols.priceCol, mathCols.totalCol);
                const oldTotal = row[mathCols.totalCol];
                if (oldTotal !== correctedTotal && correctedTotal !== null) correctedCount++;
                newRow[mathCols.totalCol] = correctedTotal;
                return newRow;
            });
            
            console.log(`✅ Auto-corrected ${correctedCount} rows with incorrect totals`);
        }
    }
    
    cleaned = detectAndFlagDuplicateIds(cleaned);
    
    // RULE 5: Null Handling & Deduplication
    const nullPlaceholders = ['N/A', 'null', 'NONE', '-', 'nan', 'undefined', 'NaN', 'None', 'Nil', 'NULL', 'NA', 'n/a'];
    cleaned = cleaned.map(row => {
        const newRow = {};
        for (let key in row) {
            let value = row[key];
            if (typeof value === 'string') {
                const trimmed = value.trim();
                newRow[key] = (nullPlaceholders.includes(trimmed) || trimmed === '') ? '' : value;
            } else {
                newRow[key] = value;
            }
        }
        return newRow;
    });
    
    cleaned = cleaned.filter(row => {
        return Object.values(row).some(val => 
            val !== '' && val !== null && val !== undefined && 
            val !== 'Invalid Date' && val !== 'Invalid Email' && val !== 'Invalid Phone' &&
            val !== 'Invalid Age' && val !== 'Invalid (Negative)' && val !== 'Invalid Value' &&
            val !== 'Invalid Format' && val !== 'Invalid Characters' &&
            !String(val).includes('Possible Typo')
        );
    });
    
    const seen = new Set();
    cleaned = cleaned.filter(row => {
        const { _duplicate_status, ...rowWithoutStatus } = row;
        const key = JSON.stringify(rowWithoutStatus).toLowerCase();
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
    });
    
    return cleaned;
}

// ============================================
// CLEAN DATA BUTTON (WITH PAYWALL)
// ============================================
document.getElementById('cleanBtn').addEventListener('click', function() {
    if (originalData.length === 0) {
        alert('Please upload an Excel file first!');
        return;
    }
    
    // === PAYWALL CHECK ===
    if (hasUsedTrial() && !isSubscribed()) {
        showPaywall();
        return;
    }
    if (!isSubscribed()) {
        markTrialUsed();  // Pehli baar free
    }
    // === END PAYWALL CHECK ===
    
    setStatus('⚙️ working');
    updateFooterStatus('⚙️ cleaning...');
    
    cleanedData = cleanExcelData(originalData);
    renderTable(cleanedData);
    
    let flaggedCount = 0;
    cleanedData.forEach(row => {
        Object.values(row).forEach(val => {
            if (typeof val === 'string' && 
                (val.includes('Invalid') || val.includes('Possible Typo') || val.includes('Duplicate'))) {
                flaggedCount++;
            }
        });
    });
    
    updateStats(originalData, cleanedData, flaggedCount);
    setStatus('✓ cleaned');
    updateFooterStatus(`✓ ${cleanedData.length} rows, ${flaggedCount} flagged`);
    updateFooterStats(`${cleanedData.length} rows · ${flaggedCount} flagged`);
    
    const summary = generateCleaningSummary(originalData, cleanedData);
    let message = `✅ Data cleaned successfully!\n\n`;
    message += `📊 Overall:\n`;
    message += `   Original Rows: ${summary.originalRows}\n`;
    message += `   Cleaned Rows: ${summary.cleanedRows}\n`;
    message += `   Removed: ${summary.removedRows} rows\n\n`;
    message += `📋 Invalid Data Breakdown:\n`;
    message += `   Invalid Emails: ${summary.invalidEmail}\n`;
    message += `   Invalid Phones: ${summary.invalidPhone}\n`;
    message += `   Invalid Quantities: ${summary.invalidQuantity}\n`;
    message += `   Invalid Prices: ${summary.invalidPrice}\n`;
    message += `   Invalid Totals: ${summary.invalidTotal}\n`;
    message += `   Invalid Ages: ${summary.invalidAge}\n`;
    message += `   Invalid Dates: ${summary.invalidDate}\n`;
    message += `   Negative Values (Other): ${summary.negativeValues}\n`;
    message += `   Invalid Values (Garbage Text): ${summary.invalidValue}\n`;
    message += `   Duplicate IDs Flagged: ${summary.duplicateIds}\n`;
    alert(message);
});

// ============================================
// DOWNLOAD BUTTON
// ============================================
document.getElementById('downloadBtn').addEventListener('click', function() {
    if (cleanedData.length === 0) {
        alert('No data to download! Please upload and clean data first.');
        return;
    }
    
    setStatus('⬇ downloading');
    updateFooterStatus('⬇ preparing download...');
    
    try {
        let invalidCount = 0;
        cleanedData.forEach(row => {
            Object.values(row).forEach(val => {
                if (typeof val === 'string' && (val.includes('Invalid') || val.includes('Possible Typo') || val.includes('Duplicate'))) {
                    invalidCount++;
                }
            });
        });
        
        if (invalidCount > 0) {
            if (!confirm(`⚠️ Warning: ${invalidCount} cells have invalid data (marked as "Invalid...", "Possible Typo", or "Duplicate").\n\nDo you still want to download?`)) {
                setStatus('✓ ready');
                updateFooterStatus('✓ download cancelled');
                return;
            }
        }
        
        const ws = XLSX.utils.json_to_sheet(cleanedData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Cleaned Data");
        XLSX.writeFile(wb, "cleaned_data.xlsx");
        
        setStatus('✓ downloaded');
        updateFooterStatus('✓ file downloaded');
        alert('✅ File downloaded successfully!');
    } catch (error) {
        alert('Error downloading file: ' + error.message);
        setStatus('✗ error');
        updateFooterStatus('✗ download error');
    }
});

console.log('✅ Excel Data Cleaner V4.7 with Paywall Loaded!');
