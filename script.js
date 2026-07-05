const compareForm = document.getElementById('compareForm');
const compareStatus = document.getElementById('compareStatus');
const analyticsCard = document.getElementById('analyticsCard');
const summaryGrid = document.getElementById('summaryGrid');
const issueTableBody = document.getElementById('issueTableBody');
const clusterTableBody = document.getElementById('clusterTableBody');
const heatmapGrid = document.getElementById('heatmapGrid');
const geoLocationBody = document.getElementById('geoLocationBody');
const geoPinBody = document.getElementById('geoPinBody');
const orphanGrid = document.getElementById('orphanGrid');
const orphanTableBody = document.getElementById('orphanTableBody');
const clusterSubtext = document.getElementById('clusterSubtext');
const dedupeGrid = document.getElementById('dedupeGrid');
const downloadReportButton = document.getElementById('downloadReport');
const accountUpload = document.getElementById('accountUpload');
const contactUpload = document.getElementById('contactUpload');
const combinedUpload = document.getElementById('combinedUpload');

const progressOverlay = document.getElementById('progressOverlay');
const progressTitle = document.getElementById('progressTitle');
const progressFill = document.getElementById('progressFill');
const progressStepList = document.getElementById('progressStepList');
const progressFootnote = document.getElementById('progressFootnote');

const detailOverlay = document.getElementById('detailOverlay');
const detailEyebrow = document.getElementById('detailEyebrow');
const detailTitle = document.getElementById('detailTitle');
const detailBody = document.getElementById('detailBody');
const detailClose = document.getElementById('detailClose');

let lastRunData = null;
let lastExportWorkbook = null;
let lastExportFilename = 'account_contact_quality_report.xlsx';

// Pagination state
let clusterPaginationState = {
  allClusters: [],
  displayed: 0,
  pageSize: 100
};

const LOCATION_NAMES = [
  'Maharashtra', 'Delhi', 'Karnataka', 'Tamil Nadu', 'Uttar Pradesh', 'Gujarat', 'Rajasthan',
  'West Bengal', 'Bihar', 'Telangana', 'Andhra Pradesh', 'Kerala', 'Punjab', 'Haryana',
  'Chennai', 'Mumbai', 'Bengaluru', 'Hyderabad', 'Pune', 'Kolkata', 'Jaipur', 'Lucknow',
  'Ahmedabad', 'Surat', 'Nagpur', 'Noida', 'Gurgaon', 'Faridabad', 'Kanpur', 'Agra',
  'Chandigarh', 'Vadodara', 'Coimbatore', 'Mysore', 'Pondicherry', 'Visakhapatnam', 'Indore'
];

const PROGRESS_STEPS = [
  { title: 'Uploading your file…', footnote: 'Reading bytes from your selected workbook.', fill: 10 },
  { title: 'Reading the workbook…', footnote: 'Parsing sheets, headers, and rows.', fill: 28 },
  { title: 'Cleaning & validating…', footnote: 'Applying cleanup rules and standardization.', fill: 52 },
  { title: 'Building clusters & insights…', footnote: 'Grouping likely duplicates and scoring data quality.', fill: 78 },
  { title: 'Building your report…', footnote: 'Assembling summary, analytics, and cleaned data sheets.', fill: 94 }
];

function setStatus(message) {
  compareStatus.textContent = message;
}

function toStringSafe(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return '';
  return String(value);
}

function normalizeText(value) {
  return toStringSafe(value).trim().replace(/\s+/g, ' ');
}

function normalizeHeaderValue(value) {
  return normalizeText(value).toLowerCase();
}

function properCase(value) {
  return normalizeText(value).toLowerCase().replace(/\b[a-z]/g, (match) => match.toUpperCase());
}

function stripSpecial(value) {
  return normalizeText(value).replace(/[^a-zA-Z0-9\s&/\-\.]/g, ' ').replace(/\s+/g, ' ').trim();
}

function removeTextInBrackets(value) {
  return normalizeText(value).replace(/[\(\[\{][^\)\]\}]*[\)\]\}]/g, ' ').replace(/\s+/g, ' ').trim();
}

function removeInvalidAccountPrefix(value) {
  return normalizeText(value).replace(/^(m\/s\.?|m\.s\.?|messrs\.?\s*)/i, '').trim();
}

function replaceEntToEnterprise(value) {
  return normalizeText(value).replace(/\b(ent)\b/gi, 'Enterprise');
}

function removePrivateDotLimited(value) {
  return normalizeText(value).replace(/private\.\s*limited/gi, 'Private Limited');
}

function stripTerminalDotsAndCommas(value) {
  return normalizeText(value).replace(/^[\.,\s]+|[\.,\s]+$/g, '').trim();
}

function cleanAccountNameLocation(value) {
  const text = normalizeText(value);
  const lower = text.toLowerCase();
  for (const location of LOCATION_NAMES) {
    const escaped = location.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&');
    const pattern = new RegExp(`[,-]\\s*${escaped}(?:\\b|$)`, 'i');
    const match = lower.match(pattern);
    if (match) {
      return stripSpecial(text.slice(0, match.index).trim());
    }
  }
  return text;
}

function cleanPhone(value) {
  const digits = toStringSafe(value).replace(/\D/g, '');
  if (digits.length === 11 && digits.startsWith('0')) return digits;
  if (digits.length === 10) return `0${digits}`;
  if (digits.length === 12 && digits.startsWith('91')) return `0${digits.slice(2)}`;
  return digits;
}

function cleanMobile(value) {
  const digits = toStringSafe(value).replace(/\D/g, '');
  if (digits.length === 10) return digits;
  if (digits.length === 11 && digits.startsWith('0')) return digits.slice(1);
  if (digits.length === 12 && digits.startsWith('91')) return digits.slice(2);
  return digits;
}

function cleanPinCode(value) {
  return toStringSafe(value).replace(/\D/g, '');
}

function cleanEmail(value) {
  return normalizeText(value).toLowerCase();
}

function readFileAsWorkbook(file) {
  return new Promise((resolve, reject) => {
    const ext = file.name.slice(file.name.lastIndexOf('.')).toLowerCase();
    const reader = new FileReader();
    reader.onerror = () => reject(new Error('Unable to read file.'));
    reader.onload = (event) => {
      try {
        const data = event.target.result;
        const workbook = ext === '.csv'
          ? XLSX.read(data, { type: 'string', raw: false })
          : XLSX.read(data, { type: 'array', raw: false });
        resolve(workbook);
      } catch (error) {
        reject(error);
      }
    };
    if (ext === '.csv') {
      reader.readAsText(file, 'utf8');
    } else {
      reader.readAsArrayBuffer(file);
    }
  });
}

function workbookToJsonArray(workbook) {
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, { defval: null, raw: false });
}

function findHeaderKey(sampleRow, candidates) {
  const headers = Object.keys(sampleRow).reduce((acc, key) => {
    acc[normalizeHeaderValue(key)] = key;
    return acc;
  }, {});
  const found = candidates.find((candidate) => headers[normalizeHeaderValue(candidate)]);
  return found ? headers[normalizeHeaderValue(found)] : null;
}

function buildActualColumns(records, columnMapping) {
  const actual = {};
  if (!records.length) return actual;
  const sample = records[0];
  const headers = Object.keys(sample).reduce((acc, key) => {
    acc[normalizeHeaderValue(key)] = key;
    return acc;
  }, {});

  Object.entries(columnMapping).forEach(([logical, candidates]) => {
    const found = candidates.find((candidate) => headers[normalizeHeaderValue(candidate)]);
    if (found) actual[logical] = headers[normalizeHeaderValue(found)];
  });
  return actual;
}

function ensureContactNameColumn(records, columnMapping) {
  if (!records.length) return;
  const sample = records[0];
  const contactNameCandidates = columnMapping['Contact Name'] || ['Contact Name'];
  if (findHeaderKey(sample, contactNameCandidates)) return;

  const firstNameCandidates = ['First Name', 'FirstName', 'Fname', 'First', 'Given Name'];
  const lastNameCandidates = ['Last Name', 'LastName', 'Lname', 'Last', 'Surname', 'Family Name'];
  const firstKey = findHeaderKey(sample, firstNameCandidates);
  const lastKey = findHeaderKey(sample, lastNameCandidates);
  if (!firstKey && !lastKey) return;

  records.forEach((row) => {
    const first = firstKey ? toStringSafe(row[firstKey]) : '';
    const last = lastKey ? toStringSafe(row[lastKey]) : '';
    row['Contact Name'] = normalizeText(`${first} ${last}`);
  });
}

function ensureCorrectedColumns(records, actualCols) {
  records.forEach((row) => {
    Object.keys(actualCols).forEach((logical) => {
      const correctedKey = `${logical}_corrected`;
      if (!(correctedKey in row)) {
        row[correctedKey] = toStringSafe(row[actualCols[logical]]);
      }
    });
  });
}

function appendIssue(existing, issueName) {
  const current = normalizeText(existing);
  if (!current) return issueName;
  const parts = current.split(',').map((part) => part.trim()).filter(Boolean);
  if (parts.includes(issueName)) return current;
  return `${current}, ${issueName}`;
}

function autoSizeColumns(worksheet, minWidth = 10, maxWidth = 48) {
  worksheet.columns.forEach((column) => {
    let maxLength = minWidth;
    column.eachCell({ includeEmpty: true }, (cell) => {
      const text = cell.value == null ? '' : String(cell.value).replace(/\r\n/g, ' ');
      maxLength = Math.max(maxLength, Math.min(text.length + 2, maxWidth));
    });
    column.width = maxLength;
  });
}

function styleHeaderRow(worksheet) {
  worksheet.getRow(1).eachCell((cell) => {
    cell.font = { bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEAFE' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
  });
}

function saveWorkbookToFile(workbook, filename) {
  return workbook.xlsx.writeBuffer().then((buffer) => {
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
  });
}

function buildRuleFlags() {
  return {
    'Account Name - Remove Brackets': 'Y',
    'Account Name - Contains Location': 'Y',
    'Account Name - Invalid Prefixes': 'Y',
    'Account Name - Remove Punctuation (. and ,)': 'Y',
    'Account Name - Ent/ent to Enterprise': 'Y',
    'Account Name - Remove Dot: Private. Limited': 'Y',
    'Account Name - Short form correction': 'Y',
    'Account Name - Proper case': 'Y',
    'Address - Clean Leading/Trailing Punctuation': 'Y',
    'Address - Apartment': 'Y',
    'Address - Avenue': 'Y',
    'Address - Building': 'Y',
    'Address - District': 'Y',
    'Address - Extension': 'Y',
    'Address - Floor': 'Y',
    'Address - Industrial': 'Y',
    'Address - Lane': 'Y',
    'Address - Market': 'Y',
    'Address - Near': 'Y',
    'Address - Opposite': 'Y',
    'Address - Post Office': 'Y',
    'Address - Road': 'Y',
    'Address - Street': 'Y',
    'Address - Town': 'Y',
    'Address - Village': 'Y',
    'Address - House Number': 'Y',
    'Address - Colony': 'Y',
    'Address - Proper case': 'Y',
    'Location - Keep Only Last Value After Final Comma': 'Y',
    'Location - Remove Brackets & Content': 'Y',
    'Location - Remove Digits': 'Y',
    'Location - Clean Leading/Trailing Dots & Commas': 'Y',
    'Location - Proper case': 'Y',
    'Pin Code - Missing': 'Y',
    'Job Title - Missing': 'Y',
    'Department - Missing': 'Y',
    'Mobile Number - Missing': 'Y',
    'Phone Number - Missing': 'Y',
    'Fax Number - Missing': 'Y',
    'Pin Code - 6 digits': 'Y',
    'Phone Number - 11 digits': 'Y',
    'Mobile Number - 10 digits': 'Y',
    'Fax Number - Format/Length': 'Y',
    'Email - lowercase': 'Y',
    'Job Title - Proper case': 'Y',
    'Department - Proper case': 'Y',
    'Contact Name - Proper case': 'Y'
  };
}

function buildColumnMapping() {
  return {
    'Account Name': ['Account Name', 'AccountName', 'Company Name', 'Legal Name'],
    'Address': ['Address', 'Street Address', 'Address Line 1', 'Address Line'],
    'Location': ['Location', 'City', 'Town', 'Region', 'State'],
    'Pin Code': ['Pin Code', 'Pincode', 'Postal Code', 'Zip', 'Zip Code'],
    'Account Record ID': ['Account Record ID', 'Account ID', 'Account Id', 'Account RecoRoadID'],
    'Contact Name': ['Contact Name', 'Name', 'Contact Full Name'],
    'Phone Number': ['Phone', 'Phone Number', 'Telephone', 'Landline'],
    'Mobile Number': ['Mobile Number', 'Mobile', 'Mobile No', 'Mobile No.'],
    'Fax Number': ['Fax Number', 'Fax', 'Fax No', 'Fax No.'],
    'Email ID': ['Email ID', 'Email', 'Email Address', 'Email Addresss'],
    'Job Title': ['Job Title', 'Designation', 'Role'],
    'Department': ['Department', 'Dept'],
    'Contact Record ID': ['Contact Record ID', 'Contact ID', 'ContactID']
  };
}
function cleanData(records, ruleFlags, columnMapping, mode) {
  ensureContactNameColumn(records, columnMapping);
  const df = records.map((record) => ({ ...record, Issues: '' }));
  const issueSummary = {};
  const actualCols = buildActualColumns(df, columnMapping);
  ensureCorrectedColumns(df, actualCols);

  const ruleGroups = [
    {
      name: 'Account Name - Remove Brackets', logical: 'Account Name', modes: ['account', 'combined'], action: removeTextInBrackets
    },
    {
      name: 'Account Name - Contains Location', logical: 'Account Name', modes: ['account', 'combined'], action: cleanAccountNameLocation
    },
    {
      name: 'Account Name - Invalid Prefixes', logical: 'Account Name', modes: ['account', 'combined'], action: removeInvalidAccountPrefix
    },
    {
      name: 'Account Name - Remove Punctuation (. and ,)', logical: 'Account Name', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/[.,]/g, ' '))
    },
    {
      name: 'Account Name - Ent/ent to Enterprise', logical: 'Account Name', modes: ['account', 'combined'], action: replaceEntToEnterprise
    },
    {
      name: 'Account Name - Remove Dot: Private. Limited', logical: 'Account Name', modes: ['account', 'combined'], action: removePrivateDotLimited
    },
    {
      name: 'Account Name - Short form correction', logical: 'Account Name', modes: ['account', 'combined'], action: (value) => {
        let next = normalizeText(value);
        const patterns = [
          { regex: /\bpvt\.?\s*ltd\.?\b/gi, replace: 'Private Limited' },
          { regex: /\bpvt\.?\b/gi, replace: 'Private' },
          { regex: /\bltd\.?\b/gi, replace: 'Limited' },
          { regex: /\bllp\b/gi, replace: 'LLP' },
          { regex: /\bbros\.?\b/gi, replace: 'Brothers' },
          { regex: /\b(?:&|and)\s+sons\b/gi, replace: 'and Sons' },
          { regex: /\b(?:&|and)\s+co\.?\b/gi, replace: 'Company' },
          { regex: /\bcorp\.?\b/gi, replace: 'Corporation' },
          { regex: /\binc\.?\b/gi, replace: 'Incorporated' },
          { regex: /\bmfg\.?\b/gi, replace: 'Manufacturing' },
          { regex: /\btrdrs\.?\b/gi, replace: 'Traders' },
          { regex: /\bassoc\.?\b/gi, replace: 'Associates' },
          { regex: /\bcons\.?\b/gi, replace: 'Consultants' },
          { regex: /\bsvc\.?\b/gi, replace: 'Services' },
          { regex: /\belec\.?\b/gi, replace: 'Electrical' },
          { regex: /\bengg\.?\b/gi, replace: 'Engineering' }
        ];
        patterns.forEach(({ regex, replace }) => {
          next = next.replace(regex, replace);
        });
        return normalizeText(next);
      }
    },
    {
      name: 'Account Name - Proper case', logical: 'Account Name', modes: ['account', 'combined'], action: (value) => properCase(stripSpecial(value))
    },
    {
      name: 'Contact Name - Proper case', logical: 'Contact Name', modes: ['contact', 'combined'], action: (value) => properCase(stripSpecial(value))
    },
    {
      name: 'Job Title - Proper case', logical: 'Job Title', modes: ['contact', 'combined'], action: (value) => properCase(stripSpecial(value))
    },
    {
      name: 'Department - Proper case', logical: 'Department', modes: ['contact', 'combined'], action: (value) => properCase(stripSpecial(value))
    },
    {
      name: 'Address - Clean Leading/Trailing Punctuation', logical: 'Address', modes: ['account', 'combined'], action: stripTerminalDotsAndCommas
    },
    {
      name: 'Address - Apartment', logical: 'Address', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/\bapt\.?\b/gi, 'Apartment'))
    },
    {
      name: 'Address - Avenue', logical: 'Address', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/\b(av[e]?|avn?)\.?\b/gi, 'Avenue'))
    },
    {
      name: 'Address - Building', logical: 'Address', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/\b(bldg|bld)\.?\b/gi, 'Building'))
    },
    {
      name: 'Address - District', logical: 'Address', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/\b(dist|dis)\.?\b/gi, 'District'))
    },
    {
      name: 'Address - Extension', logical: 'Address', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/\bext\.?\b/gi, 'Extension'))
    },
    {
      name: 'Address - Floor', logical: 'Address', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/\b(fl|flr)\.?\b/gi, 'Floor'))
    },
    {
      name: 'Address - Industrial', logical: 'Address', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/\b(ind|indust)\.?\b/gi, 'Industrial'))
    },
    {
      name: 'Address - Lane', logical: 'Address', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/\b(lane|ln)\.?\b/gi, 'Lane'))
    },
    {
      name: 'Address - Market', logical: 'Address', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/\b(mkt|mrkt)\.?\b/gi, 'Market'))
    },
    {
      name: 'Address - Near', logical: 'Address', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/\b(nr|near)\.?\b/gi, 'Near'))
    },
    {
      name: 'Address - Opposite', logical: 'Address', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/\b(opp|oppo)\.?\b/gi, 'Opposite'))
    },
    {
      name: 'Address - Post Office', logical: 'Address', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/\b(p\.?o\.?|post office)\b/gi, 'Post Office'))
    },
    {
      name: 'Address - Road', logical: 'Address', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/\b(r\.?d\.?|road)\b/gi, 'Road'))
    },
    {
      name: 'Address - Street', logical: 'Address', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/\b(str|st)\.?\b/gi, 'Street'))
    },
    {
      name: 'Address - Town', logical: 'Address', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/\b(twn|tn)\.?\b/gi, 'Town'))
    },
    {
      name: 'Address - Village', logical: 'Address', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/\b(vill|vlg)\.?\b/gi, 'Village'))
    },
    {
      name: 'Address - House Number', logical: 'Address', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/\b(h\.?\s*no\.?|h\/?no\.?|house no\.?)\b/gi, 'House Number'))
    },
    {
      name: 'Address - Colony', logical: 'Address', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/\bcol\.?\b/gi, 'Colony'))
    },
    {
      name: 'Address - Proper case', logical: 'Address', modes: ['account', 'combined'], action: properCase
    },
    {
      name: 'Location - Keep Only Last Value After Final Comma', logical: 'Location', modes: ['account', 'combined'], action: (value) => {
        const parts = normalizeText(value).split(',').map((part) => part.trim()).filter(Boolean);
        return parts.length > 1 ? parts[parts.length - 1] : normalizeText(value);
      }
    },
    {
      name: 'Location - Remove Brackets & Content', logical: 'Location', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/[\(\[\{][^\)\]\}]*[\)\]\}]/g, ' '))
    },
    {
      name: 'Location - Remove Digits', logical: 'Location', modes: ['account', 'combined'], action: (value) => normalizeText(toStringSafe(value).replace(/\d+/g, ''))
    },
    {
      name: 'Location - Clean Leading/Trailing Dots & Commas', logical: 'Location', modes: ['account', 'combined'], action: stripTerminalDotsAndCommas
    },
    {
      name: 'Location - Proper case', logical: 'Location', modes: ['account', 'combined'], action: properCase
    }
  ];

  const validations = [
    { name: 'Pin Code - Missing', logical: 'Pin Code', modes: ['account', 'combined'], validateMissing: true, action: (value) => value },
    { name: 'Job Title - Missing', logical: 'Job Title', modes: ['contact', 'combined'], validateMissing: true, action: (value) => value },
    { name: 'Department - Missing', logical: 'Department', modes: ['contact', 'combined'], validateMissing: true, action: (value) => value },
    { name: 'Mobile Number - Missing', logical: 'Mobile Number', modes: ['contact', 'combined'], validateMissing: true, action: (value) => value },
    { name: 'Phone Number - Missing', logical: 'Phone Number', modes: ['account', 'combined'], validateMissing: true, action: (value) => value },
    { name: 'Fax Number - Missing', logical: 'Fax Number', modes: ['account', 'combined'], validateMissing: true, action: (value) => value },
    { name: 'Pin Code - 6 digits', logical: 'Pin Code', modes: ['account', 'combined'], action: cleanPinCode, valid: (value) => /^\d{6}$/.test(value) },
    { name: 'Phone Number - 11 digits', logical: 'Phone Number', modes: ['account', 'combined'], action: cleanPhone, valid: (value) => /^0\d{10}$/.test(value) },
    { name: 'Mobile Number - 10 digits', logical: 'Mobile Number', modes: ['contact', 'combined'], action: cleanMobile, valid: (value) => /^\d{10}$/.test(value) },
    { name: 'Fax Number - Format/Length', logical: 'Fax Number', modes: ['account', 'combined'], action: cleanPhone, valid: (value) => /^0\d{10}$/.test(value) },
    { name: 'Email - lowercase', logical: 'Email ID', modes: ['contact', 'combined'], action: cleanEmail, valid: (value) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value) }
  ];

  [...ruleGroups, ...validations].forEach((rule) => {
    if (!rule.modes.includes(mode)) return;
    if (ruleFlags[rule.name] !== 'Y') {
      issueSummary[rule.name] = { count: 'Not Checked', corrections: 'Not Checked' };
      return;
    }
    if (!(rule.logical in actualCols)) {
      issueSummary[rule.name] = { count: 'Not Checked', corrections: 'Not Checked' };
      return;
    }

    let issueCount = 0;
    let correctionCount = 0;
    const correctedKey = `${rule.logical}_corrected`;

    df.forEach((row) => {
      const before = toStringSafe(row[correctedKey]);
      const after = rule.action(before);
      const changed = after !== before;
      if (changed) {
        row[correctedKey] = after;
        correctionCount += 1;
      }

      let flagged = false;
      if (rule.validateMissing) {
        flagged = !normalizeText(after);
      } else if (rule.valid) {
        flagged = after !== '' && !rule.valid(after);
      } else {
        flagged = changed;
      }

      if (flagged) {
        row.Issues = appendIssue(row.Issues, rule.name);
        issueCount += 1;
      }
    });

    issueSummary[rule.name] = { count: issueCount, corrections: correctionCount };
  });

  df.forEach((row, index) => {
    row._rowNumber = index + 2;
    row.Issues = normalizeText(row.Issues) || null;
  });

  return { df, issueSummary, actualCols };
}

const DEDUPE_THRESHOLD = 0.67;

function levenshteinDistance(a, b) {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  let prev = new Array(b.length + 1);
  let curr = new Array(b.length + 1);
  for (let j = 0; j <= b.length; j += 1) prev[j] = j;
  for (let i = 1; i <= a.length; i += 1) {
    curr[0] = i;
    for (let j = 1; j <= b.length; j += 1) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      curr[j] = Math.min(prev[j] + 1, curr[j - 1] + 1, prev[j - 1] + cost);
    }
    [prev, curr] = [curr, prev];
  }
  return prev[b.length];
}

function fuzzyRatio(a, b) {
  const left = normalizeText(a).toLowerCase();
  const right = normalizeText(b).toLowerCase();
  if (!left && !right) return 1;
  if (!left || !right) return 0;
  if (left === right) return 1;
  const distance = levenshteinDistance(left, right);
  return 1 - distance / Math.max(left.length, right.length);
}

function tokenSortRatio(a, b) {
  const sortTokens = (value) => normalizeText(value).toLowerCase().split(/\s+/).filter(Boolean).sort().join(' ');
  return fuzzyRatio(sortTokens(a), sortTokens(b));
}

function exactMatch(a, b) {
  const left = normalizeText(a).toLowerCase();
  const right = normalizeText(b).toLowerCase();
  return Boolean(left) && Boolean(right) && left === right;
}

function weightedAverage(components) {
  const active = components.filter((item) => item.available);
  if (!active.length) return 0;
  const totalWeight = active.reduce((sum, item) => sum + item.weight, 0);
  const score = active.reduce((sum, item) => sum + (item.score * item.weight), 0);
  return totalWeight ? score / totalWeight : 0;
}

function getRowLabel(row, index, actualCols, logicalId, fallbackLogicalName) {
  const idField = actualCols[logicalId];
  const id = idField ? normalizeText(row[idField]) : '';
  const nameField = actualCols[fallbackLogicalName] ? `${fallbackLogicalName}_corrected` : '';
  const name = nameField ? normalizeText(row[nameField]) : '';
  return id || name || `Row ${index + 2}`;
}

function getCompletenessFields(mode, actualCols) {
  const accountFields = ['Account Name', 'Address', 'Location', 'Pin Code', 'Phone Number'];
  const contactFields = ['Contact Name', 'Email ID', 'Mobile Number', 'Job Title', 'Department'];
  const combined = mode === 'account' ? accountFields : mode === 'contact' ? contactFields : [...new Set([...accountFields, ...contactFields])];
  return combined.filter((logical) => actualCols[logical]);
}

function countFilledFields(row, logicalFields) {
  return logicalFields.reduce((sum, logical) => sum + (normalizeText(row[`${logical}_corrected`]) ? 1 : 0), 0);
}

function pickMasterRecord(cluster, kind, actualCols) {
  const logicalFields = kind === 'account'
    ? getCompletenessFields('account', actualCols)
    : getCompletenessFields('contact', actualCols);
  const nameField = kind === 'account' ? 'Account Name_corrected' : 'Contact Name_corrected';

  const ranked = [...cluster.members].sort((a, b) => {
    if (b.completenessScore !== a.completenessScore) return b.completenessScore - a.completenessScore;
    const aNameLen = normalizeText(a.row[nameField]).length;
    const bNameLen = normalizeText(b.row[nameField]).length;
    if (bNameLen !== aNameLen) return bNameLen - aNameLen;
    return a.index - b.index;
  });

  const winner = ranked[0];
  return {
    rowNumber: winner.index + 2,
    label: winner.label,
    name: normalizeText(winner.row[nameField]) || winner.label,
    completenessScore: winner.completenessScore,
    maxCompleteness: logicalFields.length,
    reason: `Most complete record (${winner.completenessScore}/${logicalFields.length} populated priority fields)`
  };
}

function buildConnectedComponents(nodes, edges) {
  const adjacency = new Map();
  nodes.forEach((node) => adjacency.set(node.index, []));
  edges.forEach((edge) => {
    adjacency.get(edge.left.index).push(edge);
    adjacency.get(edge.right.index).push({ ...edge, left: edge.right, right: edge.left });
  });

  const visited = new Set();
  const clusters = [];

  nodes.forEach((node) => {
    if (visited.has(node.index)) return;
    const stack = [node.index];
    const memberIndexes = [];
    while (stack.length) {
      const current = stack.pop();
      if (visited.has(current)) continue;
      visited.add(current);
      memberIndexes.push(current);
      adjacency.get(current).forEach((edge) => {
        if (!visited.has(edge.right.index)) stack.push(edge.right.index);
      });
    }
    if (memberIndexes.length > 1) {
      const memberSet = new Set(memberIndexes);
      const clusterEdges = edges.filter((edge) => memberSet.has(edge.left.index) && memberSet.has(edge.right.index));
      clusters.push({ memberIndexes, edges: clusterEdges });
    }
  });

  return clusters;
}
function buildAccountClusterInsights(records, actualCols, mode) {
  if (!(mode === 'account' || mode === 'combined')) return { ran: false, kind: 'account', reason: 'Account mode not selected.', clusters: [], duplicateRows: 0, averageConfidence: 0, duplicateRowSet: new Set() };
  if (!actualCols['Account Name'] || !actualCols['Pin Code']) {
    return { ran: false, kind: 'account', reason: 'Account Name and Pin Code are required for account clustering.', clusters: [], duplicateRows: 0, averageConfidence: 0, duplicateRowSet: new Set() };
  }

  const nodes = records.map((row, index) => ({
    row,
    index,
    label: getRowLabel(row, index, actualCols, 'Account Record ID', 'Account Name'),
    completenessScore: countFilledFields(row, getCompletenessFields('account', actualCols))
  }));

  const buckets = new Map();
  nodes.forEach((node) => {
    const pin = normalizeText(node.row['Pin Code_corrected']);
    if (!pin) return;
    if (!buckets.has(pin)) buckets.set(pin, []);
    buckets.get(pin).push(node);
  });

  const edges = [];
  buckets.forEach((bucket) => {
    for (let i = 0; i < bucket.length; i += 1) {
      for (let j = i + 1; j < bucket.length; j += 1) {
        const left = bucket[i];
        const right = bucket[j];
        const nameScore = tokenSortRatio(left.row['Account Name_corrected'], right.row['Account Name_corrected']);
        if (nameScore < DEDUPE_THRESHOLD) continue;
        const addressScore = actualCols['Address'] ? tokenSortRatio(left.row['Address_corrected'], right.row['Address_corrected']) : 0;
        const locationScore = actualCols['Location'] ? fuzzyRatio(left.row['Location_corrected'], right.row['Location_corrected']) : 0;
        const phoneScore = actualCols['Phone Number'] ? (exactMatch(left.row['Phone Number_corrected'], right.row['Phone Number_corrected']) ? 1 : 0) : 0;
        const confidence = weightedAverage([
          { available: true, weight: 0.58, score: nameScore },
          { available: Boolean(actualCols['Address']), weight: 0.2, score: addressScore },
          { available: Boolean(actualCols['Location']), weight: 0.1, score: locationScore },
          { available: Boolean(actualCols['Phone Number']), weight: 0.12, score: phoneScore }
        ]);
        edges.push({ left, right, confidence });
      }
    }
  });

  const components = buildConnectedComponents(nodes, edges);
  const duplicateRowSet = new Set();
  const clusters = components.map((component, clusterIndex) => {
    const members = component.memberIndexes.map((index) => nodes[index]).sort((a, b) => a.index - b.index);
    members.forEach((member) => duplicateRowSet.add(member.index));
    const avgConfidence = component.edges.length
      ? component.edges.reduce((sum, edge) => sum + edge.confidence, 0) / component.edges.length
      : DEDUPE_THRESHOLD;
    const master = pickMasterRecord({ members }, 'account', actualCols);
    return {
      id: `ACC-${String(clusterIndex + 1).padStart(3, '0')}`,
      kind: 'account',
      typeLabel: 'Account',
      size: members.length,
      averageConfidence: avgConfidence,
      master,
      members,
      edgeCount: component.edges.length,
      topSignals: 'Same pin code + fuzzy account name (>67%)'
    };
  }).sort((a, b) => b.size - a.size || b.averageConfidence - a.averageConfidence);

  clusters.forEach((cluster) => {
    cluster.members.forEach((member) => {
      const row = member.row;
      row.Account_Cluster_ID = cluster.id;
      row.Account_Cluster_Size = cluster.size;
      row.Account_Duplicate_Confidence = `${Math.round(cluster.averageConfidence * 100)}%`;
      row.Account_Master_Suggestion = cluster.master.label;
      row.Account_Cluster_Members = cluster.members.map((item) => item.label).join(', ');
      row.Account_Cluster_Status = member.label === cluster.master.label ? 'Suggested Master' : 'Duplicate In Cluster';
    });
  });

  const averageConfidence = clusters.length
    ? clusters.reduce((sum, cluster) => sum + cluster.averageConfidence, 0) / clusters.length
    : 0;

  return {
    ran: true,
    kind: 'account',
    clusters,
    duplicateRows: duplicateRowSet.size,
    averageConfidence,
    duplicateRowSet,
    reason: ''
  };
}

function buildContactClusterInsights(records, actualCols, mode) {
  if (!(mode === 'contact' || mode === 'combined')) return { ran: false, kind: 'contact', reason: 'Contact mode not selected.', clusters: [], duplicateRows: 0, averageConfidence: 0, duplicateRowSet: new Set() };
  if (!actualCols['Contact Name'] || (!actualCols['Email ID'] && !actualCols['Mobile Number'])) {
    return { ran: false, kind: 'contact', reason: 'Contact Name plus Email ID or Mobile Number are required for contact clustering.', clusters: [], duplicateRows: 0, averageConfidence: 0, duplicateRowSet: new Set() };
  }

  const nodes = records.map((row, index) => ({
    row,
    index,
    label: getRowLabel(row, index, actualCols, 'Contact Record ID', 'Contact Name'),
    completenessScore: countFilledFields(row, getCompletenessFields('contact', actualCols))
  }));

  const edges = [];
  const seenPairs = new Set();
  const maybeAddEdge = (left, right) => {
    const key = `${Math.min(left.index, right.index)}::${Math.max(left.index, right.index)}`;
    if (seenPairs.has(key)) return;
    seenPairs.add(key);
    const nameScore = tokenSortRatio(left.row['Contact Name_corrected'], right.row['Contact Name_corrected']);
    if (nameScore < DEDUPE_THRESHOLD) return;
    const emailScore = actualCols['Email ID'] ? (exactMatch(left.row['Email ID_corrected'], right.row['Email ID_corrected']) ? 1 : 0) : 0;
    const mobileScore = actualCols['Mobile Number'] ? (exactMatch(left.row['Mobile Number_corrected'], right.row['Mobile Number_corrected']) ? 1 : 0) : 0;
    if (!emailScore && !mobileScore) return;
    const confidence = weightedAverage([
      { available: true, weight: 0.45, score: nameScore },
      { available: Boolean(actualCols['Email ID']), weight: 0.3, score: emailScore },
      { available: Boolean(actualCols['Mobile Number']), weight: 0.25, score: mobileScore }
    ]);
    edges.push({ left, right, confidence });
  };

  ['Email ID_corrected', 'Mobile Number_corrected'].forEach((fieldKey) => {
    if (!records.some((row) => fieldKey in row)) return;
    const buckets = new Map();
    nodes.forEach((node) => {
      const value = normalizeText(node.row[fieldKey]).toLowerCase();
      if (!value) return;
      if (!buckets.has(value)) buckets.set(value, []);
      buckets.get(value).push(node);
    });
    buckets.forEach((bucket) => {
      for (let i = 0; i < bucket.length; i += 1) {
        for (let j = i + 1; j < bucket.length; j += 1) maybeAddEdge(bucket[i], bucket[j]);
      }
    });
  });

  const components = buildConnectedComponents(nodes, edges);
  const duplicateRowSet = new Set();
  const clusters = components.map((component, clusterIndex) => {
    const members = component.memberIndexes.map((index) => nodes[index]).sort((a, b) => a.index - b.index);
    members.forEach((member) => duplicateRowSet.add(member.index));
    const avgConfidence = component.edges.length
      ? component.edges.reduce((sum, edge) => sum + edge.confidence, 0) / component.edges.length
      : DEDUPE_THRESHOLD;
    const master = pickMasterRecord({ members }, 'contact', actualCols);
    return {
      id: `CON-${String(clusterIndex + 1).padStart(3, '0')}`,
      kind: 'contact',
      typeLabel: 'Contact',
      size: members.length,
      averageConfidence: avgConfidence,
      master,
      members,
      edgeCount: component.edges.length,
      topSignals: 'Fuzzy contact name + exact email/mobile match'
    };
  }).sort((a, b) => b.size - a.size || b.averageConfidence - a.averageConfidence);

  clusters.forEach((cluster) => {
    cluster.members.forEach((member) => {
      const row = member.row;
      row.Contact_Cluster_ID = cluster.id;
      row.Contact_Cluster_Size = cluster.size;
      row.Contact_Duplicate_Confidence = `${Math.round(cluster.averageConfidence * 100)}%`;
      row.Contact_Master_Suggestion = cluster.master.label;
      row.Contact_Cluster_Members = cluster.members.map((item) => item.label).join(', ');
      row.Contact_Cluster_Status = member.label === cluster.master.label ? 'Suggested Master' : 'Duplicate In Cluster';
    });
  });

  const averageConfidence = clusters.length
    ? clusters.reduce((sum, cluster) => sum + cluster.averageConfidence, 0) / clusters.length
    : 0;

  return {
    ran: true,
    kind: 'contact',
    clusters,
    duplicateRows: duplicateRowSet.size,
    averageConfidence,
    duplicateRowSet,
    reason: ''
  };
}

function buildCompletenessHeatmap(records, mode, actualCols) {
  const logicalFields = getCompletenessFields(mode, actualCols);
  return logicalFields.map((logical) => {
    const correctedKey = `${logical}_corrected`;
    const filled = records.filter((row) => normalizeText(row[correctedKey])).length;
    const pct = records.length ? filled / records.length : 0;
    return { logical, filled, total: records.length, percentage: pct };
  }).sort((a, b) => b.percentage - a.percentage || a.logical.localeCompare(b.logical));
}

function calculateGeography(records, actualCols, duplicateRowSet) {
  const buildBreakdown = (logical, correctedKey) => {
    if (!actualCols[logical]) return [];
    const map = new Map();
    records.forEach((row, index) => {
      const value = normalizeText(row[correctedKey]) || 'Missing';
      if (!map.has(value)) map.set(value, { label: value, rows: 0, duplicateRows: 0 });
      const entry = map.get(value);
      entry.rows += 1;
      if (duplicateRowSet.has(index)) entry.duplicateRows += 1;
    });
    return [...map.values()].sort((a, b) => b.rows - a.rows || a.label.localeCompare(b.label)).slice(0, 10);
  };

  return {
    byLocation: buildBreakdown('Location', 'Location_corrected'),
    byPin: buildBreakdown('Pin Code', 'Pin Code_corrected')
  };
}

function calculateOrphans(records, mode, actualCols) {
  const accountSignals = ['Account Record ID', 'Account Name', 'Address', 'Pin Code'].filter((logical) => actualCols[logical]);
  const contactSignals = ['Contact Record ID', 'Contact Name', 'Email ID', 'Mobile Number'].filter((logical) => actualCols[logical]);

  const orphanAccounts = [];
  const orphanContacts = [];

  if (mode === 'combined') {
    records.forEach((row, index) => {
      const hasAccount = accountSignals.some((logical) => normalizeText(row[`${logical}_corrected`])) || (actualCols['Account Record ID'] && normalizeText(row[actualCols['Account Record ID']]));
      const hasContact = contactSignals.some((logical) => normalizeText(row[`${logical}_corrected`])) || (actualCols['Contact Record ID'] && normalizeText(row[actualCols['Contact Record ID']]));
      if (hasAccount && !hasContact) orphanAccounts.push({ row, index });
      if (hasContact && !hasAccount) orphanContacts.push({ row, index });
    });
  }

  return { orphanAccounts, orphanContacts, total: orphanAccounts.length + orphanContacts.length };
}

function buildAnalytics(records, issueSummary, mode, actualCols, accountClusters, contactClusters, completeness, geography, orphans) {
  const numericIssues = Object.entries(issueSummary).filter(([_, entry]) => typeof entry.count === 'number').sort((a, b) => b[1].count - a[1].count);
  const issueRows = records.filter((row) => row.Issues).length;
  const totalIssues = numericIssues.reduce((sum, [_, entry]) => sum + entry.count, 0);
  const duplicateRowSet = new Set([...accountClusters.duplicateRowSet, ...contactClusters.duplicateRowSet]);
  const completenessAverage = completeness.length
    ? completeness.reduce((sum, item) => sum + item.percentage, 0) / completeness.length
    : 0;
  const qualityScore = Math.max(0, Math.min(100,
    Math.round((
      (completenessAverage * 0.55) +
      ((records.length ? (1 - (issueRows / records.length)) : 1) * 0.25) +
      ((records.length ? (1 - (duplicateRowSet.size / records.length)) : 1) * 0.2)
    ) * 100)
  ));

  const allClusters = [...accountClusters.clusters, ...contactClusters.clusters].sort((a, b) => b.size - a.size || b.averageConfidence - a.averageConfidence);
  const overallConfidence = allClusters.length
    ? allClusters.reduce((sum, cluster) => sum + cluster.averageConfidence, 0) / allClusters.length
    : 0;

  return {
    mode,
    totalRecords: records.length,
    issueRows,
    totalIssues,
    topIssues: numericIssues.slice(0, 8),
    qualityScore,
    overallConfidence,
    duplicateRows: duplicateRowSet.size,
    duplicateRowSet,
    clusters: allClusters,
    clusterLookup: Object.fromEntries(allClusters.map((cluster) => [cluster.id, cluster])),
    accountClusters,
    contactClusters,
    completeness,
    geography,
    orphans
  };
}

function scoreBadgeClass(value) {
  if (value >= 85) return 'status-good';
  if (value >= 65) return 'status-warn';
  return 'status-bad';
}

function showProgressOverlay() {
  if (!progressOverlay) return;
  progressOverlay.classList.remove('hidden');
  Array.from(progressStepList.children).forEach((item) => item.classList.remove('active', 'done'));
  progressFill.style.width = '4%';
}

function hideProgressOverlay() {
  progressOverlay.classList.add('hidden');
}

function setProcessStep(index) {
  const step = PROGRESS_STEPS[index];
  if (!step) return;
  progressTitle.textContent = step.title;
  progressFootnote.textContent = step.footnote;
  progressFill.style.width = `${step.fill}%`;
  Array.from(progressStepList.children).forEach((item, position) => {
    item.classList.toggle('active', position === index);
    item.classList.toggle('done', position < index);
  });
}

function completeProgressOverlay() {
  progressFill.style.width = '100%';
  progressTitle.textContent = 'All done!';
  progressFootnote.textContent = 'Insights are ready below.';
  Array.from(progressStepList.children).forEach((item) => {
    item.classList.remove('active');
    item.classList.add('done');
  });
  return new Promise((resolve) => setTimeout(() => {
    hideProgressOverlay();
    resolve();
  }, 420));
}

function pause(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function escapeHtml(value) {
  return toStringSafe(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function openDetailOverlay(eyebrow, title, bodyHtml) {
  detailEyebrow.textContent = eyebrow;
  detailTitle.textContent = title;
  detailBody.innerHTML = bodyHtml;
  detailOverlay.classList.remove('hidden');
}

function closeDetailOverlay() {
  detailOverlay.classList.add('hidden');
}
function buildIssueDetailHtml(issueName, entry) {
  if (!lastRunData) return '<p class="detail-empty">No data available.</p>';
  const affected = lastRunData.records.filter((row) => normalizeText(row.Issues).split(',').map((part) => part.trim()).includes(issueName));
  const rows = affected.slice(0, 25).map((row) => `<tr><td>${row._rowNumber}</td><td>${escapeHtml(row[`${issueName.split(' - ')[0]}_corrected`] || row.Issues || '')}</td><td>${escapeHtml(row.Issues || '')}</td></tr>`).join('');
  return `
    <div class="detail-stats">
      <div class="detail-stat"><strong>${entry.count}</strong><span>Rows flagged</span></div>
      <div class="detail-stat"><strong>${entry.corrections}</strong><span>Corrections applied</span></div>
      <div class="detail-stat"><strong>${lastRunData.records.length}</strong><span>Total rows scanned</span></div>
    </div>
    <div class="detail-table-wrap">
      <table>
        <thead><tr><th>Row</th><th>Current value</th><th>Issues</th></tr></thead>
        <tbody>${rows || '<tr><td colspan="3">No matching rows found.</td></tr>'}</tbody>
      </table>
    </div>`;
}

function buildClusterDetailHtml(clusterId) {
  if (!lastRunData) return '<p class="detail-empty">No data available.</p>';
  const cluster = lastRunData.analytics.clusterLookup[clusterId];
  if (!cluster) return '<p class="detail-empty">Cluster not found.</p>';
  const rows = cluster.members.map((member) => {
    const row = member.row;
    const name = cluster.kind === 'account' ? row['Account Name_corrected'] : row['Contact Name_corrected'];
    const location = row['Location_corrected'] || row['Pin Code_corrected'] || row['Email ID_corrected'] || row['Mobile Number_corrected'] || '—';
    const masterBadge = member.label === cluster.master.label ? '<span class="badge-good">Suggested master</span>' : '<span class="badge-neutral">Cluster member</span>';
    return `<tr><td>${member.index + 2}</td><td>${escapeHtml(member.label)}</td><td>${escapeHtml(name || '—')}</td><td>${escapeHtml(location)}</td><td>${masterBadge}</td></tr>`;
  }).join('');

  return `
    <div class="detail-stats">
      <div class="detail-stat"><strong>${cluster.size}</strong><span>Rows in cluster</span></div>
      <div class="detail-stat"><strong>${Math.round(cluster.averageConfidence * 100)}%</strong><span>Average confidence</span></div>
      <div class="detail-stat"><strong>${escapeHtml(cluster.master.label)}</strong><span>Suggested master</span></div>
    </div>
    <p><strong>Why grouped:</strong> ${escapeHtml(cluster.topSignals)}.</p>
    <p><strong>Master rule:</strong> ${escapeHtml(cluster.master.reason)}.</p>
    <div class="detail-table-wrap">
      <table>
        <thead><tr><th>Row</th><th>Record label</th><th>Name</th><th>Supporting field</th><th>Status</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;
}

function buildOrphanDetailHtml(kind) {
  if (!lastRunData) return '<p class="detail-empty">No data available.</p>';
  const list = kind === 'account' ? lastRunData.analytics.orphans.orphanAccounts : lastRunData.analytics.orphans.orphanContacts;
  const title = kind === 'account' ? 'Account-side rows missing contact data' : 'Contact-side rows missing account data';
  const rows = list.slice(0, 30).map(({ row, index }) => {
    const accountName = row['Account Name_corrected'] || '—';
    const contactName = row['Contact Name_corrected'] || '—';
    return `<tr><td>${index + 2}</td><td>${escapeHtml(accountName)}</td><td>${escapeHtml(contactName)}</td></tr>`;
  }).join('');
  return `
    <div class="detail-stats">
      <div class="detail-stat"><strong>${list.length}</strong><span>Rows found</span></div>
      <div class="detail-stat"><strong>${lastRunData.analytics.totalRecords}</strong><span>Total rows scanned</span></div>
      <div class="detail-stat"><strong>${escapeHtml(title)}</strong><span>Definition</span></div>
    </div>
    <div class="detail-table-wrap">
      <table>
        <thead><tr><th>Row</th><th>Account</th><th>Contact</th></tr></thead>
        <tbody>${rows || '<tr><td colspan="3">No orphaned rows found.</td></tr>'}</tbody>
      </table>
    </div>`;
}

function renderSummaryCards(analytics) {
  summaryGrid.innerHTML = '';
  const cards = [
    { label: 'Mode', value: analytics.mode.charAt(0).toUpperCase() + analytics.mode.slice(1) },
    { label: 'Rows processed', value: analytics.totalRecords },
    { label: 'Data health score', value: `${analytics.qualityScore}%`, statusClass: scoreBadgeClass(analytics.qualityScore) },
    { label: 'Duplicate confidence', value: `${Math.round(analytics.overallConfidence * 100)}%`, statusClass: analytics.overallConfidence >= 0.8 ? 'status-good' : analytics.overallConfidence >= 0.67 ? 'status-warn' : 'status-bad' },
    { label: 'Rows with issues', value: analytics.issueRows }
  ];

  cards.forEach((card) => {
    const block = document.createElement('div');
    block.className = `summary-card ${card.statusClass || ''}`.trim();
    block.innerHTML = `<strong>${card.value}</strong><span>${card.label}</span>`;
    summaryGrid.appendChild(block);
  });
}

function renderClusterSummary(analytics) {
  dedupeGrid.innerHTML = '';
  clusterTableBody.innerHTML = '';
  const loadMoreContainer = document.getElementById('clusterLoadMoreContainer');
  const loadMoreButton = document.getElementById('clusterLoadMore');
  
  clusterSubtext.textContent = analytics.clusters.length
    ? `${analytics.clusters.length} cluster(s) found across the uploaded file.`
    : 'No duplicate clusters detected with the selected logic.';

  const cards = [];
  if (analytics.accountClusters.ran) {
    cards.push({ label: 'Account clusters', value: analytics.accountClusters.clusters.length });
    cards.push({ label: 'Account duplicate rows', value: analytics.accountClusters.duplicateRows });
  }
  if (analytics.contactClusters.ran) {
    cards.push({ label: 'Contact clusters', value: analytics.contactClusters.clusters.length });
    cards.push({ label: 'Contact duplicate rows', value: analytics.contactClusters.duplicateRows });
  }
  cards.push({ label: 'Overall duplicate rows', value: analytics.duplicateRows });
  cards.push({ label: 'Avg confidence', value: `${Math.round(analytics.overallConfidence * 100)}%` });

  cards.slice(0, 6).forEach((card) => {
    const block = document.createElement('div');
    block.className = 'summary-card';
    block.innerHTML = `<strong>${card.value}</strong><span>${card.label}</span>`;
    dedupeGrid.appendChild(block);
  });

  if (!analytics.clusters.length) {
    clusterTableBody.innerHTML = '<tr><td colspan="6">No clusters were detected.</td></tr>';
    loadMoreContainer.style.display = 'none';
    return;
  }

  // Initialize pagination state
  clusterPaginationState.allClusters = analytics.clusters;
  clusterPaginationState.displayed = 0;
  
  // Display first batch of records
  displayClusterBatch();
}

function displayClusterBatch() {
  const startIdx = clusterPaginationState.displayed;
  const endIdx = Math.min(startIdx + clusterPaginationState.pageSize, clusterPaginationState.allClusters.length);
  const loadMoreContainer = document.getElementById('clusterLoadMoreContainer');
  const loadMoreButton = document.getElementById('clusterLoadMore');
  
  for (let i = startIdx; i < endIdx; i++) {
    const cluster = clusterPaginationState.allClusters[i];
    const row = document.createElement('tr');
    row.className = 'clickable-row';
    row.innerHTML = `
      <td>${cluster.id}</td>
      <td>${cluster.typeLabel}</td>
      <td>${cluster.size}</td>
      <td>${Math.round(cluster.averageConfidence * 100)}%</td>
      <td>${escapeHtml(cluster.master.label)}</td>
      <td class="row-action">View details ›</td>`;
    row.addEventListener('click', () => openDetailOverlay('Duplicate cluster', cluster.id, buildClusterDetailHtml(cluster.id)));
    clusterTableBody.appendChild(row);
  }
  
  clusterPaginationState.displayed = endIdx;
  
  // Show or hide load more button
  if (endIdx >= clusterPaginationState.allClusters.length) {
    loadMoreContainer.style.display = 'none';
  } else {
    loadMoreContainer.style.display = 'flex';
    loadMoreButton.onclick = displayClusterBatch;
  }
}

function renderHeatmap(analytics) {
  heatmapGrid.innerHTML = '';
  analytics.completeness.forEach((item) => {
    const card = document.createElement('div');
    card.className = 'heatmap-card';
    card.innerHTML = `
      <div class="heatmap-top">
        <div>
          <h4>${escapeHtml(item.logical)}</h4>
          <p>${item.filled} of ${item.total} rows populated</p>
        </div>
        <strong>${Math.round(item.percentage * 100)}%</strong>
      </div>
      <div class="heat-track"><div class="heat-fill" style="width:${Math.round(item.percentage * 100)}%"></div></div>`;
    heatmapGrid.appendChild(card);
  });
}

function renderGeography(analytics) {
  geoLocationBody.innerHTML = analytics.geography.byLocation.length
    ? analytics.geography.byLocation.map((entry) => `<tr><td>${escapeHtml(entry.label)}</td><td>${entry.rows}</td><td>${entry.duplicateRows}</td></tr>`).join('')
    : '<tr><td colspan="3">Location field not available.</td></tr>';
  geoPinBody.innerHTML = analytics.geography.byPin.length
    ? analytics.geography.byPin.map((entry) => `<tr><td>${escapeHtml(entry.label)}</td><td>${entry.rows}</td><td>${entry.duplicateRows}</td></tr>`).join('')
    : '<tr><td colspan="3">Pin Code field not available.</td></tr>';
}

function renderOrphans(analytics) {
  orphanGrid.innerHTML = '';
  orphanTableBody.innerHTML = '';
  const cards = [
    { label: 'Orphaned accounts', value: analytics.orphans.orphanAccounts.length },
    { label: 'Orphaned contacts', value: analytics.orphans.orphanContacts.length },
    { label: 'Total orphan rows', value: analytics.orphans.total }
  ];
  cards.forEach((card) => {
    const block = document.createElement('div');
    block.className = `summary-card ${card.value ? 'status-warn' : 'status-good'}`;
    block.innerHTML = `<strong>${card.value}</strong><span>${card.label}</span>`;
    orphanGrid.appendChild(block);
  });

  const rows = [
    { kind: 'account', label: 'Account rows missing contact-side data', count: analytics.orphans.orphanAccounts.length },
    { kind: 'contact', label: 'Contact rows missing account-side data', count: analytics.orphans.orphanContacts.length }
  ];
  orphanTableBody.innerHTML = rows.map((item) => `
    <tr class="clickable-row" data-kind="${item.kind}">
      <td>${escapeHtml(item.label)}</td>
      <td>${item.count}</td>
      <td class="row-action">View details ›</td>
    </tr>`).join('');
  orphanTableBody.querySelectorAll('tr[data-kind]').forEach((row) => {
    row.addEventListener('click', () => {
      const kind = row.getAttribute('data-kind');
      openDetailOverlay('Orphan overview', kind === 'account' ? 'Orphaned account-side rows' : 'Orphaned contact-side rows', buildOrphanDetailHtml(kind));
    });
  });
}

function renderIssues(analytics) {
  issueTableBody.innerHTML = '';
  if (!analytics.topIssues.length) {
    issueTableBody.innerHTML = '<tr><td colspan="3">No issues were detected.</td></tr>';
    return;
  }
  analytics.topIssues.forEach(([name, entry]) => {
    const row = document.createElement('tr');
    row.className = 'clickable-row';
    row.innerHTML = `<td>${escapeHtml(name)}</td><td>${entry.count}</td><td class="row-action">View details ›</td>`;
    row.addEventListener('click', () => openDetailOverlay('Issue detail', name, buildIssueDetailHtml(name, entry)));
    issueTableBody.appendChild(row);
  });
}

function renderAnalytics(analytics) {
  renderSummaryCards(analytics);
  renderClusterSummary(analytics);
  renderHeatmap(analytics);
  renderGeography(analytics);
  renderOrphans(analytics);
  renderIssues(analytics);
  analyticsCard.classList.remove('hidden');
  analyticsCard.classList.add('reveal');
}

function addSheetFromObjects(workbook, sheetName, rows) {
  const sheet = workbook.addWorksheet(sheetName);
  if (!rows.length) {
    sheet.addRow(['No data available']);
    return sheet;
  }
  const headers = Object.keys(rows[0]);
  sheet.columns = headers.map((key) => ({ header: key, key }));
  rows.forEach((row) => sheet.addRow(row));
  styleHeaderRow(sheet);
  sheet.views = [{ state: 'frozen', ySplit: 1 }];
  sheet.autoFilter = { from: { row: 1, col: 1 }, to: { row: sheet.rowCount, col: headers.length } };
  autoSizeColumns(sheet);
  return sheet;
}

function buildReportWorkbook(records, issueSummary, analytics) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Account & Contact Quality Check';
  workbook.created = new Date();

  addSheetFromObjects(workbook, 'Executive Summary', [
    { Metric: 'Mode', Value: analytics.mode },
    { Metric: 'Rows processed', Value: analytics.totalRecords },
    { Metric: 'Data health score', Value: `${analytics.qualityScore}%` },
    { Metric: 'Duplicate confidence', Value: `${Math.round(analytics.overallConfidence * 100)}%` },
    { Metric: 'Rows with issues', Value: analytics.issueRows },
    { Metric: 'Total issues', Value: analytics.totalIssues },
    { Metric: 'Account clusters', Value: analytics.accountClusters.clusters.length },
    { Metric: 'Contact clusters', Value: analytics.contactClusters.clusters.length },
    { Metric: 'Total orphan rows', Value: analytics.orphans.total },
    { Metric: 'Master selection rule', Value: 'Most complete record, then longest descriptive name, then top row order' },
    { Metric: 'Account cluster rule', Value: 'Same Pin Code + fuzzy Account Name match (>67%)' }
  ]);

  addSheetFromObjects(workbook, 'Issue Summary', Object.entries(issueSummary).map(([issue, entry]) => ({
    Issue: issue,
    Count: entry.count,
    Corrections: entry.corrections
  })));

  addSheetFromObjects(workbook, 'Duplicate Clusters', analytics.clusters.map((cluster) => ({
    Cluster: cluster.id,
    Type: cluster.typeLabel,
    Size: cluster.size,
    'Average Confidence %': Math.round(cluster.averageConfidence * 100),
    'Suggested Master': cluster.master.label,
    'Master Reason': cluster.master.reason,
    Members: cluster.members.map((member) => member.label).join(', ')
  })));

  addSheetFromObjects(workbook, 'Cluster Members', analytics.clusters.flatMap((cluster) => cluster.members.map((member) => ({
    Cluster: cluster.id,
    Type: cluster.typeLabel,
    Row: member.index + 2,
    'Record Label': member.label,
    'Suggested Master': cluster.master.label,
    'Cluster Confidence %': Math.round(cluster.averageConfidence * 100),
    Status: member.label === cluster.master.label ? 'Suggested Master' : 'Duplicate In Cluster'
  }))));

  addSheetFromObjects(workbook, 'Field Completeness', analytics.completeness.map((item) => ({
    Field: item.logical,
    'Filled Rows': item.filled,
    'Total Rows': item.total,
    'Completeness %': Math.round(item.percentage * 100)
  })));

  addSheetFromObjects(workbook, 'Geography - Location', analytics.geography.byLocation.map((item) => ({
    Location: item.label,
    Rows: item.rows,
    'Duplicate Rows': item.duplicateRows
  })));

  addSheetFromObjects(workbook, 'Geography - Pin', analytics.geography.byPin.map((item) => ({
    'Pin Code': item.label,
    Rows: item.rows,
    'Duplicate Rows': item.duplicateRows
  })));

  addSheetFromObjects(workbook, 'Orphan Records', [
    ...analytics.orphans.orphanAccounts.map(({ row, index }) => ({ Type: 'Account side missing contact data', Row: index + 2, Account: row['Account Name_corrected'] || '', Contact: row['Contact Name_corrected'] || '' })),
    ...analytics.orphans.orphanContacts.map(({ row, index }) => ({ Type: 'Contact side missing account data', Row: index + 2, Account: row['Account Name_corrected'] || '', Contact: row['Contact Name_corrected'] || '' }))
  ]);

  const cleanedHeaders = Array.from(new Set(records.flatMap((row) => Object.keys(row)).filter((key) => !key.startsWith('_'))));
  const cleanedSheet = workbook.addWorksheet('Cleaned Data');
  cleanedSheet.columns = cleanedHeaders.map((key) => ({ header: key, key }));
  records.forEach((row) => {
    const output = {};
    cleanedHeaders.forEach((key) => {
      output[key] = row[key] == null ? '' : row[key];
    });
    cleanedSheet.addRow(output);
  });
  styleHeaderRow(cleanedSheet);
  cleanedSheet.views = [{ state: 'frozen', ySplit: 1 }];
  cleanedSheet.autoFilter = { from: { row: 1, col: 1 }, to: { row: cleanedSheet.rowCount, col: cleanedHeaders.length } };
  autoSizeColumns(cleanedSheet);

  return workbook;
}

function getSelectedMode() {
  const selected = document.querySelector('input[name="fileMode"]:checked');
  return selected ? selected.value : 'account';
}

function updateFileControls() {
  const mode = getSelectedMode();
  accountUpload.classList.toggle('hidden', mode !== 'account');
  contactUpload.classList.toggle('hidden', mode !== 'contact');
  combinedUpload.classList.toggle('hidden', mode !== 'combined');
  document.querySelectorAll('.radio-card').forEach((card) => {
    const input = card.querySelector('input');
    card.classList.toggle('selected', input.checked);
  });
}

function getSelectedFile(mode) {
  if (mode === 'account') return document.getElementById('accountFile').files[0];
  if (mode === 'contact') return document.getElementById('contactFile').files[0];
  return document.getElementById('combinedFile').files[0];
}

async function runComparison(event) {
  event.preventDefault();
  const mode = getSelectedMode();
  const file = getSelectedFile(mode);
  if (!file) {
    setStatus('Upload the file required for the selected mode.');
    return;
  }

  analyticsCard.classList.add('hidden');
  analyticsCard.classList.remove('reveal');
  setStatus('Reading workbook...');
  showProgressOverlay();
  setProcessStep(0);

  try {
    await pause(150);
    setProcessStep(1);
    const workbook = await readFileAsWorkbook(file);
    const records = workbookToJsonArray(workbook);
    if (!records.length) throw new Error('The uploaded file does not contain any data rows.');

    setStatus('Cleaning data and validating fields...');
    await pause(120);
    setProcessStep(2);
    const columnMapping = buildColumnMapping();
    const ruleFlags = buildRuleFlags();
    const { df, issueSummary, actualCols } = cleanData(records, ruleFlags, columnMapping, mode);

    setStatus('Grouping duplicate clusters and computing insights...');
    await pause(120);
    setProcessStep(3);
    const accountClusters = buildAccountClusterInsights(df, actualCols, mode);
    const contactClusters = buildContactClusterInsights(df, actualCols, mode);
    const duplicateRowSet = new Set([...accountClusters.duplicateRowSet, ...contactClusters.duplicateRowSet]);
    const completeness = buildCompletenessHeatmap(df, mode, actualCols);
    const geography = calculateGeography(df, actualCols, duplicateRowSet);
    const orphans = calculateOrphans(df, mode, actualCols);
    const analytics = buildAnalytics(df, issueSummary, mode, actualCols, accountClusters, contactClusters, completeness, geography, orphans);

    setStatus('Building the management report...');
    await pause(120);
    setProcessStep(4);
    const reportWorkbook = buildReportWorkbook(df, issueSummary, analytics);
    lastExportWorkbook = reportWorkbook;
    lastExportFilename = `AC_Quality_Report_${mode}_${new Date().toISOString().slice(0, 10)}.xlsx`;
    lastRunData = { records: df, issueSummary, actualCols, analytics, mode };

    await completeProgressOverlay();
    renderAnalytics(analytics);
    setStatus('Analysis complete. Review the insights above or download the report.');
  } catch (error) {
    hideProgressOverlay();
    setStatus(`Error: ${error.message}`);
    console.error(error);
  }
}

async function downloadReport() {
  if (!lastExportWorkbook) {
    setStatus('Run the analysis first to generate the report.');
    return;
  }
  await saveWorkbookToFile(lastExportWorkbook, lastExportFilename);
}

compareForm.addEventListener('submit', runComparison);
document.querySelectorAll('input[name="fileMode"]').forEach((radio) => radio.addEventListener('change', updateFileControls));
downloadReportButton.addEventListener('click', downloadReport);
if (detailClose) detailClose.addEventListener('click', closeDetailOverlay);
if (detailOverlay) {
  detailOverlay.addEventListener('click', (event) => {
    if (event.target === detailOverlay) closeDetailOverlay();
  });
}
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape') closeDetailOverlay();
});
updateFileControls();
setStatus('Ready to run your cleanup.');
