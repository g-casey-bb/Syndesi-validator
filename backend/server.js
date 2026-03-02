const express = require('express');
const cors = require('cors');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadDir),
  filename: (req, file, cb) => {
    const unique = Date.now() + '-' + Math.round(Math.random() * 1e9);
    cb(null, unique + (path.extname(file.originalname) || '.xlsx'));
  }
});

const upload = multer({
  storage,
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ext = (path.extname(file.originalname) || '').toLowerCase();
    if (['.xlsx', '.xls'].includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error('Only Excel files (.xlsx, .xls) are allowed'));
    }
  }
});

function normalizeHeader(str) {
  if (str == null || typeof str !== 'string') return '';
  return String(str).trim().toLowerCase().replace(/\s+/g, ' ');
}

function findColumnIndex(headers, ...possibleNames) {
  const normalized = possibleNames.map(normalizeHeader);
  for (let i = 0; i < headers.length; i++) {
    const h = normalizeHeader(headers[i]);
    if (normalized.some(n => h === n || h.includes(n) || n.includes(h))) return i;
  }
  return -1;
}

/** Find Employee Number column only if header contains "employee" or "emp" to avoid matching "No." or "#". */
function findEmployeeNumberIndex(headers) {
  const mustContain = ['employee', 'emp'];
  const possibleNames = ['Employee Number', 'Employee Number*', 'Employee No', 'EmployeeNo', 'Emp No', 'Emp Number'];
  const normalized = possibleNames.map(normalizeHeader);
  for (let i = 0; i < headers.length; i++) {
    const h = normalizeHeader(headers[i]);
    const nameMatch = normalized.some(n => h === n || h.includes(n) || n.includes(h));
    const hasContext = mustContain.some(c => h.includes(c));
    if (nameMatch && hasContext) return i;
  }
  return -1;
}

/** Find Employee ID column; prefer headers containing "employee" or "emp" so we don't match a bare "ID" or "No." column. */
function findEmployeeIdIndex(headers) {
  const mustContain = ['employee', 'emp'];
  const possibleNames = ['Employee ID', 'EmployeeID', 'Employee Id', 'Emp ID'];
  const normalized = possibleNames.map(normalizeHeader);
  for (let i = 0; i < headers.length; i++) {
    const h = normalizeHeader(headers[i]);
    const nameMatch = normalized.some(n => h === n || h.includes(n) || n.includes(h));
    const hasContext = mustContain.some(c => h.includes(c));
    if (nameMatch && hasContext) return i;
  }
  return findColumnIndex(headers, 'ID');
}

function isColumnSequentialFromOne(dataRows, colIndex) {
  if (colIndex < 0) return false;
  const values = [];
  for (let i = 0; i < dataRows.length; i++) {
    const v = dataRows[i][colIndex];
    if (v == null) continue;
    const s = String(v).trim();
    if (s === '') continue;
    const n = Number(s);
    if (!Number.isInteger(n) || n < 1) return false;
    values.push(n);
  }
  if (values.length === 0) return false;
  const sorted = [...new Set(values)].sort((a, b) => a - b);
  for (let i = 0; i < sorted.length; i++) {
    if (sorted[i] !== i + 1) return false;
  }
  return true;
}

/** True if column has at least one value and a majority of non-empty values have string length >= minLen. */
function columnValuesMajorityAtLeastNChars(dataRows, colIndex, minLen) {
  if (colIndex < 0 || minLen < 1) return false;
  let total = 0;
  let meeting = 0;
  for (let i = 0; i < dataRows.length; i++) {
    const v = dataRows[i][colIndex];
    if (v == null) continue;
    const s = String(v).trim();
    if (s === '') continue;
    total++;
    if (s.length >= minLen) meeting++;
  }
  return total > 0 && meeting > total / 2;
}

function cellValue(row, index) {
  if (index < 0 || index >= row.length) return '';
  const v = row[index];
  if (v == null) return '';
  return String(v).trim();
}

/** True if the string looks like a valid email address (non-empty, has @ and domain part). */
function isValidEmail(str) {
  if (str == null || typeof str !== 'string') return false;
  const s = str.trim();
  if (s === '') return false;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
}

/** Normalize for comparison: collapse all whitespace (including Unicode) to space, strip invisible chars, and trim. */
function normalizeForComparison(str) {
  if (str == null) return '';
  let s = String(str);
  if (typeof s.normalize === 'function') s = s.normalize('NFC');
  s = s.replace(/[\u200B-\u200D\uFEFF\u00AD]/g, '');
  s = s.replace(/\s/g, ' ').replace(/\u00A0/g, ' ');
  s = s.replace(/[\u2000-\u200A\u202F\u205F\u3000]/g, ' ');
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}

function cellValueRaw(row, index) {
  if (index < 0 || index >= row.length) return '';
  const v = row[index];
  if (v == null) return '';
  // Do not trim: preserve leading/trailing spaces so they are shown in the table and the relevant cell can be highlighted.
  return String(v);
}

/** Convert 0-based column index to Excel column letter(s), e.g. 0 -> A, 1 -> B, 26 -> AA. */
function colIndexToLetter(idx) {
  if (idx < 0) return '';
  let s = '';
  let n = idx;
  while (n >= 0) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

/** Get raw cell value directly from the worksheet by row (1-based) and column (0-based). Preserves leading/trailing spaces that sheet_to_json may drop. */
function getCellRawFromSheet(sheet, row1Based, colIndex0Based) {
  if (!sheet || row1Based < 1 || colIndex0Based < 0) return '';
  const addr = colIndexToLetter(colIndex0Based) + row1Based;
  const cell = sheet[addr];
  if (cell == null) return '';
  const v = cell.v;
  const w = cell.w;
  if (v != null && typeof v === 'string') return v;
  if (w != null && typeof w === 'string') return w;
  if (v != null) return String(v);
  return '';
}

/** True only when the first or last character is whitespace (so "John Paul" is allowed). */
function hasLeadingOrTrailingSpace(str) {
  const s = String(str ?? '');
  if (s.length === 0) return false;
  return /^\s/.test(s) || /\s$/.test(s);
}

function validateWorkbook(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const results = {
    fileName: '',
    sheetsProcessed: 0,
    employeeSheets: [],
    errors: [],
    warnings: [],
    summary: {
      totalRows: 0,
      validRows: 0,
      invalidRows: 0,
      duplicates: 0,
      reversedNamePairs: 0,
      sameNameDifferentId: 0,
      leadingTrailingSpaces: 0,
      firstLastNameSame: 0
    }
  };

  const employeeSheetNames = workbook.SheetNames.filter(name =>
    name.toLowerCase().includes('employees')
  );

  if (employeeSheetNames.length === 0) {
    results.warnings.push('No sheet with "Employees" in the title was found.');
    return results;
  }

  for (const sheetName of employeeSheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    if (!data || data.length < 2) {
      results.employeeSheets.push({
        name: sheetName,
        rowCount: data ? data.length : 0,
        valid: false,
        message: 'Sheet has no data or only a header row.',
        rows: []
      });
      continue;
    }

    const headers = data[0].map(h => (h != null ? String(h) : ''));
    const dataRows = data.slice(1);

    const empIdIdx = findEmployeeIdIndex(headers);
    const empNumberIdx = findEmployeeNumberIndex(headers);

    const MIN_ID_LENGTH = 4;

    let effectiveEmpIdIdx = empIdIdx;
    if (empIdIdx >= 0 && empNumberIdx >= 0) {
      const idColSequential = isColumnSequentialFromOne(dataRows, empIdIdx);
      const numberColSequential = isColumnSequentialFromOne(dataRows, empNumberIdx);
      const idColLongEnough = columnValuesMajorityAtLeastNChars(dataRows, empIdIdx, MIN_ID_LENGTH);
      const numberColLongEnough = columnValuesMajorityAtLeastNChars(dataRows, empNumberIdx, MIN_ID_LENGTH);

      if (idColSequential && !numberColSequential) {
        effectiveEmpIdIdx = empNumberIdx;
      } else if (!idColSequential && numberColSequential) {
        effectiveEmpIdIdx = empIdIdx;
      } else if (idColSequential) {
        effectiveEmpIdIdx = empNumberIdx;
      } else if (numberColSequential) {
        effectiveEmpIdIdx = empIdIdx;
      } else if (numberColLongEnough && !idColLongEnough) {
        effectiveEmpIdIdx = empNumberIdx;
      } else if (idColLongEnough && !numberColLongEnough) {
        effectiveEmpIdIdx = empIdIdx;
      } else {
        effectiveEmpIdIdx = empNumberIdx;
      }
    } else if (empNumberIdx >= 0 && empIdIdx < 0) {
      effectiveEmpIdIdx = empNumberIdx;
    }

    const firstNameIdx = findColumnIndex(
      headers,
      'First Name',
      'FirstName',
      'Given Name'
    );
    const lastNameIdx = findColumnIndex(
      headers,
      'Last Name',
      'LastName',
      'Surname',
      'Family Name'
    );
    const emailIdx = findColumnIndex(
      headers,
      'Email',
      'E-mail',
      'Email Address',
      'Email (optional)'
    );

    const sheetReport = {
      name: sheetName,
      headers: headers.filter(Boolean),
      rowCount: 0,
      valid: true,
      rows: [],
      missingFieldErrors: [],
      duplicateErrors: [],
      reversedNameErrors: [],
      sameNameDifferentIdErrors: [],
      leadingTrailingSpaceErrors: [],
      firstLastNameSameErrors: [],
      firstNameColumnIndex: firstNameIdx,
      lastNameColumnIndex: lastNameIdx,
      employeeIdentifierColumnIndex: effectiveEmpIdIdx,
      employeeIdentifierColumnLabel: effectiveEmpIdIdx >= 0
        ? (effectiveEmpIdIdx === empNumberIdx ? 'Employee Number' : 'Employee ID')
        : 'Employee ID',
      showEmailColumn: false
    };

    const seen = new Map();
    const duplicateRowIndices = new Set();
    const duplicateEmployeeNumberIndices = new Set();
    const rowsWithData = [];
    const nameToRows = new Map();
    const reversedPairsReported = new Set();

    const numCols = headers.length;
    const columnIsSequentialFromOne = [];
    for (let c = 0; c < numCols; c++) {
      columnIsSequentialFromOne[c] = isColumnSequentialFromOne(dataRows, c);
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const columnsWithData = [];
      for (let c = 0; c < numCols; c++) {
        const v = cellValue(row, c);
        if (v !== '') columnsWithData.push(c);
      }
      if (
        columnsWithData.length === 1 &&
        columnIsSequentialFromOne[columnsWithData[0]]
      ) {
        const val = cellValue(row, columnsWithData[0]);
        const n = Number(val);
        if (Number.isInteger(n) && n >= 1) continue;
      }
      if (columnsWithData.length === 1) {
        const onlyVal = (cellValue(row, columnsWithData[0]) || '').trim();
        if (onlyVal === 'AM Shift' || onlyVal === 'PM Shift') continue;
      }

      const empId = cellValue(row, effectiveEmpIdIdx);
      const firstName = cellValue(row, firstNameIdx);
      const lastName = cellValue(row, lastNameIdx);
      // Exclude rows that only have first name containing "shift" or "agency" (case-insensitive); applies to every employee sheet (all tabs).
      const firstLower = (firstName || '').trim().toLowerCase();
      if (empId === '' && lastName === '' && firstLower !== '' && (firstLower.includes('shift') || firstLower.includes('agency'))) continue;

      const email = emailIdx >= 0 ? cellValue(row, emailIdx) : '';
      const excelRowNum = i + 1;
      const empIdRaw = getCellRawFromSheet(sheet, excelRowNum, effectiveEmpIdIdx);
      const firstNameRaw = getCellRawFromSheet(sheet, excelRowNum, firstNameIdx);
      const lastNameRaw = getCellRawFromSheet(sheet, excelRowNum, lastNameIdx);
      const emailRaw = emailIdx >= 0 ? getCellRawFromSheet(sheet, excelRowNum, emailIdx) : '';

      const hasAny = empId !== '' || firstName !== '' || lastName !== '';
      if (!hasAny) continue;

      sheetReport.rowCount++;
      rowsWithData.push({
        rowIndex: i + 1,
        empId,
        firstName,
        lastName,
        email: emailIdx >= 0 ? email : undefined,
        empIdRaw,
        firstNameRaw,
        lastNameRaw,
        emailRaw: emailIdx >= 0 ? emailRaw : undefined
      });
    }

    const empIdToRowIndices = new Map();
    for (const r of rowsWithData) {
      const id = (r.empId || '').trim().toLowerCase();
      if (id === '') continue;
      if (!empIdToRowIndices.has(id)) empIdToRowIndices.set(id, []);
      empIdToRowIndices.get(id).push(r.rowIndex);
    }
    for (const indices of empIdToRowIndices.values()) {
      if (indices.length >= 2) indices.forEach(idx => duplicateEmployeeNumberIndices.add(idx));
    }

    for (const r of rowsWithData) {
      const { rowIndex, empId, firstName, lastName, email, empIdRaw, firstNameRaw, lastNameRaw, emailRaw } = r;
      const firstNorm = normalizeForComparison(firstName);
      const lastNorm = normalizeForComparison(lastName);
      const bothNamesFilled = firstName.trim() !== '' && lastName.trim() !== '';
      const firstLastSame = bothNamesFilled && (
        firstNorm.toLowerCase() === lastNorm.toLowerCase() ||
        firstName.trim().toLowerCase() === lastName.trim().toLowerCase()
      );

      const missing = [];
      if (!empId) missing.push(sheetReport.employeeIdentifierColumnLabel || 'Employee ID');
      if (!firstName) missing.push('First Name');
      if (!lastName) missing.push('Last Name');

      const key = `${empId || ''}\t${firstName || ''}\t${lastName || ''}`.toLowerCase();
      if (seen.has(key)) {
        duplicateRowIndices.add(rowIndex);
        duplicateRowIndices.add(seen.get(key));
        sheetReport.duplicateErrors.push({
          row: rowIndex,
          employeeId: empId || '(blank)',
          firstName: firstName || '(blank)',
          lastName: lastName || '(blank)',
          message: 'Duplicate employee.'
        });
      } else if (key !== '\t\t') {
        seen.set(key, rowIndex);
      }

      if (missing.length > 0) {
        sheetReport.valid = false;
        sheetReport.missingFieldErrors.push({
          row: rowIndex,
          employeeId: empId || '(blank)',
          firstName: firstName || '(blank)',
          lastName: lastName || '(blank)',
          missingFields: missing
        });
      }

      if (firstName && lastName) {
        if (firstLastSame) {
          sheetReport.valid = false;
          sheetReport.firstLastNameSameErrors.push({
            row: rowIndex,
            employeeId: empId || '(blank)',
            firstName,
            lastName,
            message: `First Name and Last Name are the same (after removing leading/trailing spaces): "${firstName}"`
          });
        }
      }

      if (firstName && lastName && !firstLastSame) {
        const nameKey = `${firstNorm.toLowerCase()}\t${lastNorm.toLowerCase()}`;
        if (!nameToRows.has(nameKey)) nameToRows.set(nameKey, []);
        nameToRows.get(nameKey).push({ rowIndex, empId });
      }

      const hasSpaces =
        hasLeadingOrTrailingSpace(empIdRaw) ||
        hasLeadingOrTrailingSpace(firstNameRaw) ||
        hasLeadingOrTrailingSpace(lastNameRaw);
      if (hasSpaces) {
        sheetReport.valid = false;
        const fields = [];
        if (hasLeadingOrTrailingSpace(empIdRaw)) fields.push(sheetReport.employeeIdentifierColumnLabel || 'Employee ID');
        if (hasLeadingOrTrailingSpace(firstNameRaw)) fields.push('First Name');
        if (hasLeadingOrTrailingSpace(lastNameRaw)) fields.push('Last Name');
        sheetReport.leadingTrailingSpaceErrors.push({
          row: rowIndex,
          employeeId: empId || '(blank)',
          firstName: firstName || '(blank)',
          lastName: lastName || '(blank)',
          fieldsWithSpaces: fields
        });
      }

      const hasConsecutiveSpacesInName = (s) => (String(s || '').trim().match(/\s{2,}/));
      const tooManySpaces = hasConsecutiveSpacesInName(firstName) || hasConsecutiveSpacesInName(lastName);
      if (tooManySpaces) sheetReport.valid = false;

      const failureReasons = [];
      if (missing.length > 0) failureReasons.push('Missing: ' + missing.join(', '));
      if (duplicateRowIndices.has(rowIndex)) failureReasons.push('Duplicate employee.');
      if (duplicateEmployeeNumberIndices.has(rowIndex)) failureReasons.push(`Duplicate ${sheetReport.employeeIdentifierColumnLabel || 'Employee Number'} for different employees`);
      if (hasSpaces) failureReasons.push('Leading or trailing spaces should be removed');
      if (tooManySpaces) failureReasons.push('Too many spaces included');
      if (firstLastSame) failureReasons.push('First name and last name are the same');
      const emailTrimmed = (email || '').trim();
      const emailIsCore = emailTrimmed.toLowerCase() === 'core';
      const emailDisplay = emailIdx >= 0 ? (emailIsCore ? '' : (emailRaw ?? '')) : '';
      const invalidEmail = emailIdx >= 0 && emailTrimmed !== '' && !emailIsCore && !isValidEmail(emailTrimmed);
      if (invalidEmail) {
        failureReasons.push('Invalid email address');
        sheetReport.valid = false;
      }

      sheetReport.rows.push({
        rowIndex,
        employeeId: empIdRaw,
        firstName: firstNameRaw,
        lastName: lastNameRaw,
        email: emailIdx >= 0 ? emailDisplay : undefined,
        isValid:
          missing.length === 0 &&
          !duplicateRowIndices.has(rowIndex) &&
          !duplicateEmployeeNumberIndices.has(rowIndex) &&
          !hasSpaces &&
          !tooManySpaces &&
          !firstLastSame &&
          !invalidEmail,
        comment: failureReasons.length > 0 ? failureReasons.join('; ') : ''
      });
    }
    sheetReport.showEmailColumn = true;

    sheetReport.rows.sort((a, b) => {
      const na = (a.employeeId != null ? String(a.employeeId) : '').trim();
      const nb = (b.employeeId != null ? String(b.employeeId) : '').trim();
      return na.localeCompare(nb, undefined, { numeric: true });
    });

    for (const r of rowsWithData) {
      if (!r.firstName || !r.lastName) continue;
      const rFirstNorm = normalizeForComparison(r.firstName);
      const rLastNorm = normalizeForComparison(r.lastName);
      if (!rFirstNorm || !rLastNorm) continue;
      const reversedKey = `${rLastNorm.toLowerCase()}\t${rFirstNorm.toLowerCase()}`;
      const others = nameToRows.get(reversedKey) || [];
      for (const o of others) {
        if (o.rowIndex === r.rowIndex) continue;
        const pairKey = [r.rowIndex, o.rowIndex].sort((a, b) => a - b).join(',');
        if (reversedPairsReported.has(pairKey)) continue;
        reversedPairsReported.add(pairKey);
        sheetReport.valid = false;
        sheetReport.reversedNameErrors.push({
          row: r.rowIndex,
          employeeId: r.empId || '(blank)',
          firstName: r.firstName,
          lastName: r.lastName,
          otherRow: o.rowIndex,
          message: `Row ${r.rowIndex} (${r.firstName} ${r.lastName}) has names reversed in row ${o.rowIndex} (${r.lastName} ${r.firstName})`
        });
      }
    }

    for (const [, entries] of nameToRows) {
      const byId = new Map();
      for (const e of entries) {
        const id = (e.empId || '').trim().toLowerCase();
        if (!byId.has(id)) byId.set(id, []);
        byId.get(id).push(e);
      }
      const distinctIds = [...byId.keys()];
      if (distinctIds.length < 2) continue;
      const firstEntry = entries[0];
      sheetReport.valid = false;
      sheetReport.sameNameDifferentIdErrors.push({
        firstName: rowsWithData.find(rd => rd.rowIndex === firstEntry.rowIndex)?.firstName || '',
        lastName: rowsWithData.find(rd => rd.rowIndex === firstEntry.rowIndex)?.lastName || '',
        rows: entries.map(e => ({ row: e.rowIndex, employeeId: e.empId || '(blank)' })),
        message: `Same First Name + Last Name associated with different Employee IDs: ${entries.map(e => `Row ${e.rowIndex} (${e.empId || 'blank'})`).join(', ')}`
      });
    }

    const dupCount = sheetReport.duplicateErrors.length;
    if (
      sheetReport.missingFieldErrors.length > 0 ||
      dupCount > 0 ||
      duplicateEmployeeNumberIndices.size > 0 ||
      sheetReport.reversedNameErrors.length > 0 ||
      sheetReport.sameNameDifferentIdErrors.length > 0 ||
      sheetReport.leadingTrailingSpaceErrors.length > 0 ||
      sheetReport.firstLastNameSameErrors.length > 0
    ) {
      sheetReport.valid = false;
    }

    const invalidForReversedOrSameId = new Set();
    for (const e of sheetReport.reversedNameErrors || []) {
      invalidForReversedOrSameId.add(e.row);
      invalidForReversedOrSameId.add(e.otherRow);
    }
    for (const e of sheetReport.sameNameDifferentIdErrors || []) {
      for (const r of e.rows) invalidForReversedOrSameId.add(r.row);
    }
    for (const e of sheetReport.firstLastNameSameErrors || []) {
      invalidForReversedOrSameId.add(e.row);
    }
    for (const r of sheetReport.rows) {
      if (invalidForReversedOrSameId.has(r.rowIndex)) {
        r.isValid = false;
        const reasons = [];
        const inReversed = (sheetReport.reversedNameErrors || []).some(
          e => e.row === r.rowIndex || e.otherRow === r.rowIndex
        );
        const inSameId = (sheetReport.sameNameDifferentIdErrors || []).some(
          e => e.rows.some(row => row.row === r.rowIndex)
        );
        if (inReversed) reasons.push('First and last name may be in wrong columns (reversed pair)');
        if (inSameId) reasons.push(`This employee has multiple ${(sheetReport.employeeIdentifierColumnLabel || 'Employee Number') === 'Employee Number' ? 'Employee Numbers' : 'Employee IDs'}`);
        if (reasons.length > 0) {
          r.comment = (r.comment ? r.comment + '; ' : '') + reasons.join('; ');
        }
      }
    }

    results.summary.totalRows += sheetReport.rowCount;
    results.summary.validRows += sheetReport.rows.filter(r => r.isValid).length;
    results.summary.invalidRows += sheetReport.rows.filter(r => !r.isValid).length;
    results.summary.duplicates += Math.ceil(dupCount / 2);
    results.summary.reversedNamePairs += sheetReport.reversedNameErrors.length;
    results.summary.sameNameDifferentId += sheetReport.sameNameDifferentIdErrors.length;
    results.summary.leadingTrailingSpaces += sheetReport.leadingTrailingSpaceErrors.length;
    results.summary.firstLastNameSame += sheetReport.firstLastNameSameErrors.length;

    results.employeeSheets.push(sheetReport);
    results.sheetsProcessed++;
  }

  return results;
}

app.post('/api/validate', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }

  try {
    const buffer = fs.readFileSync(req.file.path);
    const result = validateWorkbook(buffer);
    result.fileName = req.file.originalname;

    fs.unlink(req.file.path, () => {});

    res.json(result);
  } catch (err) {
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlink(req.file.path, () => {});
    }
    res.status(500).json({
      error: 'Failed to process file',
      message: err.message
    });
  }
});

/** Apply First/Last name column swap for given rows and return corrected Excel. */
app.post('/api/correct-export', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }
  let corrections = [];
  try {
    const raw = req.body.corrections;
    corrections = typeof raw === 'string' ? JSON.parse(raw) : raw;
    if (!Array.isArray(corrections)) {
      return res.status(400).json({ error: 'corrections must be an array' });
    }
  } catch (e) {
    return res.status(400).json({ error: 'Invalid corrections JSON' });
  }

  try {
    const buffer = fs.readFileSync(req.file.path);
    const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });

    for (const item of corrections) {
      const { sheetName, rowIndices } = item;
      if (!sheetName || !Array.isArray(rowIndices) || rowIndices.length === 0) continue;
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) continue;

      const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      if (!data || data.length < 2) continue;

      const headers = data[0].map(h => (h != null ? String(h) : ''));
      const firstNameIdx = findColumnIndex(
        headers,
        'First Name',
        'FirstName',
        'Given Name'
      );
      const lastNameIdx = findColumnIndex(
        headers,
        'Last Name',
        'LastName',
        'Surname',
        'Family Name'
      );
      if (firstNameIdx < 0 || lastNameIdx < 0) continue;

      for (const rowOneBased of rowIndices) {
        const idx = rowOneBased - 1;
        if (idx < 1 || idx >= data.length) continue;
        const row = data[idx];
        const temp = row[firstNameIdx];
        row[firstNameIdx] = row[lastNameIdx];
        row[lastNameIdx] = temp;
      }

      const newSheet = XLSX.utils.aoa_to_sheet(data);
      workbook.Sheets[sheetName] = newSheet;
    }

    const outBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    fs.unlink(req.file.path, () => {});

    const baseName = path.basename(req.file.originalname, path.extname(req.file.originalname));
    const outName = `${baseName}_corrected.xlsx`;
    res.setHeader('Content-Disposition', `attachment; filename="${outName}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(outBuffer);
  } catch (err) {
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlink(req.file.path, () => {});
    }
    res.status(500).json({
      error: 'Failed to correct and export file',
      message: err.message
    });
  }
});

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok' });
});

app.listen(PORT, () => {
  console.log(`Excel validator API running at http://localhost:${PORT}`);
});
