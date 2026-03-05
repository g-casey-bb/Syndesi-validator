const express = require('express');
const cors = require('cors');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: '20mb' }));

const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

const skillPhotosDir = path.join(__dirname, 'uploads', 'skill-photos');
if (!fs.existsSync(skillPhotosDir)) {
  fs.mkdirSync(skillPhotosDir, { recursive: true });
}

const skillPhotoIndexPath = path.join(skillPhotosDir, 'index.json');
function loadSkillPhotoIndex() {
  try {
    if (fs.existsSync(skillPhotoIndexPath)) {
      const raw = fs.readFileSync(skillPhotoIndexPath, 'utf8');
      const data = JSON.parse(raw);
      return typeof data === 'object' && data !== null ? data : {};
    }
  } catch (e) { /* ignore */ }
  return {};
}
function saveSkillPhotoIndex(index) {
  try {
    fs.writeFileSync(skillPhotoIndexPath, JSON.stringify(index, null, 2), 'utf8');
  } catch (e) { /* ignore */ }
}

/** Load skill name mapping: input skill name -> display name. Keys normalized to lowercase for lookup. Returns display options for dropdown (unique, sorted). */
function loadTrainingSkillMap() {
  const p = path.join(__dirname, '..', 'training.json');
  if (!fs.existsSync(p)) return { map: {}, displayByKey: {}, skillOptions: [] };
  const raw = JSON.parse(fs.readFileSync(p, 'utf8'));
  const displayByKey = typeof raw === 'object' && raw !== null ? raw : {};
  const map = {};
  for (const [key, display] of Object.entries(displayByKey)) {
    const k = String(key).trim().toLowerCase();
    map[k] = display == null ? key : String(display).trim();
  }
  const skillOptions = [...new Set(Object.values(map))].filter(Boolean).sort();
  return { map, displayByKey, skillOptions };
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

function generateSkillPhotoId() {
  return Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 11);
}

const skillPhotoStorage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, skillPhotosDir),
  filename: (req, file, cb) => {
    const id = generateSkillPhotoId();
    const ext = (path.extname(file.originalname) || '').toLowerCase() || '.jpg';
    cb(null, id + ext);
  }
});

const skillPhotoUpload = multer({
  storage: skillPhotoStorage,
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    if (file.mimetype && file.mimetype.startsWith('image/')) {
      cb(null, true);
    } else {
      cb(new Error('Only image files are allowed'));
    }
  }
});

function normalizeHeader(str) {
  if (str == null || typeof str !== 'string') return '';
  return String(str).trim().toLowerCase().replace(/\s+/g, ' ');
}

/** Find column index by exact header match (used when client sends column mapping). */
function findColumnIndexByMapping(headers, mappedHeaderName) {
  if (mappedHeaderName == null || String(mappedHeaderName).trim() === '') return -1;
  const n = normalizeHeader(mappedHeaderName);
  for (let i = 0; i < headers.length; i++) {
    if (normalizeHeader(headers[i]) === n) return i;
  }
  return -1;
}

function findColumnIndex(headers, ...possibleNames) {
  const normalized = possibleNames.map(normalizeHeader);
  for (let i = 0; i < headers.length; i++) {
    const h = normalizeHeader(headers[i]);
    if (normalized.some(n => h === n || h.includes(n) || n.includes(h))) return i;
  }
  return -1;
}

/** Find first column whose header starts with the given prefix (case-insensitive). */
function findColumnIndexStartsWith(headers, prefix) {
  const p = normalizeHeader(prefix);
  for (let i = 0; i < headers.length; i++) {
    if (normalizeHeader(headers[i]).startsWith(p)) return i;
  }
  return -1;
}

/** Find first column whose normalized header contains the given substring (case-insensitive). */
function findColumnIndexIncludes(headers, substring) {
  const sub = normalizeHeader(substring);
  for (let i = 0; i < headers.length; i++) {
    if (normalizeHeader(headers[i]).includes(sub)) return i;
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

/** Correct space issues for import: trim and collapse multiple consecutive spaces to one. Applied to ID, first name, last name. */
function normalizeSpaces(str) {
  if (str == null || typeof str !== 'string') return '';
  return String(str).replace(/\s+/g, ' ').replace(/\u00A0/g, ' ').trim();
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

/** True if row looks like a section title (merged cell row), e.g. "USERS WHO ARE INSTRUCTORS ONLY". Ignore for header detection. */
function isSectionTitleRow(row) {
  if (!row || !Array.isArray(row)) return true;
  const firstCell = (row[0] != null ? String(row[0]) : '').trim();
  if (firstCell.length > 20 && /users\s+who/i.test(firstCell)) return true;
  const nonEmpty = row.filter(c => (c != null ? String(c).trim() : '') !== '');
  return nonEmpty.length <= 1 && firstCell.length > 15;
}

/** True if row looks like standard instructor/employee headers (Employee ID, First Name, Last Name, Email in separate columns). Merged title rows are not header rows. Requires at least 3 distinct columns so one merged cell cannot match multiple headers. */
function isInstructorHeaderRow(row) {
  if (!row || !Array.isArray(row)) return false;
  const headers = row.map(h => (h != null ? String(h) : ''));
  const empIdIdx = findEmployeeIdIndex(headers);
  const empNumIdx = findEmployeeNumberIndex(headers);
  const firstNameIdx = findColumnIndex(headers, 'First Name', 'FirstName', 'Given Name');
  const lastNameIdx = findColumnIndex(headers, 'Last Name', 'LastName', 'Surname', 'Family Name');
  const emailIdx = findColumnIndex(headers, 'Email', 'E-mail', 'Email Address', 'Email (optional)');
  const hasId = empIdIdx >= 0 || empNumIdx >= 0;
  const hasFirst = firstNameIdx >= 0;
  const hasLast = lastNameIdx >= 0;
  const hasEmail = emailIdx >= 0;
  const count = (hasId ? 1 : 0) + (hasFirst ? 1 : 0) + (hasLast ? 1 : 0) + (hasEmail ? 1 : 0);
  const indices = [empIdIdx, empNumIdx, firstNameIdx, lastNameIdx, emailIdx].filter(i => i >= 0);
  const distinctCols = new Set(indices);
  return count >= 3 && distinctCols.size >= 3 && !isSectionTitleRow(row);
}

/**
 * Find all instructor data blocks: each block is a header row followed by data rows.
 * Merged cells / section title rows are ignored; only rows that look like real column headers are used.
 * Returns [{ headerRowIndex, headers, dataRows }, ...].
 */
function getInstructorDataBlocks(sheet, data) {
  if (!data || data.length < 2) return [];
  const blocks = [];
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    if (!row || !Array.isArray(row)) continue;
    if (isSectionTitleRow(row)) continue;
    if (!isInstructorHeaderRow(row)) continue;
    const headers = row.map(h => (h != null ? String(h) : ''));
    const empIdIdx = findEmployeeIdIndex(headers);
    const empNumIdx = findEmployeeNumberIndex(headers);
    const firstNameIdx = findColumnIndex(headers, 'First Name', 'FirstName', 'Given Name');
    const lastNameIdx = findColumnIndex(headers, 'Last Name', 'LastName', 'Surname', 'Family Name');
    const keyCols = [empIdIdx, empNumIdx, firstNameIdx, lastNameIdx].filter(i => i >= 0);
    const dataRows = [];
    for (let j = r + 1; j < data.length; j++) {
      const next = data[j];
      if (!next || !Array.isArray(next)) break;
      if (isSectionTitleRow(next)) break;
      const keyEmpty = keyCols.length === 0 || keyCols.every(c => cellValue(next, c) === '');
      if (keyEmpty) break;
      dataRows.push(next);
    }
    blocks.push({ headerRowIndex: r, headers, dataRows });
  }
  return blocks;
}

/** Build rowsWithData from one instructor block. Uses block headers and dataRows; excel row = headerRowIndex + 2 + i. */
function buildRowsWithDataFromInstructorBlock(sheet, block) {
  const { headerRowIndex, headers, dataRows } = block;
  const numCols = headers.length;
  const empIdIdx = findEmployeeIdIndex(headers);
  const empNumberIdx = findEmployeeNumberIndex(headers);
  const MIN_ID_LENGTH = 4;
  let effectiveEmpIdIdx = empIdIdx;
  if (empIdIdx >= 0 && empNumberIdx >= 0) {
    const idColSequential = isColumnSequentialFromOne(dataRows, empIdIdx);
    const numberColSequential = isColumnSequentialFromOne(dataRows, empNumberIdx);
    const idColLongEnough = columnValuesMajorityAtLeastNChars(dataRows, empIdIdx, MIN_ID_LENGTH);
    const numberColLongEnough = columnValuesMajorityAtLeastNChars(dataRows, empNumberIdx, MIN_ID_LENGTH);
    if (idColSequential && !numberColSequential) effectiveEmpIdIdx = empNumberIdx;
    else if (!idColSequential && numberColSequential) effectiveEmpIdIdx = empIdIdx;
    else if (idColSequential) effectiveEmpIdIdx = empNumberIdx;
    else if (numberColSequential) effectiveEmpIdIdx = empIdIdx;
    else if (numberColLongEnough && !idColLongEnough) effectiveEmpIdIdx = empNumberIdx;
    else if (idColLongEnough && !numberColLongEnough) effectiveEmpIdIdx = empIdIdx;
    else effectiveEmpIdIdx = empNumberIdx;
  } else if (empNumberIdx >= 0 && empIdIdx < 0) {
    effectiveEmpIdIdx = empNumberIdx;
  }
  const firstNameIdx = findColumnIndex(headers, 'First Name', 'FirstName', 'Given Name');
  const lastNameIdx = findColumnIndex(headers, 'Last Name', 'LastName', 'Surname', 'Family Name');
  const emailIdx = findColumnIndex(headers, 'Email', 'E-mail', 'Email Address', 'Email (optional)');
  const dobIdx = findColumnIndex(headers, 'DOB', 'Date of Birth', 'Birth Date', 'Date Of Birth');
  const keyIndices = [effectiveEmpIdIdx, firstNameIdx, lastNameIdx].filter(i => i >= 0);
  if (new Set(keyIndices).size < Math.min(3, keyIndices.length)) return [];
  let siteIdx = findColumnIndexStartsWith(headers, 'site id');
  if (siteIdx < 0) siteIdx = findColumnIndexStartsWith(headers, 'site');
  if (siteIdx < 0) siteIdx = findColumnIndexIncludes(headers, 'site');
  const shiftIdx = findColumnIndexStartsWith(headers, 'shift');
  let instructorYnIdx = findColumnIndex(headers, 'Instructor Y/N', 'Instructor Y/N*', 'Instructor Y/N *');
  if (instructorYnIdx < 0) instructorYnIdx = findColumnIndexStartsWith(headers, 'instructor y');
  const columnIsSequentialFromOne = [];
  for (let c = 0; c < numCols; c++) {
    columnIsSequentialFromOne[c] = isColumnSequentialFromOne(dataRows, c);
  }
  const out = [];
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    const excelRowNum = headerRowIndex + 2 + i;
    const columnsWithData = [];
    for (let c = 0; c < numCols; c++) {
      const v = cellValue(row, c);
      if (v !== '') columnsWithData.push(c);
    }
    if (columnsWithData.length === 1 && columnIsSequentialFromOne[columnsWithData[0]]) {
      const val = cellValue(row, columnsWithData[0]);
      const n = Number(val);
      if (Number.isInteger(n) && n >= 1) continue;
    }
    if (columnsWithData.length === 1) {
      const onlyVal = (cellValue(row, columnsWithData[0]) || '').trim();
      if (onlyVal === 'AM Shift' || onlyVal === 'PM Shift') continue;
    }
    const empIdRaw = getCellRawFromSheet(sheet, excelRowNum, effectiveEmpIdIdx);
    const firstNameRaw = getCellRawFromSheet(sheet, excelRowNum, firstNameIdx);
    const lastNameRaw = getCellRawFromSheet(sheet, excelRowNum, lastNameIdx);
    const empIdCorrected = normalizeSpaces(empIdRaw);
    const firstNameCorrected = normalizeSpaces(firstNameRaw);
    const lastNameCorrected = normalizeSpaces(lastNameRaw);
    const firstLower = (firstNameCorrected || '').toLowerCase();
    if (empIdCorrected === '' && lastNameCorrected === '' && firstLower !== '' && (firstLower.includes('shift') || firstLower.includes('agency'))) continue;
    const email = emailIdx >= 0 ? cellValue(row, emailIdx) : '';
    const emailRaw = emailIdx >= 0 ? getCellRawFromSheet(sheet, excelRowNum, emailIdx) : '';
    const dobRaw = dobIdx >= 0 ? getCellRawFromSheet(sheet, excelRowNum, dobIdx) : undefined;
    const siteRaw = siteIdx >= 0 ? getCellRawFromSheet(sheet, excelRowNum, siteIdx) : undefined;
    const shiftRaw = shiftIdx >= 0 ? getCellRawFromSheet(sheet, excelRowNum, shiftIdx) : undefined;
    const instructorYn = instructorYnIdx >= 0 ? getCellRawFromSheet(sheet, excelRowNum, instructorYnIdx) : '';
    const hasAny = empIdCorrected !== '' || firstNameCorrected !== '' || lastNameCorrected !== '';
    if (!hasAny) continue;
    out.push({
      rowIndex: excelRowNum,
      empId: empIdCorrected,
      firstName: firstNameCorrected,
      lastName: lastNameCorrected,
      email: emailIdx >= 0 ? email : undefined,
      empIdRaw: empIdCorrected,
      firstNameRaw: firstNameCorrected,
      lastNameRaw: lastNameCorrected,
      emailRaw: emailIdx >= 0 ? emailRaw : undefined,
      dobRaw: dobIdx >= 0 ? dobRaw : undefined,
      siteRaw: siteIdx >= 0 ? siteRaw : undefined,
      shiftRaw: shiftIdx >= 0 ? shiftRaw : undefined,
      instructorYn: instructorYnIdx >= 0 ? instructorYn : undefined
    });
  }
  return out;
}

/** Normalized key for row identity: (empId, firstName, lastName) trimmed and lowercased. */
function rowKey(row) {
  const a = (row.employeeId != null ? String(row.employeeId) : '').trim().toLowerCase();
  const b = (row.firstName != null ? String(row.firstName) : '').trim().toLowerCase();
  const c = (row.lastName != null ? String(row.lastName) : '').trim().toLowerCase();
  return `${a}\t${b}\t${c}`;
}

function validateWorkbook(buffer, options) {
  const { sheetName: singleSheetName, sheetType: singleSheetType, columnMapping: columnMappingOpt } = options || {};
  const columnMapping = columnMappingOpt && typeof columnMappingOpt === 'object' ? columnMappingOpt : {};
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

  let employeeSheetNames;
  let trainingSheetName;

  if (singleSheetName && singleSheetType && workbook.SheetNames.includes(singleSheetName)) {
    if (singleSheetType === 'training') {
      employeeSheetNames = [];
      trainingSheetName = singleSheetName;
    } else {
      employeeSheetNames = [singleSheetName];
      trainingSheetName = null;
    }
  } else {
    employeeSheetNames = workbook.SheetNames.filter(name => {
      const lower = name.toLowerCase();
      return lower.includes('employees') || lower.includes('instructor');
    });
    trainingSheetName = workbook.SheetNames.find(n => n.toLowerCase().includes('training'));
    if (employeeSheetNames.length === 0 && !singleSheetName) {
      results.warnings.push('No sheet with "Employees" or "Instructor" in the title was found.');
    }
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

    const isInstructorSheet = sheetName.toLowerCase().includes('instructor');
    const instructorBlocks = isInstructorSheet ? getInstructorDataBlocks(sheet, data) : [];
    const headers = (isInstructorSheet && instructorBlocks.length > 0)
      ? instructorBlocks[0].headers.map(h => (h != null ? String(h) : ''))
      : data[0].map(h => (h != null ? String(h) : ''));
    const dataRows = (isInstructorSheet && instructorBlocks.length > 0)
      ? instructorBlocks.flatMap(b => b.dataRows)
      : data.slice(1);

    const useMappingOnly = columnMapping && typeof columnMapping === 'object' && Object.keys(columnMapping).length > 0;

    let empIdIdx = findColumnIndexByMapping(headers, columnMapping.employeeId);
    if (!useMappingOnly && empIdIdx < 0) empIdIdx = findEmployeeIdIndex(headers);
    let empNumberIdx = useMappingOnly ? -1 : findEmployeeNumberIndex(headers);

    const MIN_ID_LENGTH = 4;

    let effectiveEmpIdIdx = empIdIdx;
    if (!useMappingOnly && empIdIdx >= 0 && empNumberIdx >= 0) {
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
    } else if (!useMappingOnly && empNumberIdx >= 0 && empIdIdx < 0) {
      effectiveEmpIdIdx = empNumberIdx;
    }

    let firstNameIdx = findColumnIndexByMapping(headers, columnMapping.firstName);
    if (!useMappingOnly && firstNameIdx < 0) firstNameIdx = findColumnIndex(headers, 'First Name', 'FirstName', 'Given Name');
    let lastNameIdx = findColumnIndexByMapping(headers, columnMapping.lastName);
    if (!useMappingOnly && lastNameIdx < 0) lastNameIdx = findColumnIndex(headers, 'Last Name', 'LastName', 'Surname', 'Family Name');
    let emailIdx = findColumnIndexByMapping(headers, columnMapping.email);
    if (!useMappingOnly && emailIdx < 0) emailIdx = findColumnIndex(headers, 'Email', 'E-mail', 'Email Address', 'Email (optional)');
    let dobIdx = findColumnIndexByMapping(headers, columnMapping.dob);
    if (!useMappingOnly && dobIdx < 0) dobIdx = findColumnIndex(headers, 'DOB', 'Date of Birth', 'Birth Date', 'Date Of Birth');
    let siteIdx = -1;
    let instructorYnIdx = -1;
    if (!useMappingOnly) {
      siteIdx = findColumnIndexStartsWith(headers, 'site id');
      if (siteIdx < 0) siteIdx = findColumnIndexStartsWith(headers, 'site');
      if (siteIdx < 0) siteIdx = findColumnIndexIncludes(headers, 'site');
      instructorYnIdx = findColumnIndex(headers, 'Instructor Y/N', 'Instructor Y/N*', 'Instructor Y/N *');
      if (instructorYnIdx < 0) instructorYnIdx = findColumnIndexStartsWith(headers, 'instructor y');
    }
    let shiftIdx = findColumnIndexByMapping(headers, columnMapping.shift);
    if (!useMappingOnly && shiftIdx < 0) shiftIdx = findColumnIndexStartsWith(headers, 'shift');

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

    if (isInstructorSheet && instructorBlocks.length > 0) {
      for (const b of instructorBlocks) {
        rowsWithData.push(...buildRowsWithDataFromInstructorBlock(sheet, b));
      }
    } else {
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
      const excelRowNum = i + 1;
      const empIdRaw = getCellRawFromSheet(sheet, excelRowNum, effectiveEmpIdIdx);
      const firstNameRaw = getCellRawFromSheet(sheet, excelRowNum, firstNameIdx);
      const lastNameRaw = getCellRawFromSheet(sheet, excelRowNum, lastNameIdx);
      const empIdCorrected = normalizeSpaces(empIdRaw);
      const firstNameCorrected = normalizeSpaces(firstNameRaw);
      const lastNameCorrected = normalizeSpaces(lastNameRaw);
      // Exclude rows that only have first name containing "shift" or "agency" (case-insensitive); applies to every employee sheet (all tabs).
      const firstLower = (firstNameCorrected || '').toLowerCase();
      if (empIdCorrected === '' && lastNameCorrected === '' && firstLower !== '' && (firstLower.includes('shift') || firstLower.includes('agency'))) continue;

      const email = emailIdx >= 0 ? cellValue(row, emailIdx) : '';
      const emailRaw = emailIdx >= 0 ? getCellRawFromSheet(sheet, excelRowNum, emailIdx) : '';

      const dobRaw = dobIdx >= 0 ? getCellRawFromSheet(sheet, excelRowNum, dobIdx) : undefined;
      const siteRaw = siteIdx >= 0 ? getCellRawFromSheet(sheet, excelRowNum, siteIdx) : undefined;
      const shiftRaw = shiftIdx >= 0 ? getCellRawFromSheet(sheet, excelRowNum, shiftIdx) : undefined;
      const instructorYn = instructorYnIdx >= 0 ? getCellRawFromSheet(sheet, excelRowNum, instructorYnIdx) : '';

      const hasAny = empIdCorrected !== '' || firstNameCorrected !== '' || lastNameCorrected !== '';
      if (!hasAny) continue;

      sheetReport.rowCount++;
      rowsWithData.push({
        rowIndex: i + 1,
        empId: empIdCorrected,
        firstName: firstNameCorrected,
        lastName: lastNameCorrected,
        email: emailIdx >= 0 ? email : undefined,
        empIdRaw: empIdCorrected,
        firstNameRaw: firstNameCorrected,
        lastNameRaw: lastNameCorrected,
        emailRaw: emailIdx >= 0 ? emailRaw : undefined,
        dobRaw: dobIdx >= 0 ? dobRaw : undefined,
        siteRaw: siteIdx >= 0 ? siteRaw : undefined,
        shiftRaw: shiftIdx >= 0 ? shiftRaw : undefined,
        instructorYn: instructorYnIdx >= 0 ? instructorYn : undefined
      });
    }
    }
    if (isInstructorSheet && instructorBlocks.length > 0) {
      sheetReport.rowCount = rowsWithData.length;
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
      const { rowIndex, empId, firstName, lastName, email, empIdRaw, firstNameRaw, lastNameRaw, emailRaw, dobRaw, siteRaw, shiftRaw, instructorYn } = r;
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

      const employeeType = sheetName.toLowerCase().includes('instructor')
        ? 'Instructor'
        : (sheetName.toLowerCase().includes('agency') ? 'Agency Worker' : 'Employee');
      const instructorYnVal = (instructorYn != null ? String(instructorYn) : '').trim().toUpperCase();
      const finalEmployeeType = (instructorYnIdx >= 0 && instructorYnVal === 'Y') ? 'Instructor' : employeeType;

      sheetReport.rows.push({
        rowIndex,
        employeeId: empIdRaw,
        firstName: firstNameRaw,
        lastName: lastNameRaw,
        email: emailIdx >= 0 ? emailDisplay : undefined,
        employeeType: finalEmployeeType,
        dob: dobRaw !== undefined ? String(dobRaw ?? '') : undefined,
        site: siteRaw !== undefined ? String(siteRaw ?? '') : undefined,
        shift: shiftRaw !== undefined ? String(shiftRaw ?? '') : undefined,
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

  // Training sheet: parse and validate
  const { map: skillMap, skillOptions } = loadTrainingSkillMap();
  const EVENT_TYPES = ['Basic', 'Refresher', 'Observation'];
  const RESULT_OPTIONS = ['Pass', 'Fail'];
  results.trainingSheet = null;
  if (trainingSheetName) {
    const trainingSheet = workbook.Sheets[trainingSheetName];
    const trainingData = XLSX.utils.sheet_to_json(trainingSheet, { header: 1, defval: '', raw: false, dateNF: 'yyyy-mm-dd' });
    if (trainingData && trainingData.length >= 2) {
      const headers = trainingData[0].map(h => (h != null ? String(h) : ''));
      const skillIdx = findColumnIndexStartsWith(headers, 'skill');
      const eventTypeIdx = findColumnIndex(headers, 'Event Type', 'EventType', 'Event type');
      const testDateIdx = findColumnIndex(headers, 'Test Date', 'TestDate', 'Test date');
      const resultIdx = findColumnIndex(headers, 'Result');
      const empNumIdxTraining = findEmployeeNumberIndex(headers);
      const empIdIdxTraining = findEmployeeIdIndex(headers);
      const empIdIdx = empNumIdxTraining >= 0 ? empNumIdxTraining : empIdIdxTraining;
      const trainingRows = [];
      let trainingValid = true;
      for (let i = 1; i < trainingData.length; i++) {
        const row = trainingData[i];
        const rowIndex = i + 1;
        const skillRaw = skillIdx >= 0 ? cellValue(row, skillIdx) : '';
        const eventTypeRaw = eventTypeIdx >= 0 ? cellValue(row, eventTypeIdx) : '';
        const testDateRaw = testDateIdx >= 0 ? (row[testDateIdx] != null ? String(row[testDateIdx]) : '') : '';
        const resultRaw = resultIdx >= 0 ? cellValue(row, resultIdx) : '';
        const employeeIdRaw = empIdIdx >= 0 ? cellValue(row, empIdIdx) : '';
        const hasAnyData = skillRaw || eventTypeRaw || testDateRaw.trim() || resultRaw || employeeIdRaw;
        if (!hasAnyData) continue;
        const missing = [];
        if (!skillRaw) missing.push('Skill');
        if (!eventTypeRaw) missing.push('Event Type');
        if (!testDateRaw.trim()) missing.push('Test Date');
        if (!employeeIdRaw) missing.push('Employee ID');
        let skillDisplay = skillRaw;
        let skillError = null;
        if (skillRaw) {
          const key = skillRaw.trim().toLowerCase();
          if (skillMap[key] !== undefined) {
            skillDisplay = skillMap[key];
          } else {
            skillError = 'Skill not recognised';
            trainingValid = false;
          }
        }
        const eventTypeNorm = eventTypeRaw.trim();
        const eventTypeConversion = eventTypeNorm.toLowerCase() === 'conversion';
        const eventTypeMatch = EVENT_TYPES.find(t => t.toLowerCase() === eventTypeNorm.toLowerCase()) || (eventTypeConversion ? 'Basic' : null);
        const eventTypeDisplay = eventTypeMatch || eventTypeNorm;
        const eventTypeError = eventTypeRaw && !eventTypeMatch ? 'Not a valid training type' : null;
        if (eventTypeError) trainingValid = false;
        const resultNorm = resultRaw.trim();
        let resultDisplay;
        let resultError = null;
        let resultDefaulted = false;
        if (!resultRaw || !resultNorm) {
          resultDisplay = 'Pass';
          resultDefaulted = true;
        } else {
          const resultMatch = RESULT_OPTIONS.find(r => r.toLowerCase() === resultNorm.toLowerCase());
          resultDisplay = resultMatch || resultNorm;
          resultError = !resultMatch ? 'Result must be Pass or Fail' : null;
          if (resultError) trainingValid = false;
        }
        let testDateValid = true;
        let testDateDisplay = testDateRaw.trim();
        if (testDateRaw.trim()) {
          const v = row[testDateIdx];
          let isoDate = '';
          if (v instanceof Date) {
            isoDate = v.toISOString().slice(0, 10);
          } else if (typeof v === 'number' && v > 0) {
            const excelEpoch = new Date(1899, 11, 30);
            const jsDate = new Date(excelEpoch.getTime() + v * 86400 * 1000);
            if (!Number.isNaN(jsDate.getTime())) isoDate = jsDate.toISOString().slice(0, 10);
            else testDateValid = false;
          } else {
            const parsed = Date.parse(testDateRaw);
            if (Number.isNaN(parsed)) testDateValid = false;
            else isoDate = new Date(parsed).toISOString().slice(0, 10);
          }
          if (isoDate) {
            const [y, m, d] = isoDate.split('-');
            testDateDisplay = `${d}/${m}/${y}`;
          }
        }
        if (!testDateValid && testDateRaw.trim()) trainingValid = false;
        const testDateError = testDateRaw.trim() && !testDateValid ? 'Test Date must be a valid date' : null;
        const rowValid = missing.length === 0 && !skillError && !eventTypeError && !testDateError && !resultError;
        if (!rowValid) trainingValid = false;
        trainingRows.push({
          rowIndex,
          skill: skillDisplay,
          skillRaw: skillRaw,
          eventType: eventTypeDisplay,
          eventTypeRaw: eventTypeRaw,
          testDate: testDateDisplay,
          testDateRaw: testDateRaw,
          result: resultDisplay,
          resultRaw: resultRaw,
          resultDefaulted: resultDefaulted,
          employeeId: employeeIdRaw,
          isValid: rowValid,
          comment: missing.length ? 'Missing: ' + missing.join(', ') : [skillError, eventTypeError, testDateError, resultError].filter(Boolean).join('; ') || undefined,
          missingFields: missing.length ? missing : undefined,
          skillError: skillError || undefined,
          eventTypeError: eventTypeError || undefined,
          testDateError: testDateError || undefined,
          resultError: resultError || undefined,
          duplicateTraining: false
        });
      }
      const trainingKey = (r) => `${(r.skill || '').trim().toLowerCase()}\t${(r.testDate || '').trim()}\t${(r.employeeId || '').trim().toLowerCase()}`;
      const keyToIndices = new Map();
      for (let idx = 0; idx < trainingRows.length; idx++) {
        const k = trainingKey(trainingRows[idx]);
        if (!keyToIndices.has(k)) keyToIndices.set(k, []);
        keyToIndices.get(k).push(idx);
      }
      for (const indices of keyToIndices.values()) {
        if (indices.length > 1) {
          for (const idx of indices) {
            trainingRows[idx].duplicateTraining = true;
            trainingRows[idx].isValid = false;
            trainingValid = false;
          }
        }
      }
      results.trainingSheet = {
        name: 'Training',
        rowCount: trainingRows.length,
        valid: trainingValid,
        rows: trainingRows,
        skillOptions
      };
    } else {
      results.trainingSheet = { name: 'Training', rowCount: 0, valid: true, rows: [], skillOptions: [] };
    }
  }

  // Merge Instructor rows with Core/Agency: same (empId, firstName, lastName) → one row with Employee = "Instructor"
  const employeeSheets = results.employeeSheets.filter(s => !s.name.toLowerCase().includes('instructor'));
  const instructorSheets = results.employeeSheets.filter(s => s.name.toLowerCase().includes('instructor'));
  if (instructorSheets.length > 0) {
    const keyToRow = new Map();
    for (const sheet of employeeSheets) {
      for (const row of sheet.rows || []) {
        const key = rowKey(row);
        keyToRow.set(key, { sheet, row });
      }
    }
    const instructorOnlyRows = [];
    for (const sheet of instructorSheets) {
      for (const row of sheet.rows || []) {
        const key = rowKey(row);
        const existing = keyToRow.get(key);
        if (existing) {
          existing.row.employeeType = 'Instructor';
          if (row.email && (existing.row.email == null || String(existing.row.email).trim() === '')) existing.row.email = row.email;
          if (row.dob && (existing.row.dob == null || String(existing.row.dob).trim() === '')) existing.row.dob = row.dob;
          if (row.site && (existing.row.site == null || String(existing.row.site).trim() === '')) existing.row.site = row.site;
          if (row.shift && (existing.row.shift == null || String(existing.row.shift).trim() === '')) existing.row.shift = row.shift;
        } else {
          instructorOnlyRows.push(row);
        }
      }
    }
    const instructorReport = {
      name: 'Instructor',
      headers: instructorSheets[0]?.headers || [],
      rowCount: instructorOnlyRows.length,
      valid: true,
      rows: instructorOnlyRows,
      missingFieldErrors: [],
      duplicateErrors: [],
      reversedNameErrors: [],
      sameNameDifferentIdErrors: [],
      leadingTrailingSpaceErrors: [],
      firstLastNameSameErrors: [],
      showEmailColumn: true,
      employeeIdentifierColumnLabel: instructorSheets[0]?.employeeIdentifierColumnLabel || 'Employee ID'
    };
    results.employeeSheets = [...employeeSheets, instructorReport];
    results.summary.totalRows = (results.employeeSheets || []).reduce((sum, s) => sum + (s.rows?.length ?? 0), 0);
    results.summary.validRows = (results.employeeSheets || []).reduce((sum, s) => sum + (s.rows?.filter(r => r.isValid).length ?? 0), 0);
    results.summary.invalidRows = results.summary.totalRows - results.summary.validRows;
  }

  return results;
}

app.post('/api/validate', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }

  try {
    const buffer = fs.readFileSync(req.file.path);
    const sheetName = typeof req.body.sheetName === 'string' && req.body.sheetName.trim() ? req.body.sheetName.trim() : null;
    const sheetType = req.body.sheetType === 'employees' || req.body.sheetType === 'training' ? req.body.sheetType : null;
    let columnMapping = null;
    if (req.body.columnMapping) {
      try {
        const raw = typeof req.body.columnMapping === 'string' ? req.body.columnMapping : JSON.stringify(req.body.columnMapping);
        columnMapping = JSON.parse(raw);
        if (typeof columnMapping !== 'object' || columnMapping === null) columnMapping = null;
      } catch (e) { columnMapping = null; }
    }
    const options = { sheetName, sheetType, columnMapping };
    const result = validateWorkbook(buffer, options);
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

/** Validate Excel from JSON body (file as base64). Use when columnMapping is required so mapping is never lost in multipart. */
app.post('/api/validate-json', (req, res) => {
  try {
    const { fileBase64, fileName, sheetName, sheetType, columnMapping } = req.body || {};
    if (!fileBase64 || typeof fileBase64 !== 'string') {
      return res.status(400).json({ error: 'Missing fileBase64' });
    }
    const buffer = Buffer.from(fileBase64, 'base64');
    if (buffer.length === 0) {
      return res.status(400).json({ error: 'Invalid fileBase64' });
    }
    const options = {
      sheetName: typeof sheetName === 'string' && sheetName.trim() ? sheetName.trim() : null,
      sheetType: sheetType === 'employees' || sheetType === 'training' ? sheetType : null,
      columnMapping: columnMapping && typeof columnMapping === 'object' ? columnMapping : {}
    };
    const result = validateWorkbook(buffer, options);
    result.fileName = typeof fileName === 'string' && fileName.trim() ? fileName.trim() : 'upload.xlsx';
    res.json(result);
  } catch (err) {
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

/** Exists so clients can verify skill-photo API is available (returns 200). Use POST /api/skill-photo/upload to upload. */
app.get('/api/skill-photo', (req, res) => {
  res.json({ status: 'ok', message: 'Use POST /api/skill-photo/upload to upload a photo' });
});

/** Upload a skill photo (image). Body: multipart file (image), folder (optional, for future OneDrive path). Returns { id } for use with download. */
app.post('/api/skill-photo/upload', skillPhotoUpload.single('file'), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }
  const folder = (req.body && req.body.folder != null) ? String(req.body.folder).trim() : '';
  const tenantId = process.env.ONEDRIVE_TENANT_ID;
  const clientId = process.env.ONEDRIVE_CLIENT_ID;
  const clientSecret = process.env.ONEDRIVE_CLIENT_SECRET;
  const driveId = process.env.ONEDRIVE_DRIVE_ID;
  const useOneDrive = tenantId && clientId && clientSecret && driveId && folder;

  if (useOneDrive) {
    try {
      const tokenRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: clientId,
          client_secret: clientSecret,
          scope: 'https://graph.microsoft.com/.default',
          grant_type: 'client_credentials'
        })
      });
      if (!tokenRes.ok) {
        const errText = await tokenRes.text();
        throw new Error('Token: ' + errText);
      }
      const tokenData = await tokenRes.json();
      const accessToken = tokenData.access_token;
      const safeName = req.file.originalname.replace(/[^a-zA-Z0-9._-]/g, '_') || 'image.jpg';
      const folderPath = folder.replace(/\\/g, '/').replace(/^\/+/, '');
      const graphPath = folderPath ? `root:/${folderPath}/${safeName}:/content` : `root:/${safeName}:/content`;
      const graphUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/${graphPath}`;
      const fileBuffer = fs.readFileSync(req.file.path);
      const uploadRes = await fetch(graphUrl, {
        method: 'PUT',
        headers: {
          'Authorization': 'Bearer ' + accessToken,
          'Content-Type': req.file.mimetype || 'application/octet-stream'
        },
        body: fileBuffer
      });
      fs.unlink(req.file.path, () => {});
      if (!uploadRes.ok) {
        const errText = await uploadRes.text();
        throw new Error('Graph: ' + errText);
      }
      const item = await uploadRes.json();
      const itemId = item && item.id;
      if (!itemId) throw new Error('No item id in response');
      return res.json({ id: itemId });
    } catch (err) {
      if (fs.existsSync(req.file.path)) fs.unlink(req.file.path, () => {});
      console.error('OneDrive upload failed, falling back to local:', err.message);
    }
  }

  const id = req.file.filename;
  const index = loadSkillPhotoIndex();
  index[id] = { path: req.file.path, originalName: req.file.originalname };
  saveSkillPhotoIndex(index);
  res.json({ id: 'local:' + id });
});

/** Download a skill photo by id. Id is either "local:filename" (local file) or a OneDrive drive item id. */
app.get('/api/skill-photo/download/:id', async (req, res) => {
  const id = (req.params.id || '').trim();
  if (!id) return res.status(400).json({ error: 'Invalid id' });

  if (id.startsWith('local:')) {
    const filename = id.slice(6).trim();
    if (!filename || filename.includes('..') || filename.includes('/') || filename.includes('\\')) {
      return res.status(400).json({ error: 'Invalid id' });
    }
    const filePath = path.join(skillPhotosDir, filename);
    if (!fs.existsSync(filePath) || !fs.statSync(filePath).isFile()) {
      return res.status(404).json({ error: 'Photo not found' });
    }
    const index = loadSkillPhotoIndex();
    const originalName = (index[filename] && index[filename].originalName) || filename;
    res.setHeader('Content-Disposition', `attachment; filename="${originalName}"`);
    const ext = path.extname(filename).toLowerCase();
    const mime = { '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png', '.gif': 'image/gif', '.webp': 'image/webp' }[ext] || 'application/octet-stream';
    res.setHeader('Content-Type', mime);
    return res.sendFile(filePath);
  }

  const tenantId = process.env.ONEDRIVE_TENANT_ID;
  const clientId = process.env.ONEDRIVE_CLIENT_ID;
  const clientSecret = process.env.ONEDRIVE_CLIENT_SECRET;
  const driveId = process.env.ONEDRIVE_DRIVE_ID;
  if (!tenantId || !clientId || !clientSecret || !driveId) {
    return res.status(404).json({ error: 'Photo not found' });
  }
  try {
    const tokenRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials'
      })
    });
    if (!tokenRes.ok) throw new Error('Token failed');
    const tokenData = await tokenRes.json();
    const graphRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${id}/content`, {
      headers: { 'Authorization': 'Bearer ' + tokenData.access_token }
    });
    if (!graphRes.ok) return res.status(404).json({ error: 'Photo not found' });
    const contentType = graphRes.headers.get('content-type') || 'application/octet-stream';
    res.setHeader('Content-Type', contentType);
    res.setHeader('Content-Disposition', 'attachment');
    const buf = await graphRes.arrayBuffer();
    res.send(Buffer.from(buf));
  } catch (err) {
    console.error('OneDrive download error:', err.message);
    res.status(500).json({ error: 'Failed to download photo' });
  }
});

/** Add a new skill to training.json so it becomes a valid selection. Body: { skill: string }. */
app.post('/api/training/skill', (req, res) => {
  const skill = req.body?.skill != null ? String(req.body.skill).trim() : '';
  if (!skill) {
    return res.status(400).json({ error: 'Skill name is required', skillOptions: loadTrainingSkillMap().skillOptions });
  }
  const p = path.join(__dirname, '..', 'training.json');
  if (!fs.existsSync(p)) {
    return res.status(500).json({ error: 'training.json not found', skillOptions: [] });
  }
  try {
    const raw = JSON.parse(fs.readFileSync(p, 'utf8'));
    const obj = typeof raw === 'object' && raw !== null ? raw : {};
    obj[skill] = skill;
    fs.writeFileSync(p, JSON.stringify(obj, null, 2), 'utf8');
    const { skillOptions } = loadTrainingSkillMap();
    res.json({ success: true, skillOptions });
  } catch (err) {
    res.status(500).json({
      error: 'Failed to add skill',
      message: err.message,
      skillOptions: loadTrainingSkillMap().skillOptions
    });
  }
});

app.listen(PORT, () => {
  console.log(`Excel validator API running at http://localhost:${PORT}`);
});
