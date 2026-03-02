export interface TableRow {
  rowIndex: number;
  employeeId: string;
  firstName: string;
  lastName: string;
  email?: string;
  isValid: boolean;
  comment?: string;
  /** Set when validation runs: which fields have leading/trailing space and where. */
  spaceErrors?: { employeeId?: 'leading' | 'trailing' | 'both'; firstName?: 'leading' | 'trailing' | 'both'; lastName?: 'leading' | 'trailing' | 'both' };
  /** True when the only validation failure for this row is leading/trailing space (so only those cells are highlighted). */
  onlySpaceErrors?: boolean;
}

function isValidEmail(str: string): boolean {
  if (str == null || typeof str !== 'string') return false;
  const s = str.trim();
  if (s === '') return false;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
}

function normalize(s: string): string {
  if (s == null) return '';
  return String(s).replace(/\s+/g, ' ').replace(/\u00A0/g, ' ').trim().toLowerCase();
}

/** True only when the first character or the last character is whitespace (so "John Paul" is allowed). */
function hasLeadingOrTrailingSpace(s: string): boolean {
  const str = String(s ?? '');
  if (str.length === 0) return false;
  return /^\s/.test(str) || /\s$/.test(str);
}

/** True when the string (after trim) contains two or more consecutive spaces (e.g. "James  Thomas"). */
function hasConsecutiveSpaces(s: string): boolean {
  return /\s{2,}/.test(String(s ?? '').trim());
}

/** Returns which space error the string has: leading, trailing, or both. */
function getSpaceError(s: string): 'leading' | 'trailing' | 'both' | undefined {
  const str = String(s ?? '');
  if (str.length === 0) return undefined;
  const leading = /^\s/.test(str);
  const trailing = /\s$/.test(str);
  if (leading && trailing) return 'both';
  if (leading) return 'leading';
  if (trailing) return 'trailing';
  return undefined;
}

/** Re-run validation for all rows in a sheet and update each row's isValid and comment. */
export function revalidateSheetRows(rows: TableRow[], idLabel: string): void {
  if (!rows?.length) return;
  /** Key by array index so each row is validated independently (fixing one row does not affect others). */
  const reasonsByIndex = new Map<number, string[]>();

  function addReason(index: number, msg: string): void {
    if (!reasonsByIndex.has(index)) reasonsByIndex.set(index, []);
    const r = reasonsByIndex.get(index)!;
    if (!r.includes(msg)) r.push(msg);
  }

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const empId = (row.employeeId ?? '').trim();
    const firstName = (row.firstName ?? '').trim();
    const lastName = (row.lastName ?? '').trim();
    if (!empId) addReason(i, `Missing: ${idLabel}`);
    if (!firstName) addReason(i, 'Missing: First Name');
    if (!lastName) addReason(i, 'Missing: Last Name');
    if (hasLeadingOrTrailingSpace(row.employeeId ?? '') || hasLeadingOrTrailingSpace(row.firstName ?? '') || hasLeadingOrTrailingSpace(row.lastName ?? '')) {
      addReason(i, 'Leading or trailing spaces should be removed');
    }
    if (hasConsecutiveSpaces(row.firstName ?? '') || hasConsecutiveSpaces(row.lastName ?? '')) {
      addReason(i, 'Too many spaces included');
    }
    const firstNorm = normalize(row.firstName ?? '');
    const lastNorm = normalize(row.lastName ?? '');
    if (firstNorm && lastNorm && firstNorm === lastNorm) {
      addReason(i, 'First name and last name are the same');
    }
    const emailVal = (row as TableRow).email;
    if (emailVal != null) {
      const e = String(emailVal).trim();
      if (e !== '' && e.toLowerCase() !== 'core' && !isValidEmail(e)) addReason(i, 'Invalid email address');
    }
  }

  const seen = new Map<string, number>();
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const key = `${(row.employeeId ?? '').trim()}\t${(row.firstName ?? '').trim()}\t${(row.lastName ?? '').trim()}`.toLowerCase();
    if (key === '\t\t') continue;
    if (seen.has(key)) {
      const firstIndex = seen.get(key)!;
      addReason(i, 'Duplicate employee.');
      addReason(firstIndex, 'Duplicate employee.');
    } else {
      seen.set(key, i);
    }
  }

  const empIdToIndices = new Map<string, number[]>();
  for (let i = 0; i < rows.length; i++) {
    const id = (rows[i].employeeId ?? '').trim().toLowerCase();
    if (id === '') continue;
    if (!empIdToIndices.has(id)) empIdToIndices.set(id, []);
    empIdToIndices.get(id)!.push(i);
  }
  for (const indices of empIdToIndices.values()) {
    if (indices.length >= 2) {
      for (const i of indices) addReason(i, `Duplicate ${idLabel} for different employees`);
    }
  }

  const nameToRows = new Map<string, { index: number; empId: string }[]>();
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const firstNorm = normalize(row.firstName ?? '');
    const lastNorm = normalize(row.lastName ?? '');
    if (!firstNorm || !lastNorm) continue;
    const key = `${firstNorm}\t${lastNorm}`;
    if (!nameToRows.has(key)) nameToRows.set(key, []);
    nameToRows.get(key)!.push({ index: i, empId: (row.employeeId ?? '').trim().toLowerCase() });
  }

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const firstNorm = normalize(row.firstName ?? '');
    const lastNorm = normalize(row.lastName ?? '');
    if (!firstNorm || !lastNorm) continue;
    const reversedKey = `${lastNorm}\t${firstNorm}`;
    const others = nameToRows.get(reversedKey) ?? [];
    for (const o of others) {
      if (o.index === i) continue;
      addReason(i, 'First and last name may be in wrong columns (reversed pair)');
      addReason(o.index, 'First and last name may be in wrong columns (reversed pair)');
    }
  }

  for (const [, entries] of nameToRows) {
    const byId = new Map<string, number[]>();
    for (const e of entries) {
      const id = (e.empId ?? '').trim().toLowerCase();
      if (!byId.has(id)) byId.set(id, []);
      byId.get(id)!.push(e.index);
    }
    const distinctIds = [...byId.keys()].filter(k => k !== '');
    if (distinctIds.length >= 2) {
      for (const e of entries) addReason(e.index, `This employee has multiple ${idLabel === 'Employee Number' ? 'Employee Numbers' : 'Employee IDs'}`);
    }
  }

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const reasons = reasonsByIndex.get(i) ?? [];
    row.comment = reasons.length > 0 ? reasons.join('; ') : '';
    row.isValid = reasons.length === 0;
    const spaceOnly = reasons.length === 1 && reasons[0] === 'Leading or trailing spaces should be removed';
    row.onlySpaceErrors = !row.isValid && spaceOnly;
    row.spaceErrors = {
      employeeId: getSpaceError(row.employeeId ?? ''),
      firstName: getSpaceError(row.firstName ?? ''),
      lastName: getSpaceError(row.lastName ?? ''),
    };
  }
}

/** Re-run validation only for the row at targetIndex; only that row's isValid and comment are updated. */
export function revalidateRow(rows: TableRow[], targetIndex: number, idLabel: string): void {
  if (!rows?.length || targetIndex < 0 || targetIndex >= rows.length) return;
  const reasons: string[] = [];

  function addReasonForTarget(index: number, msg: string): void {
    if (index === targetIndex && !reasons.includes(msg)) reasons.push(msg);
  }

  const row = rows[targetIndex];
  const empId = (row.employeeId ?? '').trim();
  const firstName = (row.firstName ?? '').trim();
  const lastName = (row.lastName ?? '').trim();
  if (!empId) addReasonForTarget(targetIndex, `Missing: ${idLabel}`);
  if (!firstName) addReasonForTarget(targetIndex, 'Missing: First Name');
  if (!lastName) addReasonForTarget(targetIndex, 'Missing: Last Name');
  if (hasLeadingOrTrailingSpace(row.employeeId ?? '') || hasLeadingOrTrailingSpace(row.firstName ?? '') || hasLeadingOrTrailingSpace(row.lastName ?? '')) {
    addReasonForTarget(targetIndex, 'Leading or trailing spaces should be removed');
  }
  const firstNorm = normalize(row.firstName ?? '');
  const lastNorm = normalize(row.lastName ?? '');
  if (firstNorm && lastNorm && firstNorm === lastNorm) {
    addReasonForTarget(targetIndex, 'First name and last name are the same');
  }

  const seen = new Map<string, number>();
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const key = `${(r.employeeId ?? '').trim()}\t${(r.firstName ?? '').trim()}\t${(r.lastName ?? '').trim()}`.toLowerCase();
    if (key === '\t\t') continue;
    if (seen.has(key)) {
      const firstIndex = seen.get(key)!;
      addReasonForTarget(i, 'Duplicate employee.');
      addReasonForTarget(firstIndex, 'Duplicate employee.');
    } else {
      seen.set(key, i);
    }
  }

  const empIdToIndices = new Map<string, number[]>();
  for (let i = 0; i < rows.length; i++) {
    const id = (rows[i].employeeId ?? '').trim().toLowerCase();
    if (id === '') continue;
    if (!empIdToIndices.has(id)) empIdToIndices.set(id, []);
    empIdToIndices.get(id)!.push(i);
  }
  for (const indices of empIdToIndices.values()) {
    if (indices.length >= 2) {
      for (const i of indices) addReasonForTarget(i, `Duplicate ${idLabel} for different employees`);
    }
  }

  const nameToRows = new Map<string, { index: number; empId: string }[]>();
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const firstNorm = normalize(r.firstName ?? '');
    const lastNorm = normalize(r.lastName ?? '');
    if (!firstNorm || !lastNorm) continue;
    const key = `${firstNorm}\t${lastNorm}`;
    if (!nameToRows.has(key)) nameToRows.set(key, []);
    nameToRows.get(key)!.push({ index: i, empId: (r.employeeId ?? '').trim().toLowerCase() });
  }

  const firstNormT = normalize(row.firstName ?? '');
  const lastNormT = normalize(row.lastName ?? '');
  if (firstNormT && lastNormT) {
    const reversedKey = `${lastNormT}\t${firstNormT}`;
    const others = nameToRows.get(reversedKey) ?? [];
    for (const o of others) {
      if (o.index === targetIndex) continue;
      addReasonForTarget(targetIndex, 'First and last name may be in wrong columns (reversed pair)');
      addReasonForTarget(o.index, 'First and last name may be in wrong columns (reversed pair)');
    }
  }

  for (const [, entries] of nameToRows) {
    const byId = new Map<string, number[]>();
    for (const e of entries) {
      const id = (e.empId ?? '').trim().toLowerCase();
      if (!byId.has(id)) byId.set(id, []);
      byId.get(id)!.push(e.index);
    }
    const distinctIds = [...byId.keys()].filter(k => k !== '');
    if (distinctIds.length >= 2) {
      for (const e of entries) addReasonForTarget(e.index, `This employee has multiple ${idLabel === 'Employee Number' ? 'Employee Numbers' : 'Employee IDs'}`);
    }
  }

  const targetRow = rows[targetIndex];
  targetRow.comment = reasons.length > 0 ? reasons.join('; ') : '';
  targetRow.isValid = reasons.length === 0;
  const spaceOnly = reasons.length === 1 && reasons[0] === 'Leading or trailing spaces should be removed';
  targetRow.onlySpaceErrors = !targetRow.isValid && spaceOnly;
  targetRow.spaceErrors = {
    employeeId: getSpaceError(targetRow.employeeId ?? ''),
    firstName: getSpaceError(targetRow.firstName ?? ''),
    lastName: getSpaceError(targetRow.lastName ?? ''),
  };
}
