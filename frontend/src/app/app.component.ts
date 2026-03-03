import { Component, ChangeDetectorRef } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ExcelValidatorService } from './services/excel-validator.service';
import { revalidateSheetRows, revalidateRow } from './services/row-validation.service';
import { ValidationResult, EmployeeSheetResult, ValidationRow, TrainingRow, TrainingSheetResult, SpaceErrorType } from './models/validation-result';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './app.component.html',
  styleUrl: './app.component.css',
})
export class AppComponent {
  result: ValidationResult | null = null;
  loading = false;
  error: string | null = null;
  selectedFile: File | null = null;
  rowsToReverse: Record<string, Set<number>> = {};
  exporting = false;

  /** Sheet name -> Set of 'rowIndex-field' for cells user has confirmed as valid. */
  confirmedCells: Record<string, Set<string>> = {};

  /** Per-sheet sort: column key and direction. */
  sortState: Record<string, { key: string; dir: 'asc' | 'desc' }> = {};

  /** Sort state key for the Training table (single sheet). */
  readonly trainingSortKey = 'training';

  /** When true, Employees table shows only rows that need attention. */
  employeesShowOnlyInvalid = false;

  /** When true, Training table shows only invalid rows. */
  trainingShowOnlyInvalid = false;

  /** Top-level tab: Employees, Training, or Assets. Always visible. */
  topLevelTab: 'Employees' | 'Training' | 'Assets' = 'Employees';

  /** Active employee sub-tab: 'Employees' (Core), 'Agency Employees', or 'Instructor'. Used when topLevelTab is Employees. */
  activeTab: 'Employees' | 'Agency Employees' | 'Instructor' = 'Employees';

  /** Top-level tabs (always three). */
  readonly topLevelTabs: { id: 'Employees' | 'Training' | 'Assets'; label: string }[] = [
    { id: 'Employees', label: 'Employees' },
    { id: 'Training', label: 'Training' },
    { id: 'Assets', label: 'Assets' },
  ];

  constructor(
    private validator: ExcelValidatorService,
    private cdr: ChangeDetectorRef
  ) {}

  onFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (file) {
      const ext = file.name.toLowerCase().slice(-5);
      if (!['.xlsx', '.xls'].some(e => ext.endsWith(e))) {
        this.error = 'Please select an Excel file (.xlsx or .xls).';
        this.result = null;
        this.selectedFile = null;
        return;
      }
      this.error = null;
      this.result = null;
      this.selectedFile = file;
    }
  }

  validate(): void {
    if (!this.selectedFile) {
      this.error = 'Please select a file first.';
      return;
    }
    this.loading = true;
    this.error = null;
    this.result = null;
    this.rowsToReverse = {};
    this.confirmedCells = {};
    this.employeesShowOnlyInvalid = false;
    this.trainingShowOnlyInvalid = false;

    this.validator.validateFile(this.selectedFile).subscribe({
      next: (res) => {
        this.result = res;
        this.runClientValidationForAllSheets();
        const visible = this.getVisibleTabs();
        if (visible.length > 0) {
          const employeesSheet = this.getSheetForTab('Employees');
          const agencySheet = this.getSheetForTab('Agency Employees');
          const instructorSheet = this.getSheetForTab('Instructor');
          this.activeTab = (employeesSheet?.rows?.length ?? 0) > 0 ? 'Employees'
            : (agencySheet?.rows?.length ?? 0) > 0 ? 'Agency Employees'
            : (instructorSheet?.rows?.length ?? 0) > 0 ? 'Instructor'
            : visible[0].id;
        }
        this.loading = false;
      },
      error: (err) => {
        this.error = err.error?.message || err.message || 'Validation failed.';
        this.loading = false;
      },
    });
  }

  reset(): void {
    this.result = null;
    this.error = null;
    this.selectedFile = null;
    this.rowsToReverse = {};
    this.confirmedCells = {};
    this.topLevelTab = 'Employees';
    this.activeTab = 'Employees';
    const input = document.getElementById('file-input') as HTMLInputElement;
    if (input) input.value = '';
  }

  toggleRowToReverse(sheetName: string, rowIndex: number): void {
    if (!this.rowsToReverse[sheetName]) this.rowsToReverse[sheetName] = new Set();
    const set = this.rowsToReverse[sheetName];
    if (set.has(rowIndex)) set.delete(rowIndex);
    else set.add(rowIndex);
  }

  isRowSelectedForReverse(sheetName: string, rowIndex: number): boolean {
    return this.rowsToReverse[sheetName]?.has(rowIndex) ?? false;
  }

  getCorrectionsForExport(): { sheetName: string; rowIndices: number[] }[] {
    return Object.entries(this.rowsToReverse)
      .filter(([, set]) => set.size > 0)
      .map(([sheetName, set]) => ({ sheetName, rowIndices: Array.from(set) }));
  }

  /** Tab IDs: Instructors first (when present), then Employees, then Agency Employees. Show Instructors when sheet exists; Agency Employees always; Employees when sheet has rows. */
  get tabs(): { id: 'Employees' | 'Agency Employees' | 'Instructor'; label: string }[] {
    const order: { id: 'Employees' | 'Agency Employees' | 'Instructor'; label: string }[] = [
      { id: 'Instructor', label: 'Instructors' },
      { id: 'Employees', label: 'Employees' },
      { id: 'Agency Employees', label: 'Agency Employees' },
    ];
    return order.filter(tab => {
      const sheet = this.getSheetForTab(tab.id);
      if (tab.id === 'Agency Employees') return true;
      if (tab.id === 'Instructor') return !!sheet;
      return sheet && (sheet.rows?.length ?? 0) > 0;
    });
  }

  /** Get the sheet for the given tab. Employees → Core Employees, Agency Employees → Agency Employees, Instructor → Instructor. */
  getSheetForTab(tabId: 'Employees' | 'Agency Employees' | 'Instructor'): EmployeeSheetResult | undefined {
    const sheets = this.result?.employeeSheets ?? [];
    if (tabId === 'Employees') return sheets.find(s => s.name === 'Core Employees');
    if (tabId === 'Agency Employees') return sheets.find(s => s.name === 'Agency Employees');
    return sheets.find(s => s.name === 'Instructor');
  }

  /** Tabs that have data to display (sheet exists and has rows). Instructor tab shown whenever the sheet exists, even with 0 rows. */
  getVisibleTabs(): { id: 'Employees' | 'Agency Employees' | 'Instructor'; label: string }[] {
    return this.tabs;
  }

  setActiveTab(tabId: 'Employees' | 'Agency Employees' | 'Instructor'): void {
    this.activeTab = tabId;
  }

  setTopLevelTab(tabId: 'Employees' | 'Training' | 'Assets'): void {
    this.topLevelTab = tabId;
  }

  readonly trainingEventTypes = ['Basic', 'Refresher', 'Observation'] as const;
  readonly trainingResultOptions = ['Pass', 'Fail'] as const;

  private updateTrainingRowValidity(row: TrainingRow): void {
    const missing: string[] = [];
    if (!(row.skill ?? '').trim()) missing.push('Skill');
    if (!(row.eventType ?? '').trim()) missing.push('Event Type');
    if (!(row.testDate ?? '').trim()) missing.push('Test Date');
    if (!(row.result ?? '').trim()) missing.push('Result');
    if (!(row.employeeId ?? '').trim()) missing.push('Employee ID');
    row.missingFields = missing.length ? missing : undefined;
    row.isValid = missing.length === 0 && !row.skillError && !row.eventTypeError && !row.testDateError && !row.resultError;
    row.comment = missing.length ? 'Missing: ' + missing.join(', ') : [row.skillError, row.eventTypeError, row.testDateError, row.resultError].filter(Boolean).join('; ') || undefined;
  }

  onTrainingSkillChange(row: TrainingRow, value: string, skillOptions: string[]): void {
    row.skill = value.trim();
    const valid = skillOptions.includes(row.skill);
    row.skillError = row.skill && !valid ? 'Skill not recognised' : undefined;
    this.updateTrainingRowValidity(row);
    this.recomputeTrainingDuplicates();
    this.cdr.markForCheck();
  }

  onTrainingEventTypeChange(row: TrainingRow, value: string): void {
    row.eventType = value;
    const valid = this.trainingEventTypes.includes(value as typeof this.trainingEventTypes[number]);
    row.eventTypeError = (row.eventType && !valid) ? 'Not a valid training type' : undefined;
    this.updateTrainingRowValidity(row);
    this.cdr.markForCheck();
  }

  onTrainingTestDateChange(row: TrainingRow, value: string): void {
    const trimmed = (value ?? '').trim();
    row.testDate = trimmed;
    if (!trimmed) {
      row.testDateError = undefined;
      this.updateTrainingRowValidity(row);
      this.recomputeTrainingDuplicates();
      this.cdr.markForCheck();
      return;
    }
    const ddmmyy = trimmed.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    const ymd = trimmed.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    let valid = false;
    let d = 0, m = 0, y = 0;
    if (ddmmyy) {
      d = parseInt(ddmmyy[1], 10);
      m = parseInt(ddmmyy[2], 10) - 1;
      y = parseInt(ddmmyy[3], 10);
    } else if (ymd) {
      y = parseInt(ymd[1], 10);
      m = parseInt(ymd[2], 10) - 1;
      d = parseInt(ymd[3], 10);
    }
    if (ddmmyy || ymd) {
      const date = new Date(y, m, d);
      if (!Number.isNaN(date.getTime()) && date.getDate() === d && date.getMonth() === m && date.getFullYear() === y) {
        valid = true;
        row.testDate = `${String(d).padStart(2, '0')}/${String(m + 1).padStart(2, '0')}/${y}`;
      }
    }
    if (!valid) row.testDateError = 'Test Date must be a valid date';
    else row.testDateError = undefined;
    this.updateTrainingRowValidity(row);
    this.recomputeTrainingDuplicates();
    this.cdr.markForCheck();
  }

  onTrainingResultChange(row: TrainingRow, value: string): void {
    row.result = value;
    row.resultDefaulted = false;
    const valid = this.trainingResultOptions.includes(value as typeof this.trainingResultOptions[number]);
    row.resultError = (row.result && !valid) ? 'Result must be Pass or Fail' : undefined;
    this.updateTrainingRowValidity(row);
    this.cdr.markForCheck();
  }

  onTrainingEmployeeIdChange(row: TrainingRow, value: string): void {
    row.employeeId = value ?? '';
    this.updateTrainingRowValidity(row);
    this.recomputeTrainingDuplicates();
    this.cdr.markForCheck();
  }

  /** Recompute duplicateTraining for all training rows (Skill + Test Date + Employee ID). */
  recomputeTrainingDuplicates(): void {
    const training = this.result?.trainingSheet;
    if (!training?.rows?.length) return;
    const key = (r: TrainingRow) => `${(r.skill || '').trim().toLowerCase()}\t${(r.testDate || '').trim()}\t${(r.employeeId || '').trim().toLowerCase()}`;
    const keyToIndices = new Map<string, number[]>();
    training.rows.forEach((r, idx) => {
      const k = key(r);
      if (!keyToIndices.has(k)) keyToIndices.set(k, []);
      keyToIndices.get(k)!.push(idx);
    });
    training.rows.forEach((r, idx) => {
      const k = key(r);
      r.duplicateTraining = (keyToIndices.get(k)?.length ?? 0) > 1;
      if (r.duplicateTraining) {
        r.isValid = false;
      } else {
        this.updateTrainingRowValidity(r);
      }
    });
    training.valid = training.rows.every(r => r.isValid);
  }

  removeTrainingRow(training: TrainingSheetResult, row: TrainingRow): void {
    const idx = training.rows.indexOf(row);
    if (idx >= 0) {
      training.rows.splice(idx, 1);
      training.rowCount = training.rows.length;
      this.recomputeTrainingDuplicates();
      this.cdr.markForCheck();
    }
  }

  addTrainingSkillToFile(training: TrainingSheetResult, row: TrainingRow): void {
    const skill = (row.skill ?? '').trim();
    if (!skill) return;
    this.validator.addTrainingSkill(skill).subscribe({
      next: (res) => {
        training.skillOptions = res.skillOptions ?? [];
        // Revalidate all rows: any row whose skill is now in the list is valid
        for (const r of training.rows) {
          const s = (r.skill ?? '').trim();
          if (s && training.skillOptions.includes(s)) {
            r.skillError = undefined;
            this.updateTrainingRowValidity(r);
          }
        }
        this.recomputeTrainingDuplicates();
        this.error = null;
        this.cdr.markForCheck();
      },
      error: (err) => {
        this.error = err?.error?.message || err?.message || 'Failed to add skill';
        this.cdr.markForCheck();
      }
    });
  }

  getTrainingSkillOptions(): string[] {
    return this.result?.trainingSheet?.skillOptions ?? [];
  }

  isTrainingEventTypeInvalid(value: string): boolean {
    if (!value) return false;
    return !this.trainingEventTypes.includes(value as 'Basic' | 'Refresher' | 'Observation');
  }

  isTrainingResultInvalid(value: string): boolean {
    if (!value) return false;
    return !this.trainingResultOptions.includes(value as 'Pass' | 'Fail');
  }

  /** True when the skill value is not in training.json options (unrecognised import). */
  isTrainingSkillUnrecognised(skill: string): boolean {
    if (!skill) return false;
    return !this.getTrainingSkillOptions().includes(skill);
  }

  getTrainingSheet(): TrainingSheetResult | null | undefined {
    return this.result?.trainingSheet;
  }

  setTrainingSort(key: string): void {
    const current = this.sortState[this.trainingSortKey];
    const nextDir = current?.key === key && current?.dir === 'asc' ? 'desc' : 'asc';
    this.sortState[this.trainingSortKey] = { key, dir: nextDir };
    this.cdr.markForCheck();
  }

  getTrainingSortIcon(key: string): 'asc' | 'desc' | null {
    const s = this.sortState[this.trainingSortKey];
    if (!s || s.key !== key) return null;
    return s.dir;
  }

  getSortedTrainingRows(training: TrainingSheetResult): TrainingRow[] {
    const rows = training.rows ?? [];
    if (rows.length === 0) return rows;
    const s = this.sortState[this.trainingSortKey];
    const key = s?.key ?? 'skill';
    const dir = s?.dir ?? 'asc';
    const mult = dir === 'asc' ? 1 : -1;
    return [...rows].sort((a, b) => {
      let aVal: string | number;
      let bVal: string | number;
      if (key === 'testDate') {
        aVal = this.parseTrainingTestDate(a.testDate) ?? '';
        bVal = this.parseTrainingTestDate(b.testDate) ?? '';
        const cmp = (aVal as number) - (bVal as number);
        return cmp * mult;
      }
      aVal = String((a as unknown as Record<string, unknown>)[key] ?? '').trim().toLowerCase();
      bVal = String((b as unknown as Record<string, unknown>)[key] ?? '').trim().toLowerCase();
      const cmp = (aVal as string).localeCompare(bVal as string, undefined, { numeric: true });
      return cmp * mult;
    });
  }

  getFilteredTrainingRows(training: TrainingSheetResult): TrainingRow[] {
    const sorted = this.getSortedTrainingRows(training);
    if (!this.trainingShowOnlyInvalid) return sorted;
    return sorted.filter(row => !row.isValid);
  }

  toggleTrainingShowFilter(): void {
    this.trainingShowOnlyInvalid = !this.trainingShowOnlyInvalid;
    this.cdr.markForCheck();
  }

  /** Parse DD/MM/YYYY to timestamp for sorting; returns null if invalid. */
  private parseTrainingTestDate(s: string | undefined): number | null {
    if (!(s ?? '').trim()) return null;
    const parts = (s ?? '').trim().split(/\D/);
    if (parts.length !== 3) return null;
    const d = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10) - 1;
    const y = parseInt(parts[2], 10);
    if (Number.isNaN(d) || Number.isNaN(m) || Number.isNaN(y)) return null;
    const date = new Date(y, m, d);
    if (Number.isNaN(date.getTime()) || date.getDate() !== d || date.getMonth() !== m || date.getFullYear() !== y) return null;
    return date.getTime();
  }

  getTrainingEmployeeIdTitle(row: TrainingRow): string | null {
    const parts: string[] = [];
    if (row.duplicateTraining) parts.push('Duplicate training');
    if (!(row.employeeId ?? '').trim()) parts.push('Employee ID is missing');
    return parts.length ? parts.join('. ') : null;
  }

  /** Count of training rows that are invalid (need attention). */
  getTrainingAttentionCount(): number {
    const training = this.result?.trainingSheet;
    if (!training?.rows?.length) return 0;
    return training.rows.filter(r => !r.isValid).length;
  }

  hasReversedNameErrors(): boolean {
    return (this.result?.employeeSheets ?? []).some(sheet => (sheet.reversedNameErrors?.length ?? 0) > 0);
  }

  /** Label for the identifier column: either "Employee ID" or "Employee Number" per backend. */
  getEmployeeIdentifierColumnLabel(): string {
    const first = this.result?.employeeSheets?.find(s => (s.rows?.length ?? 0) > 0);
    return first?.employeeIdentifierColumnLabel === 'Employee Number' ? 'Employee Number' : 'Employee ID';
  }

  /** Run client-side validation on all sheets so spaceErrors/onlySpaceErrors are set for display. */
  private runClientValidationForAllSheets(): void {
    if (!this.result?.employeeSheets) return;
    for (const sheet of this.result.employeeSheets) {
      const rows = sheet.rows ?? [];
      if (rows.length === 0) continue;
      const idLabel = sheet.employeeIdentifierColumnLabel === 'Employee Number' ? 'Employee Number' : 'Employee ID';
      revalidateSheetRows(rows, idLabel);
    }
    this.updateSummary();
  }

  /** Tooltip for a cell: only when this specific cell has an error (missing, spaces) or row-level error in the responsible cell. */
  getCellTooltip(row: ValidationRow, field: 'employeeId' | 'firstName' | 'lastName' | 'email', idLabel: string): string | null {
    if (field === 'email') {
      return row.comment && row.comment.includes('Invalid email address') ? 'Invalid email address' : null;
    }
    if ((field === 'firstName' || field === 'lastName') && row.comment && row.comment.includes('Too many spaces included')) {
      const val = field === 'firstName' ? row.firstName : row.lastName;
      if (this.hasConsecutiveSpacesInName(val)) return 'Too many spaces included';
    }
    if (row.comment && row.comment.includes('Duplicate employee.')) {
      return 'Duplicate employee.';
    }
    if ((field === 'firstName' || field === 'lastName') && row.comment && row.comment.includes('for different employees')) {
      return null;
    }
    const space = row.spaceErrors?.[field];
    if (space) return this.getSpaceTooltip(space, field === 'employeeId' ? idLabel : undefined);
    const val = field === 'employeeId' ? row.employeeId : field === 'firstName' ? row.firstName : row.lastName;
    const trimmed = (val ?? '').toString().trim();
    if (trimmed === '') {
      if (field === 'employeeId') return `Missing: ${idLabel}`;
      if (field === 'firstName') return 'Missing: First Name';
      if (field === 'lastName') return 'Missing: Last Name';
    }
    if (!row.isValid && row.comment) {
      const c = row.comment.toLowerCase();
      const isDuplicateOrSameNameId = c.includes('duplicate') || (c.includes('same') && c.includes('name') && c.includes('different')) ||
        (c.includes('multiple') && c.includes('employee'));
      if (field === 'employeeId') {
        return isDuplicateOrSameNameId ? row.comment : null;
      }
      const isNameRelated = c.includes('reversed') || (c.includes('first') && c.includes('last')) || (c.includes('same') && c.includes('name'));
      if ((field === 'firstName' || field === 'lastName') && (isNameRelated || c.includes('duplicate') || (c.includes('multiple') && c.includes('employee')))) return row.comment;
    }
    return null;
  }

  /** True when this cell should show the confirm button (has error but not a space error). Do not show for empty required fields. */
  showConfirmButton(row: ValidationRow, field: 'employeeId' | 'firstName' | 'lastName' | 'email', idLabel: string): boolean {
    if (field === 'email') return this.getCellTooltip(row, 'email', idLabel) != null;
    const val = field === 'employeeId' ? row.employeeId : field === 'firstName' ? row.firstName : row.lastName;
    if ((val ?? '').toString().trim() === '') return false;
    if (row.spaceErrors?.[field]) return false;
    return this.getCellTooltip(row, field, idLabel) != null;
  }

  /** True when the string (after trim) contains two or more consecutive spaces. */
  private hasConsecutiveSpacesInName(val: string | undefined): boolean {
    return /\s{2,}/.test(String(val ?? '').trim());
  }

  getSpaceTooltip(type: SpaceErrorType, fieldLabel?: string): string {
    let msg: string;
    if (type === 'leading') msg = 'Space at start';
    else if (type === 'trailing') msg = 'Space at end';
    else msg = 'Space at start and end';
    return fieldLabel ? `${fieldLabel}: ${msg}` : msg;
  }

  /** Identifier column label for a sheet (Employee ID or Employee Number). */
  getIdLabel(sheet: { employeeIdentifierColumnLabel?: string }): string {
    return sheet.employeeIdentifierColumnLabel === 'Employee Number' ? 'Employee Number' : 'Employee ID';
  }

  /** Email to display in table: "Core" (case-insensitive) shows as empty. */
  getEmailDisplay(row: ValidationRow): string {
    const e = (row.email ?? '').toString().trim();
    return e.toLowerCase() === 'core' ? '' : (row.email ?? '');
  }

  /** Value for Employee dropdown: one of Employee, Agency Worker, Instructor, Admin; defaults by sheet when unset. */
  getEmployeeTypeValue(row: ValidationRow, sheet: { name: string }): string {
    const options = ['Employee', 'Agency Worker', 'Instructor', 'Admin'];
    const current = (row.employeeType ?? '').trim();
    if (options.includes(current)) return current;
    return sheet.name.toLowerCase().includes('agency') ? 'Agency Worker' : 'Employee';
  }

  readonly sortableKeys = ['employeeId', 'firstName', 'lastName', 'email', 'dob', 'site', 'shift', 'employeeType'] as const;

  getSortKey(sheet: { name: string }): string {
    const s = this.sortState[sheet.name];
    return s?.key ?? 'employeeId';
  }

  getSortDir(sheet: { name: string }): 'asc' | 'desc' {
    const s = this.sortState[sheet.name];
    return s?.dir ?? 'asc';
  }

  setSort(sheet: { name: string }, key: string): void {
    const current = this.sortState[sheet.name];
    const nextDir = current?.key === key && current?.dir === 'asc' ? 'desc' : 'asc';
    this.sortState[sheet.name] = { key, dir: nextDir };
    this.cdr.markForCheck();
  }

  getSortedRows(sheet: { name: string; rows?: ValidationRow[] }): ValidationRow[] {
    const rows = sheet.rows ?? [];
    if (rows.length === 0) return rows;
    const key = this.getSortKey(sheet);
    const dir = this.getSortDir(sheet);
    const mult = dir === 'asc' ? 1 : -1;
    return [...rows].sort((a, b) => {
      const aVal = String((a as unknown as Record<string, unknown>)[key] ?? '').trim().toLowerCase();
      const bVal = String((b as unknown as Record<string, unknown>)[key] ?? '').trim().toLowerCase();
      const cmp = aVal.localeCompare(bVal, undefined, { numeric: true });
      return cmp * mult;
    });
  }

  getFilteredEmployeeRows(sheet: { name: string; rows?: ValidationRow[]; employeeIdentifierColumnLabel?: string }): ValidationRow[] {
    const sorted = this.getSortedRows(sheet);
    if (!this.employeesShowOnlyInvalid) return sorted;
    const idLabel = this.getIdLabel(sheet);
    return sorted.filter(row => this.hasUnconfirmedRowErrors(row, sheet, idLabel));
  }

  toggleEmployeesShowFilter(): void {
    this.employeesShowOnlyInvalid = !this.employeesShowOnlyInvalid;
    this.cdr.markForCheck();
  }

  getSortIcon(sheet: { name: string }, key: string): 'asc' | 'desc' | null {
    const s = this.sortState[sheet.name];
    if (!s || s.key !== key) return null;
    return s.dir;
  }

  private cellKey(rowIndex: number, field: string): string {
    return `${rowIndex}-${field}`;
  }

  isCellConfirmed(sheetName: string, rowIndex: number, field: 'employeeId' | 'firstName' | 'lastName' | 'email'): boolean {
    return this.confirmedCells[sheetName]?.has(this.cellKey(rowIndex, field)) ?? false;
  }

  /** True when this row is a duplicate employee (same Employee Number + First + Last name); confirm shows as delete. */
  isDuplicateEmployeeRow(row: ValidationRow): boolean {
    return !!(row.comment && row.comment.includes('Duplicate employee.'));
  }

  /** True when on Instructor tab and this row has no Employee Number (so we show remove/confirm actions). */
  isInstructorRowWithMissingId(row: ValidationRow, sheet: { name: string }): boolean {
    return sheet.name === 'Instructor' && ((row.employeeId ?? '').toString().trim() === '');
  }

  /** Show delete button at row end (one per row) when row is a duplicate with unconfirmed errors. */
  showRowDeleteButton(row: ValidationRow, sheet: { name: string; rows?: ValidationRow[]; showEmailColumn?: boolean }, idLabel: string): boolean {
    return this.isDuplicateEmployeeRow(row) && this.hasUnconfirmedRowErrors(row, sheet, idLabel);
  }

  /** Show delete button for Instructor rows with missing Employee Number (remove row from list). */
  showInstructorMissingIdDeleteButton(row: ValidationRow, sheet: { name: string }): boolean {
    return this.isInstructorRowWithMissingId(row, sheet);
  }

  /** Show confirm button for Instructor rows with missing Employee Number (confirm as valid / keep row). */
  showInstructorMissingIdConfirmButton(row: ValidationRow, sheet: { name: string }): boolean {
    return this.isInstructorRowWithMissingId(row, sheet) && !this.isCellConfirmed(sheet.name, row.rowIndex, 'employeeId');
  }

  /** Show confirm button at row end (one per row) when row has at least one unconfirmed error that is confirmable (not missing data, not space-related). */
  showRowConfirmButton(row: ValidationRow, sheet: { name: string; rows?: ValidationRow[]; showEmailColumn?: boolean }, idLabel: string): boolean {
    if (this.isDuplicateEmployeeRow(row) || !this.hasUnconfirmedRowErrors(row, sheet, idLabel)) return false;
    const fields: ('employeeId' | 'firstName' | 'lastName' | 'email')[] = ['employeeId', 'firstName', 'lastName', 'email'];
    for (const field of fields) {
      const tip = this.getCellTooltip(row, field, idLabel);
      if (tip && !this.isCellConfirmed(sheet.name, row.rowIndex, field) && this.isConfirmableTooltip(tip)) return true;
    }
    return false;
  }

  /** True when this tooltip represents an error the user can confirm (excludes missing data and space-only issues). */
  private isConfirmableTooltip(tip: string | null): boolean {
    if (!tip) return false;
    if (tip.startsWith('Missing:')) return false;
    if (tip === 'Leading or trailing spaces should be removed') return false;
    if (tip.includes('Space at start') || tip.includes('Space at end')) return false;
    if (tip === 'Too many spaces included') return false;
    return true;
  }

  /** Confirm all unconfirmed errors in this row (called when user clicks the single row confirm button). */
  confirmRow(sheet: { name: string; rows?: ValidationRow[]; showEmailColumn?: boolean; employeeIdentifierColumnLabel?: string }, row: ValidationRow): void {
    const idLabel = sheet.employeeIdentifierColumnLabel === 'Employee Number' ? 'Employee Number' : 'Employee ID';
    const fields: ('employeeId' | 'firstName' | 'lastName' | 'email')[] = ['employeeId', 'firstName', 'lastName', 'email'];
    for (const field of fields) {
      if (this.showConfirmButton(row, field, idLabel)) this.confirmCell(sheet, row, field);
    }
    this.cdr.markForCheck();
  }

  /** Title for the confirm button (e.g. "Delete duplicate?" for duplicate employee rows). */
  getConfirmButtonTitle(row: ValidationRow): string {
    if (row.comment && row.comment.includes('Duplicate employee.')) return 'Delete duplicate?';
    return 'Click to confirm';
  }

  getInstructorMissingIdDeleteTitle(): string {
    return 'Remove row?';
  }

  getInstructorMissingIdConfirmTitle(): string {
    return 'Confirm as valid without Employee Number?';
  }

  /** Remove an Instructor row with missing Employee Number from the list. */
  removeInstructorRow(sheet: { name: string; rows?: ValidationRow[] }, row: ValidationRow): void {
    const rows = sheet.rows ?? [];
    const index = rows.findIndex(r => r.rowIndex === row.rowIndex);
    if (index >= 0) {
      rows.splice(index, 1);
      this.removeConfirmedCellsForRow(sheet.name, row.rowIndex);
      this.revalidateSheet(sheet);
      this.updateSummary();
    }
    this.cdr.markForCheck();
  }

  confirmCell(sheet: { name: string; rows?: ValidationRow[] }, row: ValidationRow, field: 'employeeId' | 'firstName' | 'lastName' | 'email'): void {
    if (field === 'email') {
      const sheetName = sheet.name;
      if (!this.confirmedCells[sheetName]) this.confirmedCells[sheetName] = new Set();
      this.confirmedCells[sheetName].add(this.cellKey(row.rowIndex, 'email'));
      this.cdr.markForCheck();
      return;
    }
    if (field === 'employeeId' && sheet.name === 'Instructor' && ((row.employeeId ?? '').toString().trim() === '')) {
      const sheetName = sheet.name;
      if (!this.confirmedCells[sheetName]) this.confirmedCells[sheetName] = new Set();
      this.confirmedCells[sheetName].add(this.cellKey(row.rowIndex, 'employeeId'));
      this.cdr.markForCheck();
      return;
    }
    if (row.comment && row.comment.includes('Duplicate employee.')) {
      const rows = sheet.rows ?? [];
      const index = rows.findIndex(r => r.rowIndex === row.rowIndex);
      if (index >= 0) {
        rows.splice(index, 1);
        this.removeConfirmedCellsForRow(sheet.name, row.rowIndex);
        this.revalidateSheet(sheet);
        this.updateSummary();
      }
      this.cdr.markForCheck();
      return;
    }
    const val = field === 'employeeId' ? row.employeeId : field === 'firstName' ? row.firstName : row.lastName;
    if ((val ?? '').toString().trim() === '') return;
    const sheetName = sheet.name;
    if (!this.confirmedCells[sheetName]) this.confirmedCells[sheetName] = new Set();
    this.confirmedCells[sheetName].add(this.cellKey(row.rowIndex, field));
    this.cdr.markForCheck();
  }

  private removeConfirmedCellsForRow(sheetName: string, rowIndex: number): void {
    const set = this.confirmedCells[sheetName];
    if (!set) return;
    for (const key of Array.from(set)) {
      if (key.startsWith(`${rowIndex}-`)) set.delete(key);
    }
  }

  /** Row has at least one validation error that is not confirmed (so row should show as invalid). */
  hasUnconfirmedRowErrors(row: ValidationRow, sheet: { name: string; rows?: ValidationRow[]; showEmailColumn?: boolean }, idLabel: string): boolean {
    if (row.isValid) return false;
    const fields: ('employeeId' | 'firstName' | 'lastName' | 'email')[] = ['employeeId', 'firstName', 'lastName', 'email'];
    for (const field of fields) {
      const tip = this.getCellTooltip(row, field, idLabel);
      if (tip && !this.isCellConfirmed(sheet.name, row.rowIndex, field)) return true;
    }
    return false;
  }

  /** Count of rows that still need attention (have unconfirmed errors). Updates when edits are made or cells are confirmed/deleted. */
  getEmployeesWhoNeedAttentionCount(): number {
    if (!this.result?.employeeSheets) return 0;
    let count = 0;
    for (const sheet of this.result.employeeSheets) {
      const rows = sheet.rows ?? [];
      const idLabel = sheet.employeeIdentifierColumnLabel === 'Employee Number' ? 'Employee Number' : 'Employee ID';
      for (const row of rows) {
        if (this.hasUnconfirmedRowErrors(row, sheet, idLabel)) count++;
      }
    }
    return count;
  }

  /** True when row had errors but every error has been confirmed (so row should show as validated). */
  allRowErrorsConfirmed(row: ValidationRow, sheet: { name: string; showEmailColumn?: boolean }, idLabel: string): boolean {
    if (row.isValid) return false;
    let hasError = false;
    let allConfirmed = true;
    const fields: ('employeeId' | 'firstName' | 'lastName' | 'email')[] = ['employeeId', 'firstName', 'lastName', 'email'];
    for (const field of fields) {
      const tip = this.getCellTooltip(row, field, idLabel);
      if (tip) {
        hasError = true;
        if (!this.isCellConfirmed(sheet.name, row.rowIndex, field)) allConfirmed = false;
      }
    }
    return hasError && allConfirmed;
  }

  /** Re-run validation for the whole sheet after a cell edit so all involved rows (duplicate emp id, same name different id) are updated. */
  onCellEdit(sheet: { name: string; rows?: { rowIndex: number; employeeId: string; firstName: string; lastName: string; isValid: boolean; comment?: string }[]; employeeIdentifierColumnLabel?: string }, _row: { rowIndex: number; employeeId: string; firstName: string; lastName: string; isValid: boolean; comment?: string }): void {
    const rows = sheet.rows ?? [];
    if (rows.length === 0) return;
    const idLabel = sheet.employeeIdentifierColumnLabel === 'Employee Number' ? 'Employee Number' : 'Employee ID';
    revalidateSheetRows(rows, idLabel);
    this.updateSummary();
    this.cdr.markForCheck();
  }

  /** Re-run validation for the entire sheet (e.g. if needed elsewhere). */
  revalidateSheet(sheet: { name: string; rows?: { rowIndex: number; employeeId: string; firstName: string; lastName: string; isValid: boolean; comment?: string }[]; employeeIdentifierColumnLabel?: string }): void {
    const rows = sheet.rows ?? [];
    if (rows.length === 0) return;
    const idLabel = sheet.employeeIdentifierColumnLabel === 'Employee Number' ? 'Employee Number' : 'Employee ID';
    revalidateSheetRows(rows, idLabel);
    this.updateSummary();
    this.cdr.markForCheck();
  }

  /** Update summary counts from current sheet data. */
  private updateSummary(): void {
    if (!this.result?.summary) return;
    let total = 0, valid = 0;
    for (const sheet of this.result.employeeSheets ?? []) {
      const rows = sheet.rows ?? [];
      for (const r of rows) {
        total++;
        if (r.isValid) valid++;
      }
    }
    this.result.summary.totalRows = total;
    this.result.summary.validRows = valid;
    this.result.summary.invalidRows = total - valid;
  }

  exportCorrected(): void {
    const corrections = this.getCorrectionsForExport();
    if (!this.selectedFile || corrections.length === 0) {
      this.error = 'Select at least one row to reverse and ensure a file was validated.';
      return;
    }
    this.exporting = true;
    this.error = null;
    this.validator.correctAndExport(this.selectedFile, corrections).subscribe({
      next: (blob) => {
        const base = this.result?.fileName?.replace(/\.(xlsx?|xls)$/i, '') || 'export';
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${base}_corrected.xlsx`;
        a.click();
        URL.revokeObjectURL(url);
        this.exporting = false;
      },
      error: (err) => {
        this.error = err.error?.message || err.message || 'Export failed.';
        this.exporting = false;
      }
    });
  }
}
