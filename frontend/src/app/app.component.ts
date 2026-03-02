import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ExcelValidatorService } from './services/excel-validator.service';
import { revalidateSheetRows, revalidateRow } from './services/row-validation.service';
import { ValidationResult, EmployeeSheetResult, ValidationRow, SpaceErrorType } from './models/validation-result';

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

  /** Active tab: 'Employees' (Core Employees) or 'Agency Employees'. Default is Employees. */
  activeTab: 'Employees' | 'Agency Employees' = 'Employees';

  constructor(private validator: ExcelValidatorService) {}

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

    this.validator.validateFile(this.selectedFile).subscribe({
      next: (res) => {
        this.result = res;
        this.runClientValidationForAllSheets();
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

  /** Tab IDs for Employees and Agency Employees. */
  readonly tabs: { id: 'Employees' | 'Agency Employees'; label: string }[] = [
    { id: 'Employees', label: 'Employees' },
    { id: 'Agency Employees', label: 'Agency Employees' },
  ];

  /** Get the sheet for the given tab. Employees → Core Employees, Agency Employees → Agency Employees. */
  getSheetForTab(tabId: 'Employees' | 'Agency Employees'): EmployeeSheetResult | undefined {
    const sheets = this.result?.employeeSheets ?? [];
    if (tabId === 'Employees') {
      return sheets.find(s => s.name === 'Core Employees') ?? sheets[0];
    }
    return sheets.find(s => s.name === 'Agency Employees') ?? sheets[1];
  }

  setActiveTab(tabId: 'Employees' | 'Agency Employees'): void {
    this.activeTab = tabId;
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
  getCellTooltip(row: ValidationRow, field: 'employeeId' | 'firstName' | 'lastName', idLabel: string): string | null {
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
  showConfirmButton(row: ValidationRow, field: 'employeeId' | 'firstName' | 'lastName', idLabel: string): boolean {
    const val = field === 'employeeId' ? row.employeeId : field === 'firstName' ? row.firstName : row.lastName;
    if ((val ?? '').toString().trim() === '') return false;
    if (row.spaceErrors?.[field]) return false;
    return this.getCellTooltip(row, field, idLabel) != null;
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

  /** Title for the confirm button (e.g. "Delete duplicate?" for duplicate employee rows). */
  getConfirmButtonTitle(row: ValidationRow): string {
    if (row.comment && row.comment.includes('Duplicate employee.')) return 'Delete duplicate?';
    return 'Click to confirm';
  }

  confirmCell(sheet: { name: string; rows?: ValidationRow[] }, row: ValidationRow, field: 'employeeId' | 'firstName' | 'lastName'): void {
    if (row.comment && row.comment.includes('Duplicate employee.')) {
      const rows = sheet.rows ?? [];
      const index = rows.findIndex(r => r.rowIndex === row.rowIndex);
      if (index >= 0) {
        rows.splice(index, 1);
        this.removeConfirmedCellsForRow(sheet.name, row.rowIndex);
        this.revalidateSheet(sheet);
        this.updateSummary();
      }
      return;
    }
    const val = field === 'employeeId' ? row.employeeId : field === 'firstName' ? row.firstName : row.lastName;
    if ((val ?? '').toString().trim() === '') return;
    const sheetName = sheet.name;
    if (!this.confirmedCells[sheetName]) this.confirmedCells[sheetName] = new Set();
    this.confirmedCells[sheetName].add(this.cellKey(row.rowIndex, field));
  }

  private removeConfirmedCellsForRow(sheetName: string, rowIndex: number): void {
    const set = this.confirmedCells[sheetName];
    if (!set) return;
    for (const key of Array.from(set)) {
      if (key.startsWith(`${rowIndex}-`)) set.delete(key);
    }
  }

  /** Row has at least one validation error that is not confirmed (so row should show as invalid). */
  hasUnconfirmedRowErrors(row: ValidationRow, sheet: { name: string; rows?: ValidationRow[] }, idLabel: string): boolean {
    if (row.isValid) return false;
    const fields: ('employeeId' | 'firstName' | 'lastName')[] = ['employeeId', 'firstName', 'lastName'];
    for (const field of fields) {
      const tip = this.getCellTooltip(row, field, idLabel);
      if (tip && !this.isCellConfirmed(sheet.name, row.rowIndex, field)) return true;
    }
    return false;
  }

  /** True when row had errors but every error has been confirmed (so row should show as validated). */
  allRowErrorsConfirmed(row: ValidationRow, sheet: { name: string }, idLabel: string): boolean {
    if (row.isValid) return false;
    let hasError = false;
    let allConfirmed = true;
    const fields: ('employeeId' | 'firstName' | 'lastName')[] = ['employeeId', 'firstName', 'lastName'];
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
  }

  /** Re-run validation for the entire sheet (e.g. if needed elsewhere). */
  revalidateSheet(sheet: { name: string; rows?: { rowIndex: number; employeeId: string; firstName: string; lastName: string; isValid: boolean; comment?: string }[]; employeeIdentifierColumnLabel?: string }): void {
    const rows = sheet.rows ?? [];
    if (rows.length === 0) return;
    const idLabel = sheet.employeeIdentifierColumnLabel === 'Employee Number' ? 'Employee Number' : 'Employee ID';
    revalidateSheetRows(rows, idLabel);
    this.updateSummary();
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
