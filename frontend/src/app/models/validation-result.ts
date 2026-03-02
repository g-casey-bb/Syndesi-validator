export interface ValidationSummary {
  totalRows: number;
  validRows: number;
  invalidRows: number;
  duplicates: number;
  reversedNamePairs?: number;
  sameNameDifferentId?: number;
  leadingTrailingSpaces?: number;
  firstLastNameSame?: number;
}

export interface MissingFieldError {
  row: number;
  employeeId: string;
  firstName: string;
  lastName: string;
  missingFields: string[];
}

export interface DuplicateError {
  row: number;
  employeeId: string;
  firstName: string;
  lastName: string;
  message: string;
}

export interface ReversedNameError {
  row: number;
  employeeId: string;
  firstName: string;
  lastName: string;
  otherRow: number;
  message: string;
}

export interface SameNameDifferentIdError {
  firstName: string;
  lastName: string;
  rows: { row: number; employeeId: string }[];
  message: string;
}

export interface LeadingTrailingSpaceError {
  row: number;
  employeeId: string;
  firstName: string;
  lastName: string;
  fieldsWithSpaces: string[];
}

export interface FirstLastNameSameError {
  row: number;
  employeeId: string;
  firstName: string;
  lastName: string;
  message: string;
}

/** Space error type for a single field (used for cell-level highlighting). */
export type SpaceErrorType = 'leading' | 'trailing' | 'both';

export interface ValidationRow {
  rowIndex: number;
  employeeId: string;
  firstName: string;
  lastName: string;
  isValid: boolean;
  comment?: string;
  spaceErrors?: { employeeId?: SpaceErrorType; firstName?: SpaceErrorType; lastName?: SpaceErrorType };
  onlySpaceErrors?: boolean;
}

export interface EmployeeSheetResult {
  name: string;
  headers: string[];
  rowCount: number;
  valid: boolean;
  message?: string;
  rows?: ValidationRow[];
  missingFieldErrors: MissingFieldError[];
  duplicateErrors: DuplicateError[];
  reversedNameErrors?: ReversedNameError[];
  sameNameDifferentIdErrors?: SameNameDifferentIdError[];
  leadingTrailingSpaceErrors?: LeadingTrailingSpaceError[];
  firstLastNameSameErrors?: FirstLastNameSameError[];
  firstNameColumnIndex?: number;
  lastNameColumnIndex?: number;
  employeeIdentifierColumnIndex?: number;
  employeeIdentifierColumnLabel?: string;
}

export interface ValidationResult {
  fileName: string;
  sheetsProcessed: number;
  employeeSheets: EmployeeSheetResult[];
  errors: string[];
  warnings: string[];
  summary: ValidationSummary;
}
