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
  email?: string;
  /** Default "Employee" for Core Employees sheet, "Agency Worker" for Agency Employees sheet. */
  employeeType?: string;
  /** Optional; no validation. */
  dob?: string;
  /** Optional; no validation. */
  site?: string;
  /** Optional; no validation. */
  shift?: string;
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
  /** When true, the sheet has an Email column with at least one non-empty value; show it as fourth column. */
  showEmailColumn?: boolean;
}

export interface TrainingRow {
  rowIndex: number;
  skill: string;
  skillRaw: string;
  eventType: string;
  eventTypeRaw: string;
  testDate: string;
  testDateRaw: string;
  result: string;
  employeeId: string;
  dueDate?: string;
  isValid: boolean;
  comment?: string;
  missingFields?: string[];
  skillError?: string;
  eventTypeError?: string;
  testDateError?: string;
  resultError?: string;
  /** True when result was missing in the file and was defaulted to Pass. */
  resultDefaulted?: boolean;
  /** True when this row has the same Skill + Test Date + Employee ID as another row. */
  duplicateTraining?: boolean;
}

export interface TrainingSheetResult {
  name: string;
  rowCount: number;
  valid: boolean;
  rows: TrainingRow[];
  skillOptions?: string[];
}

export interface ValidationResult {
  fileName: string;
  sheetsProcessed: number;
  employeeSheets: EmployeeSheetResult[];
  trainingSheet?: TrainingSheetResult | null;
  errors: string[];
  warnings: string[];
  summary: ValidationSummary;
}
