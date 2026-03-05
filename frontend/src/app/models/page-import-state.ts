import { ValidationResult } from './validation-result';

export type PageKey = 'Employees' | 'Agency Workers' | 'Training' | 'Assets';

/** One Excel column option for the mapping dropdown (column has some data). */
export interface ExcelColumnOption {
  index: number;
  title: string;
}

/** Target columns for mapping (key used in API). */
export type ColumnMappingKey = 'employeeId' | 'firstName' | 'lastName' | 'email' | 'dob' | 'shift';

/** Training sheet column mapping keys (used when topLevelTab === 'Training'). */
export type TrainingColumnMappingKey = 'skill' | 'eventType' | 'testDate' | 'result' | 'employeeId';

/** Column mapping config: display name and whether required. */
export const COLUMN_MAPPING_COLUMNS: { key: ColumnMappingKey; label: string; required: boolean }[] = [
  { key: 'employeeId', label: 'Employee ID', required: true },
  { key: 'firstName', label: 'First Name', required: true },
  { key: 'lastName', label: 'Last Name', required: true },
  { key: 'email', label: 'Email address', required: false },
  { key: 'dob', label: 'Date of Birth', required: false },
  { key: 'shift', label: 'Shift', required: false },
];

/** Training Assign Fields columns: Skill, Event Type, Test Date, Result, Employee Number. All mandatory. */
export const COLUMN_MAPPING_COLUMNS_TRAINING: { key: TrainingColumnMappingKey; label: string; required: boolean }[] = [
  { key: 'skill', label: 'Skill', required: true },
  { key: 'eventType', label: 'Event Type', required: true },
  { key: 'testDate', label: 'Test Date', required: true },
  { key: 'result', label: 'Result', required: true },
  { key: 'employeeId', label: 'Employee Number', required: true },
];

/** Import and display state for a single page (Employees, Agency Workers, or Training). */
export interface PageImportState {
  selectedFile: File | null;
  result: ValidationResult | null;
  excelPreviewHtml: string | null;
  previewLoading: boolean;
  excelSheetNames: string[];
  selectedSheetName: string | null;
  showSheetSelectDialog: boolean;
  /** Column titles from the Excel sheet that have at least one data cell (for mapping dropdown). */
  excelColumnOptions: ExcelColumnOption[];
  /** User-selected mapping: our column key -> Excel column title (empty string = Skip). Keys from ColumnMappingKey or TrainingColumnMappingKey depending on page. */
  columnMapping: Partial<Record<ColumnMappingKey | TrainingColumnMappingKey, string>>;
  /** Current page index (0-based) in the column mapping dialog. */
  columnMappingDialogPage: number;
  showColumnMappingDialog: boolean;
  importedFileLabel: string | null;
  error: string | null;
  loading: boolean;
  employeesSubTab: 'Import' | 'Employee Data';
  agencySubTab: 'Import' | 'Agency Worker Data';
  trainingSubTab: 'Import' | 'Training Data';
  activeTab: 'Employees' | 'Agency Employees' | 'Instructor';
  employeeTabShowOnlyInvalid: boolean;
  agencyTabShowOnlyInvalid: boolean;
  /** Row indices to show when "Show only invalid" is on (fixed at toggle time so edits don't remove rows). */
  employeeTabFilterInvalidRowIndices: number[] | null;
  agencyTabFilterInvalidRowIndices: number[] | null;
  trainingShowOnlyInvalid: boolean;
  rowsToReverse: Record<string, Set<number>>;
  confirmedCells: Record<string, Set<string>>;
  nameCheckReversedProbability: Record<string, Record<number, number>>;
  nameCheckError: string | null;
}

export function createDefaultPageImportState(): PageImportState {
  return {
    selectedFile: null,
    result: null,
    excelPreviewHtml: null,
    previewLoading: false,
    excelSheetNames: [],
    selectedSheetName: null,
    showSheetSelectDialog: false,
    excelColumnOptions: [],
    columnMapping: {},
    columnMappingDialogPage: 0,
    showColumnMappingDialog: false,
    importedFileLabel: null,
    error: null,
    loading: false,
    employeesSubTab: 'Import',
    agencySubTab: 'Import',
    trainingSubTab: 'Import',
    activeTab: 'Employees',
    employeeTabShowOnlyInvalid: false,
    agencyTabShowOnlyInvalid: false,
    employeeTabFilterInvalidRowIndices: null,
    agencyTabFilterInvalidRowIndices: null,
    trainingShowOnlyInvalid: false,
    rowsToReverse: {},
    confirmedCells: {},
    nameCheckReversedProbability: {},
    nameCheckError: null,
  };
}

/** URL path for each page. */
export const PAGE_PATHS: Record<PageKey, string> = {
  'Employees': '/employees',
  'Agency Workers': '/agency-workers',
  'Training': '/training',
  'Assets': '/assets',
};

export function pathToPageKey(path: string): PageKey {
  if (path.startsWith('/agency-workers')) return 'Agency Workers';
  if (path.startsWith('/training')) return 'Training';
  if (path.startsWith('/assets')) return 'Assets';
  return 'Employees';
}
