import { ValidationResult } from './validation-result';

export type PageKey = 'Employees' | 'Agency Workers' | 'Training' | 'Assets';

/** Import and display state for a single page (Employees, Agency Workers, or Training). */
export interface PageImportState {
  selectedFile: File | null;
  result: ValidationResult | null;
  excelPreviewHtml: string | null;
  previewLoading: boolean;
  excelSheetNames: string[];
  selectedSheetName: string | null;
  showSheetSelectDialog: boolean;
  importedFileLabel: string | null;
  error: string | null;
  loading: boolean;
  employeesSubTab: 'Import' | 'Employee Data';
  agencySubTab: 'Import' | 'Agency Worker Data';
  trainingSubTab: 'Import' | 'Training Data';
  activeTab: 'Employees' | 'Agency Employees' | 'Instructor';
  employeeTabShowOnlyInvalid: boolean;
  agencyTabShowOnlyInvalid: boolean;
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
    importedFileLabel: null,
    error: null,
    loading: false,
    employeesSubTab: 'Import',
    agencySubTab: 'Import',
    trainingSubTab: 'Import',
    activeTab: 'Employees',
    employeeTabShowOnlyInvalid: false,
    agencyTabShowOnlyInvalid: false,
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
