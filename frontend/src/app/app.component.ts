import { Component, ChangeDetectorRef, OnInit, HostListener } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Router, RouterModule } from '@angular/router';
import { DomSanitizer, SafeHtml } from '@angular/platform-browser';
import { ExcelValidatorService } from './services/excel-validator.service';
import { revalidateSheetRows, revalidateRow } from './services/row-validation.service';
import { ValidationResult, EmployeeSheetResult, ValidationRow, TrainingRow, TrainingSheetResult, SpaceErrorType } from './models/validation-result';
import { PageKey, PageImportState, createDefaultPageImportState, PAGE_PATHS, pathToPageKey, COLUMN_MAPPING_COLUMNS, COLUMN_MAPPING_COLUMNS_TRAINING, ColumnMappingKey, TrainingColumnMappingKey, ExcelColumnOption } from './models/page-import-state';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, RouterModule],
  templateUrl: './app.component.html',
  styleUrl: './app.component.css',
})
export class AppComponent implements OnInit {
  /** Import and display state per page; keyed by page. */
  pageState: Record<PageKey, PageImportState> = {
    'Employees': createDefaultPageImportState(),
    'Agency Workers': createDefaultPageImportState(),
    'Training': createDefaultPageImportState(),
    'Assets': createDefaultPageImportState(),
  };

  /** Current page derived from router URL. */
  topLevelTab: PageKey = 'Employees';

  /** State for the current page (convenience accessor). */
  get currentPageState(): PageImportState {
    return this.pageState[this.topLevelTab];
  }

  exporting = false;

  /** Per-sheet sort: column key and direction (shared across pages; keyed by sheet name). */
  sortState: Record<string, { key: string; dir: 'asc' | 'desc' }> = {};

  /** Sort state key for the Training table (single sheet). */
  readonly trainingSortKey = 'training';

  /** Top-level pages (four). */
  readonly topLevelTabs: { id: PageKey; label: string }[] = [
    { id: 'Employees', label: 'Employees' },
    { id: 'Agency Workers', label: 'Agency Workers' },
    { id: 'Training', label: 'Training' },
    { id: 'Assets', label: 'Assets' },
  ];

  get employeesSubTabs(): { id: 'Import' | 'Employee Data'; label: string }[] {
    const s = this.currentPageState;
    const importLabel = (s.excelPreviewHtml && s.selectedFile?.name) ? s.selectedFile.name : (s.importedFileLabel ?? 'Import');
    return [
      { id: 'Import', label: importLabel },
      { id: 'Employee Data', label: 'Employee Data' },
    ];
  }

  get agencySubTabs(): { id: 'Import' | 'Agency Worker Data'; label: string }[] {
    const s = this.currentPageState;
    const importLabel = (s.excelPreviewHtml && s.selectedFile?.name) ? s.selectedFile.name : (s.importedFileLabel ?? 'Import');
    return [
      { id: 'Import', label: importLabel },
      { id: 'Agency Worker Data', label: 'Agency Worker Data' },
    ];
  }

  get trainingSubTabs(): { id: 'Import' | 'Training Data'; label: string }[] {
    const s = this.currentPageState;
    const importLabel = (s.excelPreviewHtml && s.selectedFile?.name) ? s.selectedFile.name : (s.importedFileLabel ?? 'Import');
    return [
      { id: 'Import', label: importLabel },
      { id: 'Training Data', label: 'Training Data' },
    ];
  }

  /** Safe HTML for iframe srcdoc (Excel preview) for current page. Cached to avoid new SafeHtml on every CD and iframe reload flicker. */
  private _cachedPreviewHtml: string | null = null;
  private _cachedSanitizedPreview: SafeHtml | null = null;
  get sanitizedExcelPreview(): SafeHtml | null {
    const html = this.currentPageState.excelPreviewHtml;
    if (html === this._cachedPreviewHtml) return this._cachedSanitizedPreview;
    this._cachedPreviewHtml = html;
    this._cachedSanitizedPreview = html ? this.sanitizer.bypassSecurityTrustHtml(html) : null;
    return this._cachedSanitizedPreview;
  }

  /** Column widths (px) measured from the preview table; used so mapping row and overlay match iframe columns. */
  previewColumnWidths: number[] = [];

  /** Sum of previewColumnWidths for min-width. */
  get previewTableTotalWidth(): number {
    return this.previewColumnWidths.length ? this.previewColumnWidths.reduce((a, b) => a + b, 0) : 0;
  }

  /** CSS grid-template-columns from measured widths, or fallback to equal 1fr. */
  get previewMappingGridColumns(): string {
    const n = this.previewMappingCells.length;
    if (n && this.previewColumnWidths.length === n) {
      return this.previewColumnWidths.map(w => w + 'px').join(' ');
    }
    return n ? `repeat(${n}, 1fr)` : '1fr';
  }

  /** Called when the preview iframe has loaded; measure table column widths. */
  onPreviewIframeLoad(ev: Event): void {
    const iframe = ev?.target as HTMLIFrameElement | null;
    if (!iframe?.contentDocument) return;
    const table = iframe.contentDocument.getElementById('excel-preview-table') as HTMLTableElement | null;
    if (!table?.rows?.length) return;
    const firstRow = table.rows[0];
    const widths: number[] = [];
    for (let i = 0; i < firstRow.cells.length; i++) {
      widths.push((firstRow.cells[i] as HTMLTableCellElement).offsetWidth);
    }
    if (widths.length > 0) {
      this.previewColumnWidths = widths;
      this.cdr.markForCheck();
    }
  }

  /** Path for a page (for routerLink). */
  getPathForPage(page: PageKey): string {
    return PAGE_PATHS[page];
  }

  /** Whether the Settings dialog is open. */
  settingsDialogOpen = false;

  /** Whether the Maintenance page is visible. */
  maintenancePageOpen = false;

  /** Active tab on Maintenance page. */
  maintenanceTab: 'Skills' | 'Testing' = 'Skills';

  /** Skills list from Maintenance upload, persisted in localStorage. */
  skillsList: string[] = [];

  /** Testing tab: Customer, Site, User (persisted in localStorage). */
  testingCustomer = '';
  testingSite = '';
  testingUser = '';

  /** Whether the skills list section is expanded. */
  skillsListExpanded = false;

  private readonly SKILLS_STORAGE_KEY = 'syndesi_skills';
  private readonly SKILL_QUERIES_STORAGE_KEY = 'syndesi_skill_queries';
  private readonly NEXT_SKILL_QUERY_ID_KEY = 'syndesi_skill_query_next_id';
  private readonly SKILL_QUERIES_ADDED_IDS_KEY = 'syndesi_skill_queries_added_ids';
  private readonly TESTING_CUSTOMER_KEY = 'syndesi_testing_customer';
  private readonly TESTING_SITE_KEY = 'syndesi_testing_site';
  private readonly TESTING_USER_KEY = 'syndesi_testing_user';

  /** Unknown Skill dialog: when + is clicked for unrecognised skill. */
  unknownSkillDialogOpen = false;
  unknownSkillDialogSkill = '';
  unknownSkillForm: { assetId: string; make: string; model: string; attachment: string; photoUrl: string; imageData: string | null } = {
    assetId: '',
    make: '',
    model: '',
    attachment: '',
    photoUrl: '',
    imageData: null,
  };

  /** Sent skill queries (stored in localStorage). */
  skillQueries: { id: number; customer: string; site: string; user: string; assetId: string; skillName: string; make: string; model: string; attachment: string; photoUrl: string; imageData?: string; imageDriveItemId?: string }[] = [];
  private nextSkillQueryId = 1;

  /** Query being reviewed in the Review Skill dialog (null when closed). */
  reviewSkillQuery: { id: number; skillName: string; make: string; model: string; assetId: string; attachment: string; photoUrl: string; imageDriveItemId?: string; imageData?: string } | null = null;

  /** Editable form in Review Skill dialog; saved on Add, discarded on Close. */
  reviewSkillForm: { skillName: string; make: string; model: string; assetId: string; attachment: string; photoUrl: string } = {
    skillName: '', make: '', model: '', assetId: '', attachment: '', photoUrl: '',
  };

  /** IDs of skill queries that have been added to the Skills list (rows are hidden, not deleted). */
  skillQueriesAddedIds = new Set<number>();

  /** ChatGPT API key (in-memory; persisted to localStorage on Settings OK). */
  chatGptApiKey = '';

  /** Temporary value for the API key while the Settings dialog is open (Cancel discards). */
  settingsApiKeyInput = '';

  /** Skill photos folder (OneDrive path or folder id), persisted in Settings. */
  skillPhotosFolder = '';

  /** Temporary value for Skill photos folder while Settings dialog is open. */
  settingsSkillPhotosFolderInput = '';

  /** Probability (0–1) that first/last names are reversed, by sheet name then rowIndex. Set by "Check names". */
  /** Now stored per-page in pageState[].nameCheckReversedProbability */

  nameCheckLoading = false;

  /** Set while uploading skill photo in Unknown Skill dialog. */
  unknownSkillPhotoUploading = false;
  unknownSkillPhotoUploadError: string | null = null;

  /** True when "Finished checking names" dialog is open. */
  nameCheckCompleteDialogOpen = false;

  /** Whether we have any name-check results to show (for dialog). */
  get hasNameCheckResults(): boolean {
    return Object.keys(this.currentPageState.nameCheckReversedProbability).some(
      sheetName => Object.keys(this.currentPageState.nameCheckReversedProbability[sheetName] ?? {}).length > 0
    );
  }

  private readonly CHATGPT_API_KEY_STORAGE = 'syndesi_chatgpt_api_key';
  private readonly SKILL_PHOTOS_FOLDER_STORAGE = 'syndesi_skill_photos_folder';
  private readonly OPENAI_CHAT_URL = 'https://api.openai.com/v1/chat/completions';

  constructor(
    private validator: ExcelValidatorService,
    private cdr: ChangeDetectorRef,
    private http: HttpClient,
    private sanitizer: DomSanitizer,
    private router: Router
  ) {}

  ngOnInit(): void {
    this.syncPageFromRouter();
    this.router.events.subscribe(() => this.syncPageFromRouter());
    try {
      this.chatGptApiKey = localStorage.getItem(this.CHATGPT_API_KEY_STORAGE) ?? '';
      this.skillPhotosFolder = localStorage.getItem(this.SKILL_PHOTOS_FOLDER_STORAGE) ?? '';
      this.testingCustomer = localStorage.getItem(this.TESTING_CUSTOMER_KEY) ?? '';
      this.testingSite = localStorage.getItem(this.TESTING_SITE_KEY) ?? '';
      this.testingUser = localStorage.getItem(this.TESTING_USER_KEY) ?? '';
    } catch {
      this.chatGptApiKey = '';
      this.skillPhotosFolder = '';
      this.testingCustomer = '';
      this.testingSite = '';
      this.testingUser = '';
    }
    this.loadSkillsFromStorage();
    this.loadSkillQueriesFromStorage();
  }

  private syncPageFromRouter(): void {
    const path = this.router.url.split('?')[0];
    const page = pathToPageKey(path);
    if (page !== this.topLevelTab) {
      this.topLevelTab = page;
      this.cdr.markForCheck();
    }
  }

  navigateToPage(page: PageKey): void {
    this.closeMaintenance();
    this.router.navigate([PAGE_PATHS[page].replace(/^\//, '')]);
  }

  openMaintenance(): void {
    this.maintenancePageOpen = true;
    this.loadSkillsFromStorage();
    this.loadSkillQueriesFromStorage();
    this.loadTestingFromStorage();
    this.cdr.markForCheck();
  }

  loadTestingFromStorage(): void {
    try {
      this.testingCustomer = localStorage.getItem(this.TESTING_CUSTOMER_KEY) ?? '';
      this.testingSite = localStorage.getItem(this.TESTING_SITE_KEY) ?? '';
      this.testingUser = localStorage.getItem(this.TESTING_USER_KEY) ?? '';
    } catch {
      this.testingCustomer = '';
      this.testingSite = '';
      this.testingUser = '';
    }
  }

  setTestingCustomer(value: string): void {
    this.testingCustomer = value;
    try {
      localStorage.setItem(this.TESTING_CUSTOMER_KEY, value);
    } catch { /* ignore */ }
    this.cdr.markForCheck();
  }

  setTestingSite(value: string): void {
    this.testingSite = value;
    try {
      localStorage.setItem(this.TESTING_SITE_KEY, value);
    } catch { /* ignore */ }
    this.cdr.markForCheck();
  }

  setTestingUser(value: string): void {
    this.testingUser = value;
    try {
      localStorage.setItem(this.TESTING_USER_KEY, value);
    } catch { /* ignore */ }
    this.cdr.markForCheck();
  }

  closeMaintenance(): void {
    this.maintenancePageOpen = false;
    this.cdr.markForCheck();
  }

  loadSkillsFromStorage(): void {
    try {
      const raw = localStorage.getItem(this.SKILLS_STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw) as unknown;
        this.skillsList = Array.isArray(parsed)
          ? parsed.filter((x): x is string => typeof x === 'string').map(s => String(s).trim()).filter(Boolean)
          : [];
      } else {
        this.skillsList = [];
      }
    } catch {
      this.skillsList = [];
    }
  }

  saveSkillsToStorage(skills: string[]): void {
    try {
      localStorage.setItem(this.SKILLS_STORAGE_KEY, JSON.stringify(skills));
    } catch {
      /* ignore */
    }
  }

  onSkillsFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input?.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const content = reader.result as string;
      const skills = this.parseSkillsJson(content);
      this.skillsList = skills;
      this.saveSkillsToStorage(skills);
      this.cdr.markForCheck();
    };
    reader.readAsText(file);
    input.value = '';
  }

  /** Parse JSON file: array of strings, or array of objects with name/skill. */
  private parseSkillsJson(content: string): string[] {
    try {
      const parsed = JSON.parse(content) as unknown;
      if (!Array.isArray(parsed)) return [];
      return parsed
        .map(item => {
          if (typeof item === 'string') return item.trim();
          if (item && typeof item === 'object') {
            const o = item as Record<string, unknown>;
            const s = (o['name'] ?? o['skill'] ?? o['title']) as string | undefined;
            return s != null ? String(s).trim() : '';
          }
          return '';
        })
        .filter(Boolean);
    } catch {
      return [];
    }
  }

  clearSkills(): void {
    this.skillsList = [];
    try {
      localStorage.removeItem(this.SKILLS_STORAGE_KEY);
    } catch {
      /* ignore */
    }
    this.cdr.markForCheck();
  }

  toggleSkillsListExpanded(): void {
    this.skillsListExpanded = !this.skillsListExpanded;
    this.cdr.markForCheck();
  }

  loadSkillQueriesFromStorage(): void {
    try {
      const raw = localStorage.getItem(this.SKILL_QUERIES_STORAGE_KEY);
      const parsed = raw ? (JSON.parse(raw) as unknown) : [];
      this.skillQueries = Array.isArray(parsed)
        ? parsed.map((q: Record<string, unknown>) => ({
            id: Number(q['id']),
            customer: typeof q['customer'] === 'string' ? q['customer'] : '',
            site: typeof q['site'] === 'string' ? q['site'] : '',
            user: typeof q['user'] === 'string' ? q['user'] : '',
            assetId: typeof q['assetId'] === 'string' ? q['assetId'] : '',
            skillName: String(q['skillName'] ?? ''),
            make: String(q['make'] ?? ''),
            model: String(q['model'] ?? ''),
            attachment: String(q['attachment'] ?? ''),
            photoUrl: String(q['photoUrl'] ?? ''),
            ...(q['imageData'] ? { imageData: q['imageData'] as string } : {}),
            ...(typeof q['imageDriveItemId'] === 'string' ? { imageDriveItemId: q['imageDriveItemId'] } : {}),
          }))
        : [];
      const nextRaw = localStorage.getItem(this.NEXT_SKILL_QUERY_ID_KEY);
      this.nextSkillQueryId = nextRaw ? Math.max(1, parseInt(nextRaw, 10)) : 1;
      const addedRaw = localStorage.getItem(this.SKILL_QUERIES_ADDED_IDS_KEY);
      const addedArr = addedRaw ? (JSON.parse(addedRaw) as unknown) : [];
      this.skillQueriesAddedIds = new Set(Array.isArray(addedArr) ? addedArr.map((n: unknown) => Number(n)).filter((n: number) => !Number.isNaN(n)) : []);
    } catch {
      this.skillQueries = [];
      this.nextSkillQueryId = 1;
      this.skillQueriesAddedIds = new Set();
    }
  }

  private saveSkillQueriesToStorage(): void {
    try {
      localStorage.setItem(this.SKILL_QUERIES_STORAGE_KEY, JSON.stringify(this.skillQueries));
      localStorage.setItem(this.NEXT_SKILL_QUERY_ID_KEY, String(this.nextSkillQueryId));
    } catch {
      /* ignore */
    }
  }

  private saveSkillQueriesAddedIdsToStorage(): void {
    try {
      localStorage.setItem(this.SKILL_QUERIES_ADDED_IDS_KEY, JSON.stringify([...this.skillQueriesAddedIds]));
    } catch {
      /* ignore */
    }
  }

  wasSkillQueryAddedToSkills(id: number): boolean {
    return this.skillQueriesAddedIds.has(id);
  }

  clearAllSkillQueries(): void {
    this.skillQueries = [];
    this.nextSkillQueryId = 1;
    this.skillQueriesAddedIds = new Set();
    this.saveSkillQueriesToStorage();
    this.saveSkillQueriesAddedIdsToStorage();
    this.cdr.markForCheck();
  }

  updateSkillQueryField(id: number, field: 'customer' | 'site' | 'user' | 'assetId' | 'skillName' | 'make' | 'model' | 'attachment' | 'photoUrl', value: string): void {
    const q = this.skillQueries.find(x => x.id === id);
    if (q) {
      q[field] = value;
      this.saveSkillQueriesToStorage();
      this.cdr.markForCheck();
    }
  }

  addSkillQueryToSkillsList(id: number, skillName: string): void {
    const name = (skillName ?? '').trim();
    if (!name) return;
    if (!this.skillsList.includes(name)) {
      this.skillsList = [...this.skillsList, name];
      this.saveSkillsToStorage(this.skillsList);
    }
    this.skillQueriesAddedIds.add(id);
    this.saveSkillQueriesAddedIdsToStorage();
    this.cdr.markForCheck();
  }

  /** Whether the skill query has a photo (uploaded or legacy inline). */
  hasSkillQueryPhoto(q: { imageDriveItemId?: string; imageData?: string }): boolean {
    return !!(q.imageDriveItemId || (q.imageData && (q.imageData as string).startsWith('data:')));
  }

  /** URL to open/download the skill query photo (backend proxy or data URL). */
  getSkillQueryPhotoUrl(q: { imageDriveItemId?: string; imageData?: string }): string {
    if (q.imageDriveItemId) return this.validator.getSkillPhotoDownloadUrl(q.imageDriveItemId);
    if (q.imageData && (q.imageData as string).startsWith('data:')) return q.imageData as string;
    return '';
  }

  openReviewSkill(q: typeof this.skillQueries[0]): void {
    this.reviewSkillQuery = {
      id: q.id,
      skillName: q.skillName,
      make: q.make,
      model: q.model,
      assetId: q.assetId,
      attachment: q.attachment,
      photoUrl: q.photoUrl,
      ...(q.imageDriveItemId ? { imageDriveItemId: q.imageDriveItemId } : {}),
      ...(q.imageData ? { imageData: q.imageData } : {}),
    };
    this.reviewSkillForm = {
      skillName: q.skillName,
      make: q.make,
      model: q.model,
      assetId: q.assetId,
      attachment: q.attachment,
      photoUrl: q.photoUrl,
    };
    this.cdr.markForCheck();
  }

  closeReviewSkill(): void {
    this.reviewSkillQuery = null;
    this.cdr.markForCheck();
  }

  setReviewSkillFormSkillName(value: string): void {
    this.reviewSkillForm = { ...this.reviewSkillForm, skillName: value };
    this.cdr.markForCheck();
  }
  setReviewSkillFormMake(value: string): void {
    this.reviewSkillForm = { ...this.reviewSkillForm, make: value };
    this.cdr.markForCheck();
  }
  setReviewSkillFormModel(value: string): void {
    this.reviewSkillForm = { ...this.reviewSkillForm, model: value };
    this.cdr.markForCheck();
  }
  setReviewSkillFormAssetId(value: string): void {
    this.reviewSkillForm = { ...this.reviewSkillForm, assetId: value };
    this.cdr.markForCheck();
  }
  setReviewSkillFormAttachment(value: string): void {
    this.reviewSkillForm = { ...this.reviewSkillForm, attachment: value };
    this.cdr.markForCheck();
  }
  setReviewSkillFormPhotoUrl(value: string): void {
    this.reviewSkillForm = { ...this.reviewSkillForm, photoUrl: value };
    this.cdr.markForCheck();
  }

  saveReviewSkill(): void {
    if (!this.reviewSkillQuery) return;
    const id = this.reviewSkillQuery.id;
    const skillName = (this.reviewSkillForm.skillName ?? '').trim();
    const q = this.skillQueries.find(x => x.id === id);
    if (q) {
      if (skillName && !this.skillsList.includes(skillName)) {
        this.skillsList = [...this.skillsList, skillName];
        this.saveSkillsToStorage(this.skillsList);
      }
      this.skillQueries = this.skillQueries.filter(x => x.id !== id);
      this.saveSkillQueriesToStorage();
      this.skillQueriesAddedIds.delete(id);
      this.saveSkillQueriesAddedIdsToStorage();
    }
    this.closeReviewSkill();
  }

  openUnknownSkillDialog(skill: string): void {
    this.unknownSkillDialogSkill = (skill ?? '').trim();
    this.unknownSkillForm = { assetId: '', make: '', model: '', attachment: '', photoUrl: '', imageData: null };
    this.unknownSkillPhotoUploadError = null;
    this.unknownSkillDialogOpen = true;
    this.cdr.markForCheck();
  }

  closeUnknownSkillDialog(): void {
    this.unknownSkillDialogOpen = false;
    this.unknownSkillDialogSkill = '';
    this.unknownSkillForm = { assetId: '', make: '', model: '', attachment: '', photoUrl: '', imageData: null };
    this.cdr.markForCheck();
  }

  canSendUnknownSkill(): boolean {
    const f = this.unknownSkillForm;
    return (f.assetId ?? '').trim().length > 0 && (f.make ?? '').trim().length > 0 && (f.model ?? '').trim().length > 0;
  }

  sendUnknownSkillQuery(): void {
    if (!this.canSendUnknownSkill()) return;
    const f = this.unknownSkillForm;
    this.unknownSkillPhotoUploadError = null;

    if (f.imageData) {
      const file = this.dataUrlToFile(f.imageData, 'skill-photo.jpg');
      this.unknownSkillPhotoUploading = true;
      this.cdr.markForCheck();
      this.validator.uploadSkillPhoto(file, this.skillPhotosFolder).subscribe({
        next: (result) => {
          this.unknownSkillPhotoUploading = false;
          this.unknownSkillPhotoUploadError = null;
          this.addSkillQueryEntry(f, result.id);
          this.closeUnknownSkillDialog();
          this.cdr.markForCheck();
        },
        error: (err) => {
          this.unknownSkillPhotoUploading = false;
          this.unknownSkillPhotoUploadError = err?.message || err?.error?.error || 'Upload failed';
          this.cdr.markForCheck();
        }
      });
      return;
    }

    this.addSkillQueryEntry(f, undefined);
    this.closeUnknownSkillDialog();
    this.cdr.markForCheck();
  }

  private dataUrlToFile(dataUrl: string, fileName: string): File {
    const arr = dataUrl.split(',');
    const mime = (arr[0].match(/:(.*?);/) || [])[1] || 'image/jpeg';
    const bstr = atob(arr[1] || '');
    let n = bstr.length;
    const u8arr = new Uint8Array(n);
    while (n--) u8arr[n] = bstr.charCodeAt(n);
    return new File([u8arr], fileName, { type: mime });
  }

  private addSkillQueryEntry(f: typeof this.unknownSkillForm, imageDriveItemId: string | undefined): void {
    const id = this.nextSkillQueryId++;
    this.skillQueries = [
      ...this.skillQueries,
      {
        id,
        customer: this.testingCustomer.trim(),
        site: this.testingSite.trim(),
        user: this.testingUser.trim(),
        assetId: (f.assetId ?? '').trim(),
        skillName: this.unknownSkillDialogSkill,
        make: (f.make ?? '').trim(),
        model: (f.model ?? '').trim(),
        attachment: (f.attachment ?? '').trim(),
        photoUrl: (f.photoUrl ?? '').trim(),
        ...(imageDriveItemId ? { imageDriveItemId } : {}),
      },
    ];
    this.saveSkillQueriesToStorage();
  }

  onUnknownSkillImageDrop(event: DragEvent): void {
    event.preventDefault();
    event.stopPropagation();
    const file = event.dataTransfer?.files?.[0];
    if (file && file.type.startsWith('image/')) {
      const reader = new FileReader();
      reader.onload = () => {
        this.unknownSkillForm = { ...this.unknownSkillForm, imageData: reader.result as string };
        this.cdr.markForCheck();
      };
      reader.readAsDataURL(file);
    }
  }

  onUnknownSkillImageDragOver(event: DragEvent): void {
    event.preventDefault();
    event.stopPropagation();
  }

  onUnknownSkillImageFileInput(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input?.files?.[0];
    if (file && file.type.startsWith('image/')) {
      const reader = new FileReader();
      reader.onload = () => {
        this.unknownSkillForm = { ...this.unknownSkillForm, imageData: reader.result as string };
        this.cdr.markForCheck();
      };
      reader.readAsDataURL(file);
    }
    input.value = '';
  }

  clearUnknownSkillImage(): void {
    this.unknownSkillForm = { ...this.unknownSkillForm, imageData: null };
    this.cdr.markForCheck();
  }

  setUnknownSkillMake(value: string): void {
    this.unknownSkillForm = { ...this.unknownSkillForm, make: value };
    this.cdr.markForCheck();
  }

  setUnknownSkillModel(value: string): void {
    this.unknownSkillForm = { ...this.unknownSkillForm, model: value };
    this.cdr.markForCheck();
  }

  setUnknownSkillAttachment(value: string): void {
    this.unknownSkillForm = { ...this.unknownSkillForm, attachment: value };
    this.cdr.markForCheck();
  }

  setUnknownSkillPhotoUrl(value: string): void {
    this.unknownSkillForm = { ...this.unknownSkillForm, photoUrl: value };
    this.cdr.markForCheck();
  }

  setUnknownSkillAssetId(value: string): void {
    this.unknownSkillForm = { ...this.unknownSkillForm, assetId: value };
    this.cdr.markForCheck();
  }

  triggerUnknownSkillImageInput(): void {
    document.getElementById('unknown-skill-image-input')?.click();
  }

  openSettings(): void {
    this.settingsApiKeyInput = this.chatGptApiKey;
    this.settingsSkillPhotosFolderInput = this.skillPhotosFolder;
    this.settingsDialogOpen = true;
  }

  closeSettings(): void {
    this.settingsDialogOpen = false;
  }

  saveSettingsAndClose(): void {
    this.chatGptApiKey = this.settingsApiKeyInput.trim();
    this.skillPhotosFolder = (this.settingsSkillPhotosFolderInput ?? '').trim();
    try {
      if (this.chatGptApiKey) {
        localStorage.setItem(this.CHATGPT_API_KEY_STORAGE, this.chatGptApiKey);
      } else {
        localStorage.removeItem(this.CHATGPT_API_KEY_STORAGE);
      }
      if (this.skillPhotosFolder) {
        localStorage.setItem(this.SKILL_PHOTOS_FOLDER_STORAGE, this.skillPhotosFolder);
      } else {
        localStorage.removeItem(this.SKILL_PHOTOS_FOLDER_STORAGE);
      }
    } catch {
      /* ignore */
    }
    this.settingsDialogOpen = false;
  }

  @HostListener('document:keydown.escape')
  onEscape(): void {
    if (this.settingsDialogOpen) {
      this.closeSettings();
    } else if (this.unknownSkillDialogOpen) {
      this.closeUnknownSkillDialog();
    }
  }

  /** Build list of { firstName, lastName } from sheet rows (for ChatGPT). */
  getEmployeeNamePairs(sheet: EmployeeSheetResult): { firstName: string; lastName: string }[] {
    return (sheet.rows ?? []).map(r => ({
      firstName: (r.firstName ?? '').trim(),
      lastName: (r.lastName ?? '').trim(),
    }));
  }

  /** Whether this row has a high probability that first/last names are reversed (≥60%). */
  isNameReversedWarning(sheetName: string, rowIndex: number): boolean {
    const bySheet = this.currentPageState.nameCheckReversedProbability[sheetName];
    if (!bySheet) return false;
    const p = bySheet[rowIndex];
    return typeof p === 'number' && p >= 0.6;
  }

  closeNameCheckCompleteDialog(): void {
    this.nameCheckCompleteDialogOpen = false;
  }

  checkNames(sheet: EmployeeSheetResult): void {
    if (!sheet?.rows?.length) {
      this.currentPageState.nameCheckError = 'No employee data to check.';
      this.cdr.markForCheck();
      return;
    }
    const key = this.chatGptApiKey?.trim();
    if (!key) {
      this.currentPageState.nameCheckError = 'Add a ChatGPT API key in Settings first.';
      this.cdr.markForCheck();
      return;
    }
    const pairs = this.getEmployeeNamePairs(sheet);
    const eligible = (sheet.rows ?? [])
      .map((row, i) => ({ row, pair: pairs[i] }))
      .filter(({ pair }) => (pair.firstName.length > 0 && pair.lastName.length > 0 && pair.firstName !== pair.lastName));
    const lines = eligible.map(({ pair }) => `${pair.firstName} ${pair.lastName}`);
    const rowIndices = eligible.map(({ row }) => row.rowIndex);
    if (lines.length === 0) {
      this.currentPageState.nameCheckError = 'No names to check (need both first and last name for each row).';
      this.cdr.markForCheck();
      return;
    }
    this.currentPageState.nameCheckError = null;
    this.nameCheckLoading = true;
    if (!this.currentPageState.nameCheckReversedProbability[sheet.name]) this.currentPageState.nameCheckReversedProbability[sheet.name] = {};
    this.currentPageState.nameCheckReversedProbability[sheet.name] = {};
    this.cdr.markForCheck();

    const prompt = `Analyse a list of {first_name, last_name} pairs and estimate the probability that the names are reversed (i.e., the surname appears in the first-name field and the given name appears in the last-name field).

Use global reference lists only:
- A large list of global first names.
- A large list of global surnames / family names.

For each pair:
1. Compare the first_name against the global first-name list and the global surname list.
2. Compare the last_name against the global surname list and the global first-name list.
3. Estimate a probability_wrong_way_round based on the signals:
 - If first_name appears primarily in the surname list and last_name appears in the first-name list, assign a high probability that the pair is reversed.
 - If both tokens strongly match the expected pattern (first_name in first-name list and last_name in surname list), assign a low probability.
 - If signals are mixed or ambiguous, assign a moderate probability.

The list of pairs is below (one per line, format "first_name last_name"). Output a JSON array of numbers: one probability_wrong_way_round per line, same order as the pairs. Each number in the range 0–1. Output ONLY the JSON array, no other text.

${lines.join('\n')}`;

    const headers = new HttpHeaders({
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${key}`,
    });
    const body = {
      model: 'gpt-4o-mini',
      messages: [{ role: 'user' as const, content: prompt }],
      temperature: 0.2,
    };

    this.http.post<{ choices?: { message?: { content?: string } }[] }>(this.OPENAI_CHAT_URL, body, { headers }).subscribe({
      next: (res) => {
        this.nameCheckLoading = false;
        const content = res?.choices?.[0]?.message?.content?.trim();
        if (!content) {
          this.currentPageState.nameCheckError = 'No response from ChatGPT.';
          this.cdr.markForCheck();
          return;
        }
        const parsed = this.parseProbabilitiesJson(content);
        if (parsed.length > 0) {
          for (let j = 0; j < parsed.length && j < rowIndices.length; j++) {
            if (typeof parsed[j] === 'number') {
              this.currentPageState.nameCheckReversedProbability[sheet.name][rowIndices[j]] = parsed[j];
            }
          }
        } else {
          this.currentPageState.nameCheckError = 'Could not parse probabilities from response.';
        }
        this.cdr.markForCheck();
      },
      error: (err) => {
        this.nameCheckLoading = false;
        this.currentPageState.nameCheckError = err.error?.error?.message || err.message || 'ChatGPT request failed.';
        this.cdr.markForCheck();
      },
    });
  }

  /** For debug: list of { firstName, lastName, probability } for a sheet. */
  getNameCheckDebugList(sheet: EmployeeSheetResult): { firstName: string; lastName: string; probability: number }[] {
    if (!sheet?.rows?.length) return [];
    const bySheet = this.currentPageState.nameCheckReversedProbability[sheet.name] ?? {};
    return sheet.rows.map(row => ({
      firstName: (row.firstName ?? '').trim(),
      lastName: (row.lastName ?? '').trim(),
      probability: bySheet[row.rowIndex] ?? 0,
    }));
  }

  /** Extract JSON array of numbers from model output (may be wrapped in markdown or text). */
  private parseProbabilitiesJson(content: string): number[] {
    const trimmed = content.trim();
    let jsonStr = trimmed;
    const arrayMatch = trimmed.match(/\[[\s\d.,eE+-]+\]/);
    if (arrayMatch) {
      jsonStr = arrayMatch[0];
    }
    try {
      const arr = JSON.parse(jsonStr) as unknown;
      if (!Array.isArray(arr)) return [];
      return arr.map(x => {
        const n = typeof x === 'number' ? x : parseFloat(String(x));
        return Number.isFinite(n) ? Math.max(0, Math.min(1, n)) : 0;
      });
    } catch {
      return [];
    }
  }

  onFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (file) {
      const ext = file.name.toLowerCase().slice(-5);
      if (!['.xlsx', '.xls'].some(e => ext.endsWith(e))) {
        this.currentPageState.error = 'Please select an Excel file (.xlsx or .xls).';
        this.currentPageState.result = null;
        this.currentPageState.selectedFile = null;
        this.currentPageState.excelPreviewHtml = null;
        this.currentPageState.previewLoading = false;
        this.currentPageState.excelSheetNames = [];
        this.currentPageState.selectedSheetName = null;
        this.currentPageState.showSheetSelectDialog = false;
        return;
      }
      this.currentPageState.error = null;
      this.currentPageState.result = null;
      this.currentPageState.selectedFile = file;
      this.currentPageState.excelPreviewHtml = null;
      this.currentPageState.previewLoading = true;
      this.currentPageState.excelSheetNames = [];
      this.currentPageState.selectedSheetName = null;
      this.currentPageState.showSheetSelectDialog = false;
      file.arrayBuffer().then((ab) => {
        try {
          const wb = XLSX.read(ab, { type: 'array' });
          const names = wb.SheetNames || [];
          this.currentPageState.excelSheetNames = names;
          this.currentPageState.previewLoading = false;
          if (names.length === 0) {
            this.currentPageState.excelPreviewHtml = '<p>No sheets in workbook.</p>';
            this.currentPageState.selectedSheetName = null;
          } else if (names.length === 1) {
            this.currentPageState.selectedSheetName = null;
          } else {
            this.currentPageState.selectedSheetName = null;
          }
          this.cdr.markForCheck();
        } catch (e) {
          this.currentPageState.previewLoading = false;
          this.currentPageState.excelPreviewHtml = '<p>Could not parse Excel file.</p>';
          this.currentPageState.error = 'Could not preview file.';
          this.cdr.markForCheck();
        }
      }).catch(() => {
        this.currentPageState.excelPreviewHtml = null;
        this.currentPageState.previewLoading = false;
        this.currentPageState.error = 'Could not read file.';
        this.currentPageState.excelSheetNames = [];
        this.currentPageState.selectedSheetName = null;
        this.cdr.markForCheck();
      });
    }
  }

  /** Called when user changes the sheet dropdown; rebuilds preview for the selected sheet. */
  onSheetChange(sheetName: string): void {
    if (!this.currentPageState.selectedFile) return;
    const name = (sheetName ?? '').trim();
    this.currentPageState.selectedSheetName = name || null;
    if (!name) {
      this.currentPageState.excelPreviewHtml = null;
      this.currentPageState.previewLoading = false;
      this.previewColumnWidths = [];
      this.cdr.markForCheck();
      return;
    }
    this.currentPageState.previewLoading = true;
    this.buildExcelPreview(this.currentPageState.selectedFile, name);
  }

  /** Confirm selected sheet from dialog and build preview. (Kept for compatibility; dialog removed.) */
  confirmSheetSelection(): void {
    if (!this.currentPageState.selectedFile || !this.currentPageState.selectedSheetName) return;
    this.currentPageState.showSheetSelectDialog = false;
    this.currentPageState.previewLoading = true;
    this.buildExcelPreview(this.currentPageState.selectedFile, this.currentPageState.selectedSheetName);
  }

  closeSheetSelectDialog(): void {
    this.currentPageState.showSheetSelectDialog = false;
    this.clearPreview();
  }

  /** Parse the Excel file and set excelPreviewHtml for iframe preview; then open column mapping dialog. */
  private buildExcelPreview(file: File, sheetName: string): void {
    const page = this.topLevelTab;
    const state = this.pageState[page];
    state.previewLoading = true;
    file.arrayBuffer().then((ab) => {
      try {
        const wb = XLSX.read(ab, { type: 'array' });
        const ws = wb.Sheets[sheetName];
        if (!ws) {
          state.excelPreviewHtml = '<p>Sheet not found.</p>';
        } else {
          const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as (string | number)[][];
          const html = XLSX.utils.sheet_to_html(ws, { id: 'excel-preview-table' });
          state.excelPreviewHtml = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>body{margin:0;padding:8px} table{border-collapse:collapse;font:14px/1.4 sans-serif;} th,td{border:1px solid #ccc;padding:4px 8px;} th{background:#eee;}</style></head><body>' + html + '</body></html>';
          this.previewColumnWidths = [];
          state.excelColumnOptions = this.getColumnsWithData(data);
          const opts = state.excelColumnOptions;
          state.columnMapping = {};
          state.columnMappingDialogPage = 0;
          state.showColumnMappingDialog = opts.length > 0;
        }
      } catch (e) {
        state.excelPreviewHtml = '<p>Could not parse Excel file.</p>';
        state.error = 'Could not preview file.';
      }
      state.previewLoading = false;
      this.cdr.markForCheck();
    }).catch(() => {
      state.excelPreviewHtml = null;
      state.previewLoading = false;
      state.error = 'Could not read file.';
      this.cdr.markForCheck();
    });
  }

  /** From sheet data (array of rows), return columns that have at least one non-empty data cell (excluding header). */
  private getColumnsWithData(data: (string | number)[][]): ExcelColumnOption[] {
    if (!data || data.length < 2) return [];
    const headers = data[0].map(c => (c != null ? String(c).trim() : ''));
    const options: ExcelColumnOption[] = [];
    for (let col = 0; col < headers.length; col++) {
      let hasData = false;
      for (let row = 1; row < data.length; row++) {
        const cell = data[row][col];
        if (cell != null && String(cell).trim() !== '') {
          hasData = true;
          break;
        }
      }
      if (hasData) {
        const headerTitle = headers[col] || `Column ${col + 1}`;
        options.push({ index: col, title: headerTitle || `Column ${col + 1}` });
      }
    }
    return options;
  }

  readonly columnMappingColumns = COLUMN_MAPPING_COLUMNS;
  readonly columnMappingColumnsTraining = COLUMN_MAPPING_COLUMNS_TRAINING;

  /** Columns for the current page's Assign Fields dialog (Training vs Employees/Agency). Uses the page that has the dialog open so the correct fields are shown even if the user switched tabs before the preview loaded. */
  get columnMappingColumnsForPage(): { key: ColumnMappingKey | TrainingColumnMappingKey; label: string; required: boolean }[] {
    if (this.pageState['Training'].showColumnMappingDialog) return this.columnMappingColumnsTraining;
    return this.columnMappingColumns;
  }

  get currentColumnMappingColumn(): { key: ColumnMappingKey | TrainingColumnMappingKey; label: string; required: boolean } | null {
    const page = this.currentPageState.columnMappingDialogPage;
    return this.columnMappingColumnsForPage[page] ?? null;
  }

  /** True when required column mappings are all set. Employees/Agency: Employee ID, First Name, Last Name. Training: Skill, Event Type, Test Date, Result, Employee Number. */
  get hasRequiredColumnMapping(): boolean {
    const m = this.currentPageState.columnMapping;
    if (this.topLevelTab === 'Training') {
      const keys: TrainingColumnMappingKey[] = ['skill', 'eventType', 'testDate', 'result', 'employeeId'];
      return keys.every(k => (m[k] ?? '').trim() !== '');
    }
    const keys: ColumnMappingKey[] = ['employeeId', 'firstName', 'lastName'];
    return keys.every(k => (m[k] ?? '').trim() !== '');
  }

  /** Mappings to show above the preview: field label and selected Excel column. */
  get columnMappingSummary(): { key: string; label: string; excelColumn: string }[] {
    const m = this.currentPageState.columnMapping;
    return this.columnMappingColumnsForPage
      .map(col => ({ key: col.key, label: col.label, excelColumn: (m[col.key] ?? '').trim() }))
      .filter(item => item.excelColumn !== '');
  }

  /** Per-column mapping for table/overlay: one entry per Excel column index. label = field name if mapped, '' otherwise; mapped = true if green. */
  get columnMappingByColumn(): { label: string; mapped: boolean }[] {
    const opts = this.currentPageState.excelColumnOptions;
    const m = this.currentPageState.columnMapping;
    if (!opts.length) return [];
    const numCols = Math.max(...opts.map(o => o.index), 0) + 1;
    const titleByIndex = new Map(opts.map(o => [o.index, o.title]));
    const labelByTitle = new Map<string, string>();
    for (const col of this.columnMappingColumnsForPage) {
      const title = (m[col.key] ?? '').trim();
      if (title) labelByTitle.set(title, col.label);
    }
    const result: { label: string; mapped: boolean }[] = [];
    for (let i = 0; i < numCols; i++) {
      const title = titleByIndex.get(i) ?? '';
      const label = labelByTitle.get(title) ?? '';
      result.push({ label, mapped: label !== '' });
    }
    return result;
  }

  /** Same as columnMappingByColumn but length matches measured table columns when available, so mapping row/overlay align with iframe. */
  get previewMappingCells(): { label: string; mapped: boolean }[] {
    const byCol = this.columnMappingByColumn;
    const n = this.previewColumnWidths.length || byCol.length;
    const out: { label: string; mapped: boolean }[] = [];
    for (let i = 0; i < n; i++) {
      out.push(byCol[i] ?? { label: '', mapped: false });
    }
    return out;
  }

  /** Column options for the current Assign Fields page: only columns not already mapped to a different field (current field's selection is always included). */
  get availableColumnOptionsForMapping(): ExcelColumnOption[] {
    const col = this.currentColumnMappingColumn;
    const opts = this.currentPageState.excelColumnOptions;
    if (!col || !opts.length) return opts;
    const mapping = this.currentPageState.columnMapping;
    const currentVal = (mapping[col.key] ?? '').trim();
    const usedByOther = new Set(
      this.columnMappingColumnsForPage
        .filter(c => c.key !== col.key)
        .map(c => (mapping[c.key] ?? '').trim())
        .filter(Boolean)
    );
    return opts.filter(opt => currentVal === opt.title || !usedByOther.has(opt.title));
  }

  get canColumnMappingGoNext(): boolean {
    const col = this.currentColumnMappingColumn;
    if (!col) return false;
    if (col.required) {
      const val = this.currentPageState.columnMapping[col.key];
      return (val ?? '').trim() !== '';
    }
    return true;
  }

  /** On optional pages with nothing selected, show "Skip" instead of "Next". */
  get columnMappingNextButtonLabel(): string {
    const col = this.currentColumnMappingColumn;
    if (!col || col.required) return 'Next';
    const val = (this.currentPageState.columnMapping[col.key] ?? '').trim();
    return val === '' ? 'Skip' : 'Next';
  }

  columnMappingNext(): void {
    if (this.currentPageState.columnMappingDialogPage < this.columnMappingColumnsForPage.length - 1) {
      this.currentPageState.columnMappingDialogPage++;
      this.cdr.markForCheck();
    } else {
      this.closeColumnMappingDialog();
    }
  }

  columnMappingPrev(): void {
    if (this.currentPageState.columnMappingDialogPage > 0) {
      this.currentPageState.columnMappingDialogPage--;
      this.cdr.markForCheck();
    }
  }

  setColumnMappingValue(key: ColumnMappingKey | TrainingColumnMappingKey, value: string): void {
    this.currentPageState.columnMapping[key] = value;
    this.cdr.markForCheck();
  }

  closeColumnMappingDialog(): void {
    this.currentPageState.showColumnMappingDialog = false;
    this.currentPageState.columnMappingDialogPage = 0;
    this.cdr.markForCheck();
  }

  /** Clear all data on the current page's data tab and switch back to Import. */
  clearCurrentTabData(): void {
    const s = this.currentPageState;
    s.result = null;
    s.rowsToReverse = {};
    s.confirmedCells = {};
    s.nameCheckReversedProbability = {};
    s.nameCheckError = null;
    s.importedFileLabel = null;
    s.employeeTabShowOnlyInvalid = false;
    s.agencyTabShowOnlyInvalid = false;
    s.employeeTabFilterInvalidRowIndices = null;
    s.agencyTabFilterInvalidRowIndices = null;
    if (this.topLevelTab === 'Employees') s.employeesSubTab = 'Import';
    else if (this.topLevelTab === 'Agency Workers') s.agencySubTab = 'Import';
    else if (this.topLevelTab === 'Training') s.trainingSubTab = 'Import';
    this.cdr.markForCheck();
  }

  /** Remove preview and restore file upload interface. Does not clear result or switch tabs. */
  clearPreview(): void {
    this.currentPageState.excelPreviewHtml = null;
    this.previewColumnWidths = [];
    this.currentPageState.previewLoading = false;
    this.currentPageState.selectedFile = null;
    this.currentPageState.selectedSheetName = null;
    this.currentPageState.excelSheetNames = [];
    this.currentPageState.showSheetSelectDialog = false;
    this.currentPageState.excelColumnOptions = [];
    this.currentPageState.columnMapping = {};
    this.currentPageState.columnMappingDialogPage = 0;
    this.currentPageState.showColumnMappingDialog = false;
    this.currentPageState.importedFileLabel = null;
    this.currentPageState.error = null;
    const input = document.getElementById('file-input') as HTMLInputElement;
    const inputAgency = document.getElementById('file-input-agency') as HTMLInputElement;
    const inputTraining = document.getElementById('file-input-training') as HTMLInputElement;
    if (input) input.value = '';
    if (inputAgency) inputAgency.value = '';
    if (inputTraining) inputTraining.value = '';
    this.cdr.markForCheck();
  }

  validate(): void {
    if (!this.currentPageState.selectedFile) {
      this.currentPageState.error = 'Please select a file first.';
      return;
    }
    this.currentPageState.loading = true;
    this.currentPageState.error = null;
    this.currentPageState.result = null;
    this.currentPageState.rowsToReverse = {};
    this.currentPageState.confirmedCells = {};
    this.currentPageState.nameCheckReversedProbability = {};
    this.currentPageState.nameCheckError = null;
    this.currentPageState.employeeTabShowOnlyInvalid = false;
    this.currentPageState.agencyTabShowOnlyInvalid = false;
    this.currentPageState.employeeTabFilterInvalidRowIndices = null;
    this.currentPageState.agencyTabFilterInvalidRowIndices = null;
    this.currentPageState.trainingShowOnlyInvalid = false;

    const sheetType = this.topLevelTab === 'Training' ? 'training' : 'employees';
    const columnMapping = this.currentPageState.columnMapping;
    const columnMappingFiltered = columnMapping && Object.keys(columnMapping).length > 0
      ? Object.fromEntries(Object.entries(columnMapping).filter(([, v]) => (v ?? '').trim() !== ''))
      : undefined;
    this.validator.validateFile(this.currentPageState.selectedFile, {
      sheetName: this.currentPageState.selectedSheetName ?? undefined,
      sheetType,
      columnMapping: columnMappingFiltered
    }).subscribe({
      next: (res) => {
        this.currentPageState.result = res;
        this.runClientValidationForAllSheets();
        const visible = this.getVisibleTabs();
        if (visible.length > 0) {
          const employeesSheet = this.getSheetForTab('Employees');
          const agencySheet = this.getSheetForTab('Agency Employees');
          const instructorSheet = this.getSheetForTab('Instructor');
          this.currentPageState.activeTab = (employeesSheet?.rows?.length ?? 0) > 0 ? 'Employees'
            : (agencySheet?.rows?.length ?? 0) > 0 ? 'Agency Employees'
            : (instructorSheet?.rows?.length ?? 0) > 0 ? 'Instructor'
            : visible[0].id;
        }
        this.currentPageState.loading = false;
        this.currentPageState.importedFileLabel = this.currentPageState.selectedFile?.name ?? null;
        if (this.topLevelTab === 'Employees') this.setEmployeesSubTab('Employee Data');
        else if (this.topLevelTab === 'Agency Workers') this.setAgencySubTab('Agency Worker Data');
        else if (this.topLevelTab === 'Training') this.setTrainingSubTab('Training Data');
      },
      error: (err) => {
        this.currentPageState.error = err.error?.message || err.message || 'Validation failed.';
        this.currentPageState.loading = false;
      },
    });
  }

  reset(): void {
    const s = this.currentPageState;
    s.result = null;
    s.error = null;
    s.selectedFile = null;
    s.excelPreviewHtml = null;
    s.previewLoading = false;
    s.excelSheetNames = [];
    s.selectedSheetName = null;
    s.showSheetSelectDialog = false;
    s.importedFileLabel = null;
    s.rowsToReverse = {};
    s.confirmedCells = {};
    s.nameCheckReversedProbability = {};
    s.nameCheckError = null;
    s.employeeTabShowOnlyInvalid = false;
    s.agencyTabShowOnlyInvalid = false;
    s.employeeTabFilterInvalidRowIndices = null;
    s.agencyTabFilterInvalidRowIndices = null;
    s.employeesSubTab = 'Import';
    s.agencySubTab = 'Import';
    s.trainingSubTab = 'Import';
    s.activeTab = 'Employees';
    const input = document.getElementById('file-input') as HTMLInputElement;
    const inputAgency = document.getElementById('file-input-agency') as HTMLInputElement;
    const inputTraining = document.getElementById('file-input-training') as HTMLInputElement;
    if (input) input.value = '';
    if (inputAgency) inputAgency.value = '';
    if (inputTraining) inputTraining.value = '';
    this.cdr.markForCheck();
  }

  toggleRowToReverse(sheetName: string, rowIndex: number): void {
    if (!this.currentPageState.rowsToReverse[sheetName]) this.currentPageState.rowsToReverse[sheetName] = new Set();
    const set = this.currentPageState.rowsToReverse[sheetName];
    if (set.has(rowIndex)) set.delete(rowIndex);
    else set.add(rowIndex);
  }

  isRowSelectedForReverse(sheetName: string, rowIndex: number): boolean {
    return this.currentPageState.rowsToReverse[sheetName]?.has(rowIndex) ?? false;
  }

  getCorrectionsForExport(): { sheetName: string; rowIndices: number[] }[] {
    return Object.entries(this.currentPageState.rowsToReverse)
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
    const sheets = this.currentPageState.result?.employeeSheets ?? [];
    if (tabId === 'Employees') return sheets.find(s => s.name === 'Core Employees');
    if (tabId === 'Agency Employees') return sheets.find(s => s.name === 'Agency Employees');
    return sheets.find(s => s.name === 'Instructor');
  }

  /** Tabs that have data to display (sheet exists and has rows). Instructor tab shown whenever the sheet exists, even with 0 rows. */
  getVisibleTabs(): { id: 'Employees' | 'Agency Employees' | 'Instructor'; label: string }[] {
    return this.tabs;
  }

  setActiveTab(tabId: 'Employees' | 'Agency Employees' | 'Instructor'): void {
    this.currentPageState.activeTab = tabId;
  }

  setEmployeesSubTab(id: 'Import' | 'Employee Data'): void {
    this.currentPageState.employeesSubTab = id;
  }

  setAgencySubTab(id: 'Import' | 'Agency Worker Data'): void {
    this.currentPageState.agencySubTab = id;
  }

  setTrainingSubTab(id: 'Import' | 'Training Data'): void {
    this.currentPageState.trainingSubTab = id;
  }

  /** Revert the Import tab label and clear file selection so a new file can be chosen. Does not clear result or data tabs. */
  clearImportedFileLabel(): void {
    this.currentPageState.importedFileLabel = null;
    this.currentPageState.selectedFile = null;
    const input = document.getElementById('file-input') as HTMLInputElement;
    const inputAgency = document.getElementById('file-input-agency') as HTMLInputElement;
    const inputTraining = document.getElementById('file-input-training') as HTMLInputElement;
    if (input) input.value = '';
    if (inputAgency) inputAgency.value = '';
    if (inputTraining) inputTraining.value = '';
    this.cdr.markForCheck();
  }

  getEmployeeSheet(): EmployeeSheetResult | undefined {
    return this.getSheetForTab('Employees');
  }

  getAgencySheet(): EmployeeSheetResult | undefined {
    return this.getSheetForTab('Agency Employees');
  }

  getEmployeeAttentionCount(): number {
    const sheet = this.getEmployeeSheet();
    if (!sheet?.rows?.length) return 0;
    const idLabel = this.getIdLabel(sheet);
    return sheet.rows.filter(row => this.hasUnconfirmedRowErrors(row, sheet, idLabel)).length;
  }

  getAgencyAttentionCount(): number {
    const sheet = this.getAgencySheet();
    if (!sheet?.rows?.length) return 0;
    const idLabel = this.getIdLabel(sheet);
    return sheet.rows.filter(row => this.hasUnconfirmedRowErrors(row, sheet, idLabel)).length;
  }

  toggleEmployeeTabShowFilter(): void {
    this.currentPageState.employeeTabShowOnlyInvalid = !this.currentPageState.employeeTabShowOnlyInvalid;
    if (this.currentPageState.employeeTabShowOnlyInvalid) {
      const sheet = this.getEmployeeSheet();
      if (sheet?.rows?.length) {
        const idLabel = this.getIdLabel(sheet);
        this.currentPageState.employeeTabFilterInvalidRowIndices = sheet.rows
          .filter(row => this.hasUnconfirmedRowErrors(row, sheet, idLabel))
          .map(r => r.rowIndex);
      } else {
        this.currentPageState.employeeTabFilterInvalidRowIndices = [];
      }
    } else {
      this.currentPageState.employeeTabFilterInvalidRowIndices = null;
    }
    this.cdr.markForCheck();
  }

  toggleAgencyTabShowFilter(): void {
    this.currentPageState.agencyTabShowOnlyInvalid = !this.currentPageState.agencyTabShowOnlyInvalid;
    if (this.currentPageState.agencyTabShowOnlyInvalid) {
      const sheet = this.getAgencySheet();
      if (sheet?.rows?.length) {
        const idLabel = this.getIdLabel(sheet);
        this.currentPageState.agencyTabFilterInvalidRowIndices = sheet.rows
          .filter(row => this.hasUnconfirmedRowErrors(row, sheet, idLabel))
          .map(r => r.rowIndex);
      } else {
        this.currentPageState.agencyTabFilterInvalidRowIndices = [];
      }
    } else {
      this.currentPageState.agencyTabFilterInvalidRowIndices = null;
    }
    this.cdr.markForCheck();
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
    const training = this.currentPageState.result?.trainingSheet;
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
        this.currentPageState.error = null;
        this.cdr.markForCheck();
      },
      error: (err) => {
        this.currentPageState.error = err?.error?.message || err?.message || 'Failed to add skill';
        this.cdr.markForCheck();
      }
    });
  }

  getTrainingSkillOptions(): string[] {
    return this.skillsList.length > 0 ? this.skillsList : (this.currentPageState.result?.trainingSheet?.skillOptions ?? []);
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
    return this.currentPageState.result?.trainingSheet;
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
    if (!this.currentPageState.trainingShowOnlyInvalid) return sorted;
    return sorted.filter(row => !row.isValid);
  }

  toggleTrainingShowFilter(): void {
    this.currentPageState.trainingShowOnlyInvalid = !this.currentPageState.trainingShowOnlyInvalid;
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
    const training = this.currentPageState.result?.trainingSheet;
    if (!training?.rows?.length) return 0;
    return training.rows.filter(r => !r.isValid).length;
  }

  hasReversedNameErrors(): boolean {
    return (this.currentPageState.result?.employeeSheets ?? []).some(sheet => (sheet.reversedNameErrors?.length ?? 0) > 0);
  }

  /** Label for the identifier column: either "Employee ID" or "Employee Number" per backend. */
  getEmployeeIdentifierColumnLabel(): string {
    const first = this.currentPageState.result?.employeeSheets?.find(s => (s.rows?.length ?? 0) > 0);
    return first?.employeeIdentifierColumnLabel === 'Employee Number' ? 'Employee Number' : 'Employee ID';
  }

  /** Run client-side validation on all sheets so spaceErrors/onlySpaceErrors are set for display. */
  private runClientValidationForAllSheets(): void {
    if (!this.currentPageState.result?.employeeSheets) return;
    for (const sheet of this.currentPageState.result.employeeSheets) {
      const rows = sheet.rows ?? [];
      if (rows.length === 0) continue;
      const idLabel = sheet.employeeIdentifierColumnLabel === 'Employee Number' ? 'Employee Number' : 'Employee ID';
      revalidateSheetRows(rows, idLabel);
    }
    this.revalidateTrainingSheetAgainstSkillsList();
    this.updateSummary();
  }

  /** When the user has a Skills list (Maintenance), re-validate Training sheet skills against it so imported files use the latest list. */
  private revalidateTrainingSheetAgainstSkillsList(): void {
    const training = this.currentPageState.result?.trainingSheet;
    if (!training?.rows?.length || this.skillsList.length === 0) return;
    const options = this.getTrainingSkillOptions();
    const normalizedOptions = options.map(s => (s ?? '').trim().toLowerCase());
    for (const row of training.rows) {
      const skillNorm = (row.skill ?? '').trim().toLowerCase();
      const valid = !skillNorm || normalizedOptions.some(opt => opt === skillNorm);
      row.skillError = row.skill && !valid ? 'Skill not recognised' : undefined;
      this.updateTrainingRowValidity(row);
    }
    training.valid = training.rows.every(r => r.isValid);
    this.recomputeTrainingDuplicates();
  }

  /** Tooltip for a cell: only when this specific cell has an error (missing, spaces) or row-level error in the responsible cell. For field 'nameReversed', pass sheet so we can check confirmation. */
  getCellTooltip(row: ValidationRow, field: 'employeeId' | 'firstName' | 'lastName' | 'email' | 'nameReversed', idLabel: string, sheet?: { name: string }): string | null {
    if (field === 'nameReversed') {
      if (sheet && this.isNameReversedWarning(sheet.name, row.rowIndex) && !this.isCellConfirmed(sheet.name, row.rowIndex, 'nameReversed')) return 'First and last names may be reversed';
      return null;
    }
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

  /** Single tooltip to show for a cell (at most one icon). For firstName/lastName prefers validation error over name-reversed. */
  getCellDisplayTooltip(row: ValidationRow, field: 'employeeId' | 'firstName' | 'lastName' | 'email', idLabel: string, sheet: { name: string }): string | null {
    if (field === 'firstName' || field === 'lastName') {
      const validationTip = this.getCellTooltip(row, field, idLabel);
      if (validationTip && !this.isCellConfirmed(sheet.name, row.rowIndex, field)) return validationTip;
      if (this.isNameReversedWarning(sheet.name, row.rowIndex) && !this.isCellConfirmed(sheet.name, row.rowIndex, 'nameReversed')) return 'First and last names may be reversed';
      return null;
    }
    const tip = this.getCellTooltip(row, field, idLabel);
    return tip && !this.isCellConfirmed(sheet.name, row.rowIndex, field) ? tip : null;
  }

  /** True when this cell has a validation error (not the reversed-name warning) and is not confirmed. Used for red overlay. */
  hasValidationError(sheet: { name: string; employeeIdentifierColumnLabel?: string }, row: ValidationRow, field: 'employeeId' | 'firstName' | 'lastName' | 'email'): boolean {
    const tip = this.getCellTooltip(row, field, this.getIdLabel(sheet));
    return tip != null && !this.isCellConfirmed(sheet.name, row.rowIndex, field);
  }

  /** True when the reversed-name check flags this row and it is not confirmed. Used for yellow overlay on first/last name cells. */
  hasReversedWarning(sheet: { name: string }, row: ValidationRow): boolean {
    return this.isNameReversedWarning(sheet.name, row.rowIndex) && !this.isCellConfirmed(sheet.name, row.rowIndex, 'nameReversed');
  }

  /** True when this cell should show the confirm button (has error but not a space error). Do not show for empty required fields. For nameReversed pass sheet. */
  showConfirmButton(row: ValidationRow, field: 'employeeId' | 'firstName' | 'lastName' | 'email' | 'nameReversed', idLabel: string, sheet?: { name: string }): boolean {
    if (field === 'nameReversed') return sheet ? this.getCellTooltip(row, 'nameReversed', idLabel, sheet) != null : false;
    if (field === 'email') return this.getCellTooltip(row, 'email', idLabel) != null;
    const val = field === 'employeeId' ? row.employeeId : field === 'firstName' ? row.firstName : row.lastName;
    if ((val ?? '').toString().trim() === '') return false;
    if (row.spaceErrors?.[field]) return false;
    return this.getCellTooltip(row, field, idLabel) != null;
  }
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

  getFilteredEmployeeRows(sheet: { name: string; rows?: ValidationRow[]; employeeIdentifierColumnLabel?: string }, showOnlyInvalid: boolean, filterRowIndices?: number[] | null): ValidationRow[] {
    const sorted = this.getSortedRows(sheet);
    if (!showOnlyInvalid) return sorted;
    if (filterRowIndices != null && Array.isArray(filterRowIndices)) {
      const set = new Set(filterRowIndices);
      return sorted.filter(row => set.has(row.rowIndex));
    }
    const idLabel = this.getIdLabel(sheet);
    return sorted.filter(row => this.hasUnconfirmedRowErrors(row, sheet, idLabel));
  }

  toggleEmployeesShowFilter(): void {
    this.currentPageState.employeeTabShowOnlyInvalid = !this.currentPageState.employeeTabShowOnlyInvalid;
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

  isCellConfirmed(sheetName: string, rowIndex: number, field: 'employeeId' | 'firstName' | 'lastName' | 'email' | 'nameReversed'): boolean {
    return this.currentPageState.confirmedCells[sheetName]?.has(this.cellKey(rowIndex, field)) ?? false;
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
    const fields: ('employeeId' | 'firstName' | 'lastName' | 'email' | 'nameReversed')[] = ['employeeId', 'firstName', 'lastName', 'email', 'nameReversed'];
    for (const field of fields) {
      const tip = field === 'nameReversed' ? this.getCellTooltip(row, field, idLabel, sheet) : this.getCellTooltip(row, field, idLabel);
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
    const fields: ('employeeId' | 'firstName' | 'lastName' | 'email' | 'nameReversed')[] = ['employeeId', 'firstName', 'lastName', 'email', 'nameReversed'];
    for (const field of fields) {
      if (this.showConfirmButton(row, field, idLabel, field === 'nameReversed' ? sheet : undefined)) this.confirmCell(sheet, row, field);
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

  confirmCell(sheet: { name: string; rows?: ValidationRow[] }, row: ValidationRow, field: 'employeeId' | 'firstName' | 'lastName' | 'email' | 'nameReversed'): void {
    if (field === 'nameReversed') {
      const sheetName = sheet.name;
      if (!this.currentPageState.confirmedCells[sheetName]) this.currentPageState.confirmedCells[sheetName] = new Set();
      this.currentPageState.confirmedCells[sheetName].add(this.cellKey(row.rowIndex, 'nameReversed'));
      this.cdr.markForCheck();
      return;
    }
    if (field === 'email') {
      const sheetName = sheet.name;
      if (!this.currentPageState.confirmedCells[sheetName]) this.currentPageState.confirmedCells[sheetName] = new Set();
      this.currentPageState.confirmedCells[sheetName].add(this.cellKey(row.rowIndex, 'email'));
      this.cdr.markForCheck();
      return;
    }
    if (field === 'employeeId' && sheet.name === 'Instructor' && ((row.employeeId ?? '').toString().trim() === '')) {
      const sheetName = sheet.name;
      if (!this.currentPageState.confirmedCells[sheetName]) this.currentPageState.confirmedCells[sheetName] = new Set();
      this.currentPageState.confirmedCells[sheetName].add(this.cellKey(row.rowIndex, 'employeeId'));
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
    if (!this.currentPageState.confirmedCells[sheetName]) this.currentPageState.confirmedCells[sheetName] = new Set();
    this.currentPageState.confirmedCells[sheetName].add(this.cellKey(row.rowIndex, field));
    this.cdr.markForCheck();
  }

  private removeConfirmedCellsForRow(sheetName: string, rowIndex: number): void {
    const set = this.currentPageState.confirmedCells[sheetName];
    if (!set) return;
    for (const key of Array.from(set)) {
      if (key.startsWith(`${rowIndex}-`)) set.delete(key);
    }
  }

  /** Row has at least one validation error that is not confirmed (so row should show as invalid). */
  hasUnconfirmedRowErrors(row: ValidationRow, sheet: { name: string; rows?: ValidationRow[]; showEmailColumn?: boolean }, idLabel: string): boolean {
    if (this.isNameReversedWarning(sheet.name, row.rowIndex) && !this.isCellConfirmed(sheet.name, row.rowIndex, 'nameReversed')) return true;
    if (row.isValid) return false;
    const fields: ('employeeId' | 'firstName' | 'lastName' | 'email')[] = ['employeeId', 'firstName', 'lastName', 'email'];
    for (const field of fields) {
      const tip = this.getCellTooltip(row, field, idLabel);
      if (tip && !this.isCellConfirmed(sheet.name, row.rowIndex, field)) return true;
    }
    return false;
  }

  /** Count of rows that still need attention (have unconfirmed errors). Updates when edits are made or cells are confirmed/deleted. */
  /** Count of rows needing attention across all employee sheets (for backward compatibility). */
  getEmployeesWhoNeedAttentionCount(): number {
    return this.getEmployeeAttentionCount() + this.getAgencyAttentionCount();
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
    if (this.isNameReversedWarning(sheet.name, row.rowIndex)) {
      hasError = true;
      if (!this.isCellConfirmed(sheet.name, row.rowIndex, 'nameReversed')) allConfirmed = false;
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
    if (!this.currentPageState.result?.summary) return;
    let total = 0, valid = 0;
    for (const sheet of this.currentPageState.result.employeeSheets ?? []) {
      const rows = sheet.rows ?? [];
      for (const r of rows) {
        total++;
        if (r.isValid) valid++;
      }
    }
    this.currentPageState.result.summary.totalRows = total;
    this.currentPageState.result.summary.validRows = valid;
    this.currentPageState.result.summary.invalidRows = total - valid;
  }

  exportCorrected(): void {
    const corrections = this.getCorrectionsForExport();
    if (!this.currentPageState.selectedFile || corrections.length === 0) {
      this.currentPageState.error = 'Select at least one row to reverse and ensure a file was validated.';
      return;
    }
    this.exporting = true;
    this.currentPageState.error = null;
    this.validator.correctAndExport(this.currentPageState.selectedFile, corrections).subscribe({
      next: (blob) => {
        const base = this.currentPageState.result?.fileName?.replace(/\.(xlsx?|xls)$/i, '') || 'export';
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${base}_corrected.xlsx`;
        a.click();
        URL.revokeObjectURL(url);
        this.exporting = false;
      },
      error: (err) => {
        this.currentPageState.error = err.error?.message || err.message || 'Export failed.';
        this.exporting = false;
      }
    });
  }
}
