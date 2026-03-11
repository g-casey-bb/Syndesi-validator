import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';
import { ValidationResult } from '../models/validation-result';

/** API base: relative so dev server proxy (proxy.conf.json → localhost:3000) is used. Run Syndesi backend on port 3000. */
const API_URL = '/api';

@Injectable({ providedIn: 'root' })
export class ExcelValidatorService {
  constructor(private http: HttpClient) {}

  validateFile(file: File, options?: { sheetName?: string; sheetType?: 'employees' | 'training'; columnMapping?: Record<string, string> }): Observable<ValidationResult> {
    const hasMapping = options?.columnMapping && Object.keys(options.columnMapping).length > 0;
    if (hasMapping && options.columnMapping) {
      return this.validateFileWithMapping(file, {
        sheetName: options.sheetName,
        sheetType: options.sheetType,
        columnMapping: options.columnMapping
      });
    }
    const formData = new FormData();
    if (options?.sheetName) formData.append('sheetName', options.sheetName);
    if (options?.sheetType) formData.append('sheetType', options.sheetType);
    formData.append('file', file);
    return this.http.post<ValidationResult>(`${API_URL}/validate`, formData);
  }

  private validateFileWithMapping(file: File, options: { sheetName?: string; sheetType?: 'employees' | 'training'; columnMapping: Record<string, string> }): Observable<ValidationResult> {
    return new Observable<ValidationResult>((observer) => {
      const reader = new FileReader();
      reader.onload = () => {
        const arrayBuffer = reader.result as ArrayBuffer;
        const bytes = new Uint8Array(arrayBuffer);
        let binary = '';
        for (let i = 0; i < bytes.length; i++) {
          binary += String.fromCharCode(bytes[i]);
        }
        const fileBase64 = btoa(binary);
        const body = {
          fileBase64,
          fileName: file.name,
          sheetName: options.sheetName ?? null,
          sheetType: options.sheetType ?? null,
          columnMapping: options.columnMapping
        };
        this.http.post<ValidationResult>(`${API_URL}/validate-json`, body).subscribe({
          next: (res) => { observer.next(res); observer.complete(); },
          error: (err) => { observer.error(err); }
        });
      };
      reader.onerror = () => observer.error(reader.error);
      reader.readAsArrayBuffer(file);
    });
  }

  /** Add a new skill to training.json; returns updated skillOptions. */
  addTrainingSkill(skill: string): Observable<{ success: boolean; skillOptions: string[] }> {
    return this.http.post<{ success: boolean; skillOptions: string[] }>(`${API_URL}/training/skill`, { skill });
  }

  /** Request corrected Excel with First/Last name swapped for the given rows per sheet. */
  correctAndExport(file: File, corrections: { sheetName: string; rowIndices: number[] }[]): Observable<Blob> {
    const formData = new FormData();
    formData.append('file', file);
    formData.append('corrections', JSON.stringify(corrections));
    return this.http.post(`${API_URL}/correct-export`, formData, {
      responseType: 'blob'
    });
  }

  /** Upload a skill photo (image). Returns { id } for download. */
  uploadSkillPhoto(file: File, folder: string): Observable<{ id: string }> {
    const formData = new FormData();
    formData.append('file', file);
    if (folder) formData.append('folder', folder);
    return this.http.post<{ id: string }>(`${API_URL}/skill-photo/upload`, formData);
  }

  /** URL to download a skill photo by id (backend proxy). */
  getSkillPhotoDownloadUrl(id: string): string {
    return `${API_URL}/skill-photo/download/${encodeURIComponent(id)}`;
  }
}
