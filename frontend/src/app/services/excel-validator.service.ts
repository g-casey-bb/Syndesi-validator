import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';
import { ValidationResult } from '../models/validation-result';

const API_URL = 'http://localhost:3000/api';

@Injectable({ providedIn: 'root' })
export class ExcelValidatorService {
  constructor(private http: HttpClient) {}

  validateFile(file: File): Observable<ValidationResult> {
    const formData = new FormData();
    formData.append('file', file);
    return this.http.post<ValidationResult>(`${API_URL}/validate`, formData);
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
