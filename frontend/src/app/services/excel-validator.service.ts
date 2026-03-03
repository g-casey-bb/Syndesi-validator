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
}
