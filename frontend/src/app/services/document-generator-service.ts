import { HttpClient } from '@angular/common/http';
import { inject, Injectable } from '@angular/core';
import { Observable } from 'rxjs';
import { environment } from '../../environments/environment.development';

@Injectable({
  providedIn: 'root',
})
export class DocumentGeneratorService {
  private http = inject(HttpClient);

  private readonly API_URL = environment.apiUrl;
  private readonly endpoint = 'smart-sheet-reporter/api/beta/generator/generate';

  generateDocuments(excel: File, template: File): Observable<Blob> {
    const formData = new FormData();
    formData.append('excel', excel);
    formData.append('template', template);

    return this.http.post(`${this.API_URL}/${this.endpoint}`, formData, {
      responseType: 'blob',
    });
  }
}
