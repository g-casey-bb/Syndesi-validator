import { HttpInterceptorFn, HttpErrorResponse } from '@angular/common/http';
import { catchError, throwError } from 'rxjs';

export const apiErrorInterceptor: HttpInterceptorFn = (req, next) => {
  return next(req).pipe(
    catchError((err: HttpErrorResponse) => {
      const msg = err?.message ?? '';
      const body = (err?.error != null && typeof err.error === 'string') ? err.error : '';
      const gotHtml =
        msg.includes('<!DOCTYPE') ||
        msg.includes('is not valid JSON') ||
        (body.trim().startsWith('<') && body.toLowerCase().includes('doctype'));
      if (gotHtml) {
        return throwError(() =>
          new Error(
            'API returned a page instead of data. Start the backend (npm start in the backend folder) and use the frontend dev server (npm start in the frontend folder), then open http://localhost:4200'
          )
        );
      }
      return throwError(() => err);
    })
  );
};
