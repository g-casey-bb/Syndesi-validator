import { Routes } from '@angular/router';
import { AppComponent } from './app.component';

export const routes: Routes = [
  { path: 'employees', component: AppComponent },
  { path: 'agency-workers', component: AppComponent },
  { path: 'training', component: AppComponent },
  { path: 'assets', component: AppComponent },
  { path: '', pathMatch: 'full', redirectTo: 'employees' },
  { path: '**', redirectTo: 'employees' },
];
