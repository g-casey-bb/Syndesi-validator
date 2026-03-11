import { Routes } from '@angular/router';
import { AppComponent } from './app.component';
import { WorkspaceEmbedComponent } from './workspace-embed.component';

export const routes: Routes = [
  { path: 'excel-google-workspace', component: WorkspaceEmbedComponent },
  { path: 'employees', component: AppComponent },
  { path: 'agency-workers', component: AppComponent },
  { path: 'users', component: AppComponent },
  { path: 'instructors', component: AppComponent },
  { path: 'training', component: AppComponent },
  { path: 'assets', component: AppComponent },
  { path: '', pathMatch: 'full', redirectTo: 'employees' },
  { path: '**', redirectTo: 'employees' },
];
