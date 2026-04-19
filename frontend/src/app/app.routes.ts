import { Routes } from '@angular/router';
import { DashboardComponent } from './pages/dashboard-component/dashboard-component';

export const routes: Routes = [
  {
    path: '',
    component: DashboardComponent,
  },
  {
    path: '**',
    redirectTo: '',
    pathMatch: 'full',
  },
];
