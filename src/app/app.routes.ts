import { Routes } from '@angular/router';
import { HomePage } from './home/home';
import { ViewPage } from './view/view';

export const routes: Routes = [
  { path: '', component: HomePage, title: 'Docs Viewer' },
  { path: 'view', component: ViewPage, title: 'Docs Viewer' },
  { path: '**', redirectTo: '' },
];
