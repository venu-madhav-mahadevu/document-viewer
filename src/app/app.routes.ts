import { Routes } from '@angular/router';

// Lazy routes: each page is its own chunk. The view page (and the viewer behind it) is
// preloaded in the background after the home page has loaded (see app.config.ts).
export const routes: Routes = [
  { path: '', loadComponent: () => import('./home/home').then((m) => m.HomePage), title: 'Docs Viewer' },
  { path: 'view', loadComponent: () => import('./view/view').then((m) => m.ViewPage), title: 'Docs Viewer' },
  { path: '**', redirectTo: '' },
];
