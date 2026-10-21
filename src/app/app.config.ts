import { ApplicationConfig, isDevMode, provideBrowserGlobalErrorListeners, provideZonelessChangeDetection } from '@angular/core';
import { PreloadAllModules, provideRouter, withComponentInputBinding, withHashLocation, withPreloading } from '@angular/router';
import { provideServiceWorker } from '@angular/service-worker';

import { routes } from './app.routes';

export const appConfig: ApplicationConfig = {
  providers: [
    provideBrowserGlobalErrorListeners(),
    // No zone.js: change detection runs only when signals change.
    provideZonelessChangeDetection(),
    provideRouter(
      routes,
      // Hash URLs (#/view?file=…) work on any static host, including GitHub Pages sub-folders.
      withHashLocation(),
      // Query parameters arrive in components as signal inputs.
      withComponentInputBinding(),
      // Lazy pages are fetched in the background once the app is idle.
      withPreloading(PreloadAllModules),
    ),
    // Caches the app and libraries so repeat visits start instantly (production builds only).
    provideServiceWorker('ngsw-worker.js', {
      enabled: !isDevMode(),
      registrationStrategy: 'registerWhenStable:30000',
    }),
  ],
};
