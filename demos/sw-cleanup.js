// The in-browser LibreOffice demo used to install its helper service worker
// (coi-serviceworker) for the whole demos/ folder. It adds cross-origin isolation
// headers that break the spreadsheet worker on other demo pages. The demo now lives
// in demos/convert/ with its own scope; this removes the old, wider registration.
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.getRegistrations().then(async (regs) => {
    let removed = false;
    for (const reg of regs) {
      const scope = new URL(reg.scope).pathname;
      const script = reg.active?.scriptURL || reg.installing?.scriptURL || reg.waiting?.scriptURL || '';
      if (scope.endsWith('/demos/') && /coi-serviceworker\.js$/.test(script)) {
        await reg.unregister();
        removed = true;
      }
    }
    // Reload once so this page runs without the old worker's headers.
    if (removed && navigator.serviceWorker.controller && !sessionStorage.getItem('dv-sw-cleaned')) {
      sessionStorage.setItem('dv-sw-cleaned', '1');
      location.reload();
    }
  }).catch(() => {});
}
