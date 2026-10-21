// Types for presence.js (the "online now" counter).
export function startPresence(container: HTMLElement | null, options: { env: 'staging' | 'angular' | 'prod' }): { setPage(page: 'home' | 'view'): void };
