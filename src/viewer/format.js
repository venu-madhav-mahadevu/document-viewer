// Small formatting helpers, in their own module so pages can use them without loading the viewer.

/** 1536 -> "1.5 KB" */
export function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  const units = ['KB', 'MB', 'GB'];
  let value = bytes / 1024;
  let i = 0;
  while (value >= 1024 && i < units.length - 1) { value /= 1024; i++; }
  return `${value >= 100 ? Math.round(value) : value.toFixed(1).replace(/\.0$/, '')} ${units[i]}`;
}

/** 1234 -> "1.2 s" */
export function formatDuration(ms) {
  if (ms < 1000) return `${(ms / 1000).toFixed(2)} s`;
  if (ms < 10000) return `${(ms / 1000).toFixed(1)} s`;
  return `${Math.round(ms / 1000)} s`;
}
