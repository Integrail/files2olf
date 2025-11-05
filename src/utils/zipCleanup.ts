import JSZip from 'jszip';

/**
 * Clean up JSZip instance to prevent memory leaks
 *
 * JSZip maintains internal references to file contents which can hold
 * significant memory (10-100MB+ for large Office files). This utility
 * ensures all references are cleared to allow garbage collection.
 *
 * @param zip - JSZip instance to clean up
 */
export function cleanupZip(zip: JSZip | undefined): void {
  if (!zip) return;

  // Remove all file references to allow garbage collection
  Object.keys(zip.files).forEach(key => {
    const file = zip!.files[key];

    // Clear internal data buffers if present
    if ((file as any)._data) {
      (file as any)._data = null;
    }

    delete zip!.files[key];
  });
}
