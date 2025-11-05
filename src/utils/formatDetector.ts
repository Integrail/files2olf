import JSZip from 'jszip';

/**
 * Supported Office file formats
 */
export type OfficeFormat = 'xlsx' | 'xls' | 'pptx' | 'ppt' | 'docx' | 'doc' | 'unknown';

/**
 * Detect Office file format by examining file signature and internal structure
 * @param buffer - File content as Buffer
 * @returns Detected format type
 */
export async function detectOfficeFormat(buffer: Buffer): Promise<OfficeFormat> {
  if (buffer.length < 4) {
    return 'unknown';
  }

  // Check magic bytes (first 4 bytes)
  const signature = buffer.slice(0, 4).toString('hex');

  // ZIP-based formats (XLSX, PPTX) - signature: 50 4B 03 04
  if (signature === '504b0304' || signature === '504b0506' || signature === '504b0708') {
    let zip: JSZip | undefined;
    try {
      zip = await JSZip.loadAsync(buffer);

      // Check for PPTX structure
      if (zip.file('ppt/presentation.xml')) {
        return 'pptx';
      }

      // Check for XLSX structure
      if (zip.file('xl/workbook.xml')) {
        return 'xlsx';
      }

      // Check for DOCX structure
      if (zip.file('word/document.xml')) {
        return 'docx';
      }

      return 'unknown';
    } catch {
      return 'unknown';
    } finally {
      // Clean up JSZip instance to prevent memory leak
      if (zip) {
        Object.keys(zip.files).forEach(key => {
          delete zip!.files[key];
        });
        zip = undefined;
      }
    }
  }

  // OLE Compound File Binary (XLS, PPT, DOC) - signature: D0 CF 11 E0
  if (signature === 'd0cf11e0') {
    // Both XLS and PPT use same container format
    // Distinguish by checking for format-specific markers
    const format = detectOleFormat(buffer);
    return format;
  }

  return 'unknown';
}

/**
 * Detect whether OLE file is XLS, PPT, or DOC by checking internal markers
 */
function detectOleFormat(buffer: Buffer): 'xls' | 'ppt' | 'doc' | 'unknown' {
  const content = buffer.toString('binary', 0, Math.min(buffer.length, 8192));

  // PowerPoint markers (UTF-16LE encoded strings)
  if (
    content.includes('PowerPoint Document') ||
    content.includes('Current User') ||
    content.includes('P\x00o\x00w\x00e\x00r\x00P\x00o\x00i\x00n\x00t')
  ) {
    return 'ppt';
  }

  // Excel markers
  if (
    content.includes('Workbook') ||
    content.includes('W\x00o\x00r\x00k\x00b\x00o\x00o\x00k') ||
    buffer.includes(Buffer.from([0x09, 0x08])) // BOF record for BIFF8
  ) {
    return 'xls';
  }

  // Word markers (UTF-16LE encoded strings)
  if (
    content.includes('Word.Document') ||
    content.includes('W\x00o\x00r\x00d\x00D\x00o\x00c\x00u\x00m\x00e\x00n\x00t') ||
    content.includes('Microsoft Word')
  ) {
    return 'doc';
  }

  return 'unknown';
}

/**
 * Check if buffer is a supported Excel format (XLSX or XLS)
 */
export async function isExcelFile(buffer: Buffer): Promise<boolean> {
  const format = await detectOfficeFormat(buffer);
  return format === 'xlsx' || format === 'xls';
}

/**
 * Check if buffer is a supported PowerPoint format (PPTX only)
 */
export async function isPowerPointFile(buffer: Buffer): Promise<boolean> {
  const format = await detectOfficeFormat(buffer);
  return format === 'pptx';
}
