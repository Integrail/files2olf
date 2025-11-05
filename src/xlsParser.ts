import * as XLSX from 'xlsx';
import { parseXlsx } from './xlsxParser';
import { XlsxParseOptions, XlsxParseResult } from './xlsxTypes';

// Maximum file size: 100MB
const MAX_XLS_FILE_SIZE = 100 * 1024 * 1024;

/**
 * Parse an XLS (Excel 97-2003 binary) file by converting to XLSX format
 *
 * IMPORTANT: Image extraction is not supported for XLS files due to limitations
 * in the SheetJS Community Edition. For files with images, convert to XLSX first.
 *
 * @param xlsBuffer - Buffer containing the XLS file content
 * @param options - Optional parsing options (same as parseXlsx)
 * @returns Promise resolving to the parsed workbook data
 * @throws TypeError if xlsBuffer is not a Buffer
 * @throws Error if the XLS file is invalid or cannot be parsed
 */
export async function parseXls(
  xlsBuffer: Buffer,
  options?: XlsxParseOptions
): Promise<XlsxParseResult> {
  // Input validation
  if (!Buffer.isBuffer(xlsBuffer)) {
    throw new TypeError('xlsBuffer must be a Buffer');
  }
  if (xlsBuffer.length === 0) {
    throw new Error('xlsBuffer is empty');
  }
  if (xlsBuffer.length > MAX_XLS_FILE_SIZE) {
    throw new Error(
      `File size ${xlsBuffer.length} bytes exceeds maximum ${MAX_XLS_FILE_SIZE} bytes (100MB)`
    );
  }

  let sheetjsWorkbook: XLSX.WorkBook | undefined;
  let xlsxBuffer: Buffer | undefined;

  try {
    // Read XLS file with SheetJS
    sheetjsWorkbook = XLSX.read(xlsBuffer, {
      type: 'buffer',
      cellDates: true,   // Parse dates as Date objects
      cellFormula: true, // Include formulas
      cellStyles: false  // We don't need styles for data extraction
    });

    // Convert to XLSX format in-memory
    const xlsxArrayBuffer = XLSX.write(sheetjsWorkbook, {
      type: 'buffer',
      bookType: 'xlsx',
      cellDates: true
    });

    // Clear SheetJS workbook before parseXlsx to reduce peak memory
    sheetjsWorkbook = undefined;

    xlsxBuffer = Buffer.from(xlsxArrayBuffer);

    // Use existing XLSX parser
    // Note: includeImages option is ignored for XLS (images cannot be extracted from XLS)
    const result = await parseXlsx(xlsxBuffer, {
      ...options,
      includeImages: false // Override - XLS cannot provide images
    });

    return result;
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    throw new Error(`Failed to parse XLS file: ${errorMessage}`);
  } finally {
    // Clean up references to allow garbage collection
    sheetjsWorkbook = undefined;
    xlsxBuffer = undefined;
  }
}
