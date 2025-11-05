import ExcelJS from 'exceljs';
import { Sheet, XlsxParseOptions, XlsxParseResult } from './xlsxTypes';
import { extractTables } from './xlsxTableExtractor';
import { extractImages } from './xlsxImageExtractor';

/**
 * Parse an XLSX file and extract sheets with their content, tables, and images
 * @param xlsxBuffer - Buffer containing the XLSX file content
 * @param options - Optional parsing options
 * @returns Promise resolving to the parsed workbook data
 * @throws TypeError if xlsxBuffer is not a Buffer
 * @throws Error if the XLSX file is invalid or cannot be parsed
 */
export async function parseXlsx(xlsxBuffer: Buffer, options?: XlsxParseOptions): Promise<XlsxParseResult> {
  // Input validation
  if (!Buffer.isBuffer(xlsxBuffer)) {
    throw new TypeError('xlsxBuffer must be a Buffer');
  }
  if (xlsxBuffer.length === 0) {
    throw new Error('xlsxBuffer is empty');
  }

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(xlsxBuffer as any);

    // Process sheets either sequentially or in parallel based on options
    const sheets = options?.parallel
      ? await processSheetsParallel(workbook, options)
      : await processSheetsSequential(workbook, options);

    return { sheets };
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    throw new Error(`Failed to parse XLSX file: ${errorMessage}`);
  }
}

/**
 * Process sheets sequentially (default behavior)
 */
async function processSheetsSequential(
  workbook: ExcelJS.Workbook,
  options?: XlsxParseOptions
): Promise<Sheet[]> {
  const sheets: Sheet[] = [];

  workbook.eachSheet((worksheet, sheetId) => {
    const sheet = processSheet(worksheet, sheetId - 1, workbook, options);
    sheets.push(sheet);
  });

  return sheets;
}

/**
 * Process sheets in parallel for better performance
 */
async function processSheetsParallel(
  workbook: ExcelJS.Workbook,
  options?: XlsxParseOptions
): Promise<Sheet[]> {
  const sheetPromises: Promise<Sheet>[] = [];

  workbook.eachSheet((worksheet, sheetId) => {
    const promise = Promise.resolve(
      processSheet(worksheet, sheetId - 1, workbook, options)
    );
    sheetPromises.push(promise);
  });

  return Promise.all(sheetPromises);
}

/**
 * Process a single worksheet
 */
function processSheet(
  worksheet: ExcelJS.Worksheet,
  index: number,
  workbook: ExcelJS.Workbook,
  options?: XlsxParseOptions
): Sheet {
  // Extract merged cells
  const mergedCells = extractMergedCells(worksheet);

  // Extract tables
  const tables = extractTables(worksheet, mergedCells, options);

  // Extract images if requested
  const images = options?.includeImages !== false
    ? extractImages(worksheet, workbook)
    : [];

  return {
    name: worksheet.name,
    index,
    tables,
    mergedCells,
    images
  };
}

/**
 * Extract merged cell ranges from worksheet
 */
function extractMergedCells(worksheet: ExcelJS.Worksheet): import('./xlsxTypes').MergedCellRange[] {
  const mergedCells: import('./xlsxTypes').MergedCellRange[] = [];

  // Access merged cell ranges from ExcelJS
  const merges: string[] = (worksheet.model as any).merges || [];

  for (const mergeRef of merges) {
    const range = parseCellRange(mergeRef);
    const topLeftCell = worksheet.getCell(range.startRow, range.startCol);

    mergedCells.push({
      ref: mergeRef,
      startRow: range.startRow,
      startCol: range.startCol,
      endRow: range.endRow,
      endCol: range.endCol,
      value: topLeftCell.value as any,
      colSpan: range.endCol - range.startCol + 1,
      rowSpan: range.endRow - range.startRow + 1
    });
  }

  return mergedCells;
}

/**
 * Parse a cell range reference (e.g., "A1:C3") to coordinates
 */
function parseCellRange(ref: string): {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
} {
  const parts = ref.split(':');
  const start = cellAddressToCoords(parts[0]);
  const end = parts.length > 1 ? cellAddressToCoords(parts[1]) : start;

  return {
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col
  };
}

/**
 * Convert cell address (e.g., "A1") to coordinates
 */
function cellAddressToCoords(address: string): { row: number; col: number } {
  const match = address.match(/^([A-Z]+)(\d+)$/);
  if (!match) throw new Error(`Invalid cell address: ${address}`);

  const colLetters = match[1];
  const row = parseInt(match[2]);

  let col = 0;
  for (let i = 0; i < colLetters.length; i++) {
    col = col * 26 + (colLetters.charCodeAt(i) - 64);
  }

  return { row, col };
}
