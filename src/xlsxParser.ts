import ExcelJS from 'exceljs';
import { Sheet, XlsxParseOptions, XlsxParseResult, MergedCellRange } from './xlsxTypes';
import { extractTables } from './xlsxTableExtractor';
import { extractImages } from './xlsxImageExtractor';
import { parseCellRange } from './utils/excelCoordinates';

// Maximum file size: 100MB
const MAX_XLSX_FILE_SIZE = 100 * 1024 * 1024;
// Maximum number of sheets
const MAX_SHEETS = 100;

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
  if (xlsxBuffer.length > MAX_XLSX_FILE_SIZE) {
    throw new Error(
      `File size ${xlsxBuffer.length} bytes exceeds maximum ${MAX_XLSX_FILE_SIZE} bytes (100MB)`
    );
  }

  let workbook: ExcelJS.Workbook | undefined;

  try {
    workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(xlsxBuffer as any);

    // Validate sheet count
    const sheetCount = workbook.worksheets.length;
    if (sheetCount > MAX_SHEETS) {
      throw new Error(
        `Workbook has ${sheetCount} sheets, maximum is ${MAX_SHEETS}`
      );
    }

    // Process sheets sequentially
    const sheets = await processSheetsSequential(workbook, options);

    return { sheets };
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    throw new Error(`Failed to parse XLSX file: ${errorMessage}`);
  } finally {
    // CRITICAL: Clean up ExcelJS workbook to prevent memory leaks
    if (workbook) {
      // Clear worksheet references
      workbook.eachSheet((worksheet) => {
        worksheet.destroy();
      });
      // Null the workbook reference to allow garbage collection
      workbook = undefined;
    }
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
function extractMergedCells(worksheet: ExcelJS.Worksheet): MergedCellRange[] {
  const mergedCells: MergedCellRange[] = [];

  // Access merged cell ranges from ExcelJS
  // Note: worksheet.model.merges is an internal property
  const worksheetModel = worksheet.model as { merges?: string[] };
  const merges: string[] = worksheetModel.merges || [];

  for (const mergeRef of merges) {
    const range = parseCellRange(mergeRef);
    const topLeftCell = worksheet.getCell(range.startRow, range.startCol);

    // Extract cell value safely
    let cellValue: string | number | boolean | Date | null = null;
    if (topLeftCell.value !== null && topLeftCell.value !== undefined) {
      if (typeof topLeftCell.value === 'object' && 'result' in topLeftCell.value) {
        cellValue = topLeftCell.value.result as any;
      } else {
        cellValue = topLeftCell.value as any;
      }
    }

    mergedCells.push({
      ref: mergeRef,
      startRow: range.startRow,
      startCol: range.startCol,
      endRow: range.endRow,
      endCol: range.endCol,
      value: cellValue,
      colSpan: range.endCol - range.startCol + 1,
      rowSpan: range.endRow - range.startRow + 1
    });
  }

  return mergedCells;
}
