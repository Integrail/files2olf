import ExcelJS from 'exceljs';
import { Table, CellData, CellType, MergedCellRange, XlsxParseOptions } from './xlsxTypes';
import { convertTableToJson } from './xlsxJsonConverter';
import { coordsToAddress } from './utils/excelCoordinates';
import { convertTableToMarkdown as convertToMarkdownTable } from './utils/markdown';

/**
 * Extract tables from a worksheet
 */
export function extractTables(
  worksheet: ExcelJS.Worksheet,
  mergedCells: MergedCellRange[],
  options?: XlsxParseOptions
): Table[] {
  const tables: Table[] = [];

  // Detect table regions from the worksheet
  const tableRegions = detectTableRegions(worksheet);

  for (const region of tableRegions) {
    const table = extractTable(worksheet, region, mergedCells, options);
    if (table) {
      tables.push(table);
    }
  }

  return tables;
}

/**
 * Detect table regions in the worksheet
 */
function detectTableRegions(worksheet: ExcelJS.Worksheet): TableRegion[] {
  const regions: TableRegion[] = [];

  // Get actual dimensions of the worksheet
  const dimension = worksheet.dimensions;
  if (!dimension) return regions;

  // Simple heuristic: Find contiguous rectangular regions
  // For MVP, treat entire used range as one table
  const startRow = dimension.top;
  const endRow = dimension.bottom;
  const startCol = dimension.left;
  const endCol = dimension.right;

  if (startRow && endRow && startCol && endCol) {
    regions.push({
      name: `Table1`,
      startRow,
      startCol,
      endRow,
      endCol
    });
  }

  return regions;
}

/**
 * Extract a single table from a region
 */
function extractTable(
  worksheet: ExcelJS.Worksheet,
  region: TableRegion,
  mergedCells: MergedCellRange[],
  options?: XlsxParseOptions
): Table | null {
  // Extract cell data
  const data: CellData[][] = [];

  for (let row = region.startRow; row <= region.endRow; row++) {
    const rowData: CellData[] = [];

    for (let col = region.startCol; col <= region.endCol; col++) {
      const cell = worksheet.getCell(row, col);
      const cellData = extractCellData(cell);
      rowData.push(cellData);
    }

    data.push(rowData);
  }

  // Identify merged cells within this table
  const tableMergedCells = mergedCells.filter(merge =>
    merge.startRow >= region.startRow && merge.endRow <= region.endRow &&
    merge.startCol >= region.startCol && merge.endCol <= region.endCol
  );

  // Extract column headers (first row)
  const columns = data.length > 0
    ? data[0].map(cell => String(cell.value || ''))
    : [];

  // Build range string
  const range = `${coordsToAddress(region.startRow, region.startCol)}:${coordsToAddress(region.endRow, region.endCol)}`;

  // Check if has hierarchical headers
  const hasHierarchicalHeaders = tableMergedCells.some(
    merge => merge.startRow < region.startRow + 2 && merge.colSpan > 1
  );

  // Generate markdown
  const rows = data.map(row => row.map(cell => String(cell.value ?? '')));
  const markdown = convertToMarkdownTable(rows);

  const table: Table = {
    name: region.name,
    range,
    columns,
    data,
    mergedHeaders: tableMergedCells,
    markdown,
    hasHierarchicalHeaders
  };

  // Convert to JSON if requested
  if (options?.convertToJson) {
    table.json = convertTableToJson(table);
  }

  return table;
}

/**
 * Type definition for complex cell values
 */
interface CellComplexValue {
  formula?: string;
  result?: any;
  text?: string;
  richText?: Array<{ text: string }>;
  hyperlink?: string;
}

/**
 * Extract data from a single cell
 */
function extractCellData(cell: ExcelJS.Cell): CellData {
  const address = cell.address;
  const row = Number(cell.row);
  const col = Number(cell.col);

  const { value, type, formula } = extractCellValue(cell);

  return {
    address,
    row,
    col,
    value,
    type,
    formula
  };
}

/**
 * Extract value, type, and formula from a cell
 */
function extractCellValue(cell: ExcelJS.Cell): {
  value: string | number | boolean | Date | null;
  type: CellType;
  formula?: string;
} {
  // Empty cell
  if (cell.value === null || cell.value === undefined) {
    return { value: null, type: CellType.Empty };
  }

  // Complex cell types (formulas, rich text, hyperlinks)
  if (typeof cell.value === 'object') {
    return extractComplexCellValue(cell);
  }

  // Primitive types
  if (typeof cell.value === 'string') {
    return { value: cell.value, type: CellType.String };
  }

  if (typeof cell.value === 'boolean') {
    return { value: cell.value, type: CellType.Boolean };
  }

  if (typeof cell.value === 'number') {
    return extractNumberCellValue(cell);
  }

  // Unknown type
  return { value: String(cell.value), type: CellType.String };
}

/**
 * Extract complex cell values (formulas, rich text, dates, hyperlinks)
 */
function extractComplexCellValue(cell: ExcelJS.Cell): {
  value: string | number | boolean | Date | null;
  type: CellType;
  formula?: string;
} {
  const cellValue = cell.value as CellComplexValue;

  // Formula cell
  if (cellValue.formula) {
    return {
      type: CellType.Formula,
      formula: cellValue.formula,
      value: cellValue.result ?? null
    };
  }

  // Rich text (simple text property)
  if (cellValue.text) {
    return { value: cellValue.text, type: CellType.String };
  }

  // Rich text (array of styled text)
  if (cellValue.richText) {
    const text = cellValue.richText.map((rt) => rt.text).join('');
    return { value: text, type: CellType.String };
  }

  // Hyperlink
  if (cellValue.hyperlink) {
    return {
      value: cellValue.text || cellValue.hyperlink,
      type: CellType.String
    };
  }

  // Date object
  if (cell.value instanceof Date) {
    return { value: cell.value, type: CellType.Date };
  }

  // Unknown object type
  return { value: String(cell.value), type: CellType.String };
}

/**
 * Extract number cell value (may be date or number)
 */
function extractNumberCellValue(cell: ExcelJS.Cell): {
  value: number | Date;
  type: CellType;
} {
  const numValue = cell.value as number;

  // Check if it's a date (Excel stores dates as numbers)
  if (cell.numFmt && isDateFormat(cell.numFmt)) {
    return {
      value: excelDateToJSDate(numValue),
      type: CellType.Date
    };
  }

  return {
    value: numValue,
    type: CellType.Number
  };
}

/**
 * Check if number format indicates a date
 */
function isDateFormat(numFmt: string): boolean {
  return numFmt.includes('d') || numFmt.includes('m') || numFmt.includes('y');
}

/**
 * Convert Excel date number to JavaScript Date
 */
function excelDateToJSDate(excelDate: number): Date {
  // Excel stores dates as days since 1900-01-01
  // JavaScript Date uses milliseconds since 1970-01-01
  const millisecondsPerDay = 24 * 60 * 60 * 1000;
  const excelEpoch = new Date(1900, 0, 1).getTime();
  // Excel incorrectly treats 1900 as a leap year, so subtract 1 day for dates after Feb 28, 1900
  const daysOffset = excelDate > 59 ? excelDate - 2 : excelDate - 1;
  return new Date(excelEpoch + daysOffset * millisecondsPerDay);
}



/**
 * Table region detected in worksheet
 */
interface TableRegion {
  name: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}
