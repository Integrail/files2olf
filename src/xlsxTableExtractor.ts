import ExcelJS from 'exceljs';
import { Table, CellData, CellType, MergedCellRange, XlsxParseOptions } from './xlsxTypes';
import { convertTableToJson } from './xlsxJsonConverter';

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
  const markdown = convertTableToMarkdown(data);

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
 * Extract data from a single cell
 */
function extractCellData(cell: ExcelJS.Cell): CellData {
  const address = cell.address;
  const row = typeof cell.row === 'number' ? cell.row : parseInt(String(cell.row));
  const col = typeof cell.col === 'number' ? cell.col : parseInt(String(cell.col));

  let value: string | number | boolean | Date | null = null;
  let type: CellType = CellType.Empty;
  let formula: string | undefined;

  // Check cell type
  if (cell.value === null || cell.value === undefined) {
    type = CellType.Empty;
    value = null;
  } else if (typeof cell.value === 'object') {
    // Handle complex cell types (formulas, rich text, hyperlinks)
    const cellValue = cell.value as any;

    if (cellValue.formula) {
      // Formula cell
      type = CellType.Formula;
      formula = cellValue.formula;
      value = cellValue.result ?? null;
    } else if (cellValue.text) {
      // Rich text
      type = CellType.String;
      value = cellValue.text;
    } else if (cellValue.richText) {
      // Rich text array
      type = CellType.String;
      value = cellValue.richText.map((rt: any) => rt.text).join('');
    } else if (cellValue.hyperlink) {
      // Hyperlink
      type = CellType.String;
      value = cellValue.text || cellValue.hyperlink;
    } else if (cellValue instanceof Date) {
      // Date
      type = CellType.Date;
      value = cellValue;
    } else {
      // Unknown object type - convert to string
      type = CellType.String;
      value = String(cellValue);
    }
  } else if (typeof cell.value === 'string') {
    type = CellType.String;
    value = cell.value;
  } else if (typeof cell.value === 'number') {
    // Check if it's a date (Excel stores dates as numbers)
    if (cell.numFmt && (cell.numFmt.includes('d') || cell.numFmt.includes('m') || cell.numFmt.includes('y'))) {
      type = CellType.Date;
      // Convert Excel date number to Date object
      value = excelDateToJSDate(cell.value);
    } else {
      type = CellType.Number;
      value = cell.value;
    }
  } else if (typeof cell.value === 'boolean') {
    type = CellType.Boolean;
    value = cell.value;
  }

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
 * Convert coordinates to cell address (e.g., (1, 1) -> "A1")
 */
function coordsToAddress(row: number, col: number): string {
  let colLetter = '';
  let tempCol = col;

  while (tempCol > 0) {
    const remainder = (tempCol - 1) % 26;
    colLetter = String.fromCharCode(65 + remainder) + colLetter;
    tempCol = Math.floor((tempCol - 1) / 26);
  }

  return `${colLetter}${row}`;
}

/**
 * Convert table data to markdown format
 */
function convertTableToMarkdown(data: CellData[][]): string {
  if (data.length === 0) return '';

  const rows = data.map(row =>
    row.map(cell => String(cell.value ?? ''))
  );

  const lines: string[] = [];

  // Header row
  if (rows.length > 0) {
    const header = rows[0].join(' | ');
    lines.push('| ' + header + ' |');
    lines.push('| ' + rows[0].map(() => '---').join(' | ') + ' |');
  }

  // Data rows
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i].join(' | ');
    lines.push('| ' + row + ' |');
  }

  return lines.join('\n');
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
