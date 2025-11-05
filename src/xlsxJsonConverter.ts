import { Table, TableJson, CellData, MergedCellRange } from './xlsxTypes';

/**
 * Convert a table to JSON array with support for nested headers
 */
export function convertTableToJson(table: Table): TableJson[] {
  if (!table.hasHierarchicalHeaders || table.mergedHeaders.length === 0) {
    // Simple flat conversion
    return convertFlatTable(table);
  }

  // Complex nested conversion with merged headers
  return convertNestedTable(table);
}

/**
 * Convert a flat table (no merged headers) to JSON
 */
function convertFlatTable(table: Table): TableJson[] {
  const data = table.data;
  if (data.length < 2) return [];

  const headers = data[0].map(cell => String(cell.value || `Column${cell.col}`));
  const jsonRows: TableJson[] = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const jsonRow: TableJson = {};

    for (let j = 0; j < row.length; j++) {
      const header = headers[j] || `Column${j + 1}`;
      jsonRow[header] = row[j].value;
    }

    jsonRows.push(jsonRow);
  }

  return jsonRows;
}

/**
 * Convert a table with nested/merged headers to hierarchical JSON
 */
function convertNestedTable(table: Table): TableJson[] {
  // Step 1: Detect header depth
  const headerDepth = detectHeaderDepth(table);

  // Step 2: Build header hierarchy from merged cells
  const headerTree = buildHeaderTree(table, headerDepth);

  // Step 3: Convert data rows to JSON using header tree
  const jsonRows: TableJson[] = [];

  for (let rowIdx = headerDepth; rowIdx < table.data.length; rowIdx++) {
    const row = table.data[rowIdx];
    const jsonRow = buildJsonRow(row, headerTree);
    jsonRows.push(jsonRow);
  }

  return jsonRows;
}

/**
 * Detect how many rows are headers (vs data)
 */
function detectHeaderDepth(table: Table): number {
  const data = table.data;
  let headerDepth = 1; // Minimum 1 header row

  // Check merged cells - if any merged cells span columns in early rows, those are headers
  const firstRow = table.data[0] && table.data[0][0] ? table.data[0][0].row : 1;
  const mergedInEarlyRows = table.mergedHeaders.filter(
    merge => merge.colSpan > 1 && merge.startRow <= firstRow + 2
  );

  if (mergedInEarlyRows.length > 0) {
    // Find the deepest row with merged headers
    const deepestMergedRow = Math.max(
      ...mergedInEarlyRows.map(merge => merge.endRow)
    );

    // Header depth is the row after the last merged header (relative to table start)
    headerDepth = deepestMergedRow - firstRow + 1;
  }

  // Ensure header depth doesn't exceed table size
  return Math.min(headerDepth, data.length - 1);
}

/**
 * Build hierarchical header tree from merged cells
 */
function buildHeaderTree(table: Table, headerDepth: number): HeaderNode[] {
  const data = table.data;
  const numCols = data[0].length;
  const headerTree: HeaderNode[] = [];

  // Build a header node for each final column
  for (let colIdx = 0; colIdx < numCols; colIdx++) {
    const col = data[0][colIdx].col;
    const node = buildHeaderNodeForColumn(col, table, headerDepth);
    headerTree.push(node);
  }

  return headerTree;
}

/**
 * Build header path for a specific column
 */
function buildHeaderNodeForColumn(
  col: number,
  table: Table,
  headerDepth: number
): HeaderNode {
  const path: string[] = [];
  const firstDataRow = table.data[0] && table.data[0][0] ? table.data[0][0].row : 1;

  // For each header level, find the header text for this column
  for (let level = 0; level < headerDepth; level++) {
    const row = firstDataRow + level;

    // Find if this cell is part of a merged range
    const mergedCell = table.mergedHeaders.find(
      merge =>
        merge.startRow === row &&
        col >= merge.startCol &&
        col <= merge.endCol
    );

    if (mergedCell) {
      // Use the merged cell's value
      path.push(String(mergedCell.value || ''));
    } else {
      // Use the cell's direct value
      const cellData = table.data[level].find(c => c.col === col);
      path.push(String(cellData?.value || ''));
    }
  }

  return {
    column: col,
    path,
    finalHeader: path[path.length - 1]
  };
}

/**
 * Build a JSON object for a data row using header tree
 */
function buildJsonRow(row: CellData[], headerTree: HeaderNode[]): TableJson {
  const jsonRow: TableJson = {};

  for (let i = 0; i < row.length && i < headerTree.length; i++) {
    const cell = row[i];
    const node = headerTree[i];

    // Build nested object following the header path
    let current = jsonRow;

    for (let pathIdx = 0; pathIdx < node.path.length - 1; pathIdx++) {
      const key = node.path[pathIdx];

      if (!current[key]) {
        current[key] = {};
      }

      current = current[key];
    }

    // Set the final value
    const finalKey = node.path[node.path.length - 1];
    current[finalKey] = cell.value;
  }

  return jsonRow;
}

/**
 * Convert coordinates to cell address
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
 * Header node representing column's header hierarchy
 */
interface HeaderNode {
  column: number;
  path: string[];
  finalHeader: string;
}

/**
 * Table region definition
 */
interface TableRegion {
  name: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}
