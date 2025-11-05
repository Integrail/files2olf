import { Table, TableJson, CellData, MergedCellRange } from './xlsxTypes';
import { coordsToAddress } from './utils/excelCoordinates';

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
  if (table.data.length === 0) return 0;

  let headerDepth = 1; // Minimum 1 header row
  const firstRow = table.data[0]?.[0]?.row ?? 1;

  // Find deepest merged header without creating intermediate arrays
  let deepestMergedRow = firstRow;
  for (const merge of table.mergedHeaders) {
    if (merge.colSpan > 1 && merge.startRow <= firstRow + 2) {
      deepestMergedRow = Math.max(deepestMergedRow, merge.endRow);
    }
  }

  if (deepestMergedRow > firstRow) {
    headerDepth = deepestMergedRow - firstRow + 1;
  }

  // Ensure header depth doesn't exceed table size
  return Math.min(headerDepth, table.data.length - 1);
}

/**
 * Build an index of merged cells by row for O(1) lookups
 */
function buildMergeIndex(mergedHeaders: MergedCellRange[]): Map<number, MergedCellRange[]> {
  const index = new Map<number, MergedCellRange[]>();

  for (const merge of mergedHeaders) {
    for (let row = merge.startRow; row <= merge.endRow; row++) {
      if (!index.has(row)) {
        index.set(row, []);
      }
      index.get(row)!.push(merge);
    }
  }

  return index;
}

/**
 * Build hierarchical header tree from merged cells
 */
function buildHeaderTree(table: Table, headerDepth: number): HeaderNode[] {
  const data = table.data;
  const numCols = data[0].length;
  const headerTree: HeaderNode[] = [];

  // Build merge index once for O(1) lookups
  const mergeIndex = buildMergeIndex(table.mergedHeaders);

  // Build a header node for each final column
  for (let colIdx = 0; colIdx < numCols; colIdx++) {
    const col = data[0][colIdx].col;
    const node = buildHeaderNodeForColumn(col, table, headerDepth, mergeIndex);
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
  headerDepth: number,
  mergeIndex: Map<number, MergedCellRange[]>
): HeaderNode {
  const path: string[] = [];
  const firstDataRow = table.data[0]?.[0]?.row ?? 1;

  // For each header level, find the header text for this column
  for (let level = 0; level < headerDepth; level++) {
    const row = firstDataRow + level;
    const rowData = table.data[level];

    // O(1) lookup from index instead of O(n) find
    const rowMerges = mergeIndex.get(row) || [];
    const mergedCell = rowMerges.find(
      merge => col >= merge.startCol && col <= merge.endCol
    );

    if (mergedCell) {
      // Use the merged cell's value
      path.push(String(mergedCell.value || ''));
    } else {
      // Direct array indexing for O(1) access
      const firstColInRow = rowData[0]?.col ?? 1;
      const colIndex = col - firstColInRow;
      const cellData = rowData[colIndex];
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
 * Header node representing column's header hierarchy
 */
interface HeaderNode {
  column: number;
  path: string[];
  finalHeader: string;
}
