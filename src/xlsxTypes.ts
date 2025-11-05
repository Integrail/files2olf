/**
 * Type definitions for XLSX parser
 */

/**
 * Options for parsing an XLSX document
 */
export interface XlsxParseOptions {
  /**
   * Process sheets in parallel for better performance (default: false)
   */
  parallel?: boolean;

  /**
   * Convert tables to nested JSON structures based on merged headers (default: false)
   */
  convertToJson?: boolean;

  /**
   * Include embedded images (default: true)
   */
  includeImages?: boolean;
}

/**
 * Result of parsing an XLSX document
 */
export interface XlsxParseResult {
  /** Array of sheets extracted from the workbook */
  sheets: Sheet[];
}

/**
 * Represents a single worksheet with its content
 */
export interface Sheet {
  /** The sheet name */
  name: string;

  /** The sheet index (0-indexed) */
  index: number;

  /** Tables detected in this sheet */
  tables: Table[];

  /** Merged cell ranges */
  mergedCells: MergedCellRange[];

  /** Images embedded in this sheet */
  images: SheetImage[];
}

/**
 * Represents a table (formal Excel table or detected table region)
 */
export interface Table {
  /** Table name (if formal table) or generated name */
  name: string;

  /** Cell range (e.g., "A1:D10") */
  range: string;

  /** Column headers */
  columns: string[];

  /** Raw table data (2D array, includes all header rows) */
  data: CellData[][];

  /** Merged cell ranges within this table (for nested headers) */
  mergedHeaders: MergedCellRange[];

  /** Markdown representation of the table */
  markdown: string;

  /** JSON representation (only if convertToJson option is true) */
  json?: TableJson[];

  /** Whether table has multi-level headers */
  hasHierarchicalHeaders: boolean;
}

/**
 * Cell data representation
 */
export interface CellData {
  /** Cell address (e.g., "A1") */
  address: string;

  /** Row index (1-based, matching Excel) */
  row: number;

  /** Column index (1-based, matching Excel) */
  col: number;

  /** Cell value (parsed) */
  value: string | number | boolean | Date | null;

  /** Cell type */
  type: CellType;

  /** Formula (if cell contains formula) */
  formula?: string;
}

/**
 * Cell type enumeration
 */
export enum CellType {
  Number = 'number',
  String = 'string',
  Boolean = 'boolean',
  Date = 'date',
  Formula = 'formula',
  Empty = 'empty'
}

/**
 * Merged cell range
 */
export interface MergedCellRange {
  /** Range reference (e.g., "A1:C1") */
  ref: string;

  /** Start row (1-based) */
  startRow: number;

  /** Start column (1-based) */
  startCol: number;

  /** End row (1-based) */
  endRow: number;

  /** End column (1-based) */
  endCol: number;

  /** Value from the top-left cell */
  value: string | number | boolean | Date | null;

  /** Column span */
  colSpan: number;

  /** Row span */
  rowSpan: number;
}

/**
 * JSON representation of table with nested headers
 */
export type TableJson = Record<string, any>;

/**
 * Image embedded in sheet
 */
export interface SheetImage {
  /** Image file name */
  fileName: string;

  /** Binary content of the image */
  content: Buffer;

  /** Content type/MIME type */
  contentType: string;

  /** Position in sheet (row, column where image starts) */
  position?: {
    row: number;
    col: number;
  };
}
