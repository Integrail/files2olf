/**
 * Type definitions for DOCX parser
 */

/**
 * Options for parsing a DOCX document
 */
export interface DocxParseOptions {
  /**
   * Convert content to markdown format (default: true)
   */
  convertToMarkdown?: boolean;

  /**
   * Include embedded images (default: true)
   */
  includeImages?: boolean;

  /**
   * Include formulas/equations (default: true)
   */
  includeFormulas?: boolean;
}

/**
 * Result of parsing a DOCX document
 */
export interface DocxParseResult {
  /** Array of pages extracted from the document */
  pages: Page[];
}

/**
 * Represents a single page in the document
 */
export interface Page {
  /** The page number (1-indexed) */
  pageNumber: number;

  /** Paragraphs in this page */
  paragraphs: Paragraph[];

  /** Tables in this page */
  tables: Table[];

  /** Math formulas/equations in this page */
  formulas: Formula[];

  /** Images in this page */
  images: DocumentImage[];

  /** Markdown representation of the page */
  markdown?: string;
}

/**
 * Represents a paragraph with text and formatting
 */
export interface Paragraph {
  /** Paragraph text content */
  text: string;

  /** Paragraph style (e.g., "Heading1", "Heading2", "Normal") */
  style?: string;

  /** Whether this is a list item */
  isList?: boolean;

  /** List level (0-indexed) */
  listLevel?: number;

  /** List numbering ID */
  listId?: string;

  /** Text runs with formatting */
  runs: TextRun[];
}

/**
 * Represents a text run with formatting
 */
export interface TextRun {
  /** Text content */
  text: string;

  /** Is bold */
  bold?: boolean;

  /** Is italic */
  italic?: boolean;

  /** Is underlined */
  underline?: boolean;

  /** Font color */
  color?: string;
}

/**
 * Represents a table in the document
 */
export interface Table {
  /** Table rows */
  rows: TableRow[];

  /** Markdown representation of the table */
  markdown: string;
}

/**
 * Represents a table row
 */
export interface TableRow {
  /** Cells in this row */
  cells: TableCell[];
}

/**
 * Represents a table cell
 */
export interface TableCell {
  /** Cell text content */
  text: string;

  /** Cell paragraphs (cells can have multiple paragraphs) */
  paragraphs: Paragraph[];

  /** Column span (for merged cells) */
  gridSpan?: number;

  /** Vertical merge indicator */
  vMerge?: string;
}

/**
 * Represents a mathematical formula/equation
 */
export interface Formula {
  /** OMML (Office Math Markup Language) XML representation */
  omml: string;

  /** Plain text representation of the formula (extracted from OMML) */
  text?: string;
}

/**
 * Image embedded in document
 */
export interface DocumentImage {
  /** Relationship ID */
  rId: string;

  /** Path within the DOCX archive (e.g., "word/media/image1.png") */
  path: string;

  /** Image file name */
  fileName: string;

  /** Binary content of the image */
  content: Buffer;

  /** Content type/MIME type */
  contentType: string;

  /** Image description/alt text */
  description?: string;
}
