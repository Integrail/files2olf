import { Table, TableRow, TableCell, Paragraph } from './docxTypes';
import { ensureArray } from './utils/array';
import { extractTextValue } from './utils/xml';
import { convertTableToMarkdown as convertToMarkdownTable } from './utils/markdown';

/**
 * Extract tables from document elements
 */
export function extractTables(elements: any[]): Table[] {
  const tables: Table[] = [];

  for (const element of elements) {
    if (element.type === 'table') {
      const table = extractTable(element.data);
      if (table) {
        tables.push(table);
      }
    }
  }

  return tables;
}

/**
 * Extract a single table
 */
function extractTable(tbl: any): Table | null {
  const rows: TableRow[] = [];

  // Get table rows
  const wRows = ensureArray(tbl['w:tr']);

  for (const wRow of wRows) {
    const row = extractTableRow(wRow);
    if (row) {
      rows.push(row);
    }
  }

  if (rows.length === 0) {
    return null;
  }

  // Convert to markdown
  const stringRows = rows.map(row => row.cells.map(cell => cell.text || ''));
  const markdown = convertToMarkdownTable(stringRows);

  return {
    rows,
    markdown
  };
}

/**
 * Extract a table row
 */
function extractTableRow(tr: any): TableRow | null {
  const cells: TableCell[] = [];

  // Get table cells
  const wCells = ensureArray(tr['w:tc']);

  for (const wCell of wCells) {
    const cell = extractTableCell(wCell);
    if (cell) {
      cells.push(cell);
    }
  }

  if (cells.length === 0) {
    return null;
  }

  return { cells };
}

/**
 * Extract a table cell
 */
function extractTableCell(tc: any): TableCell | null {
  // Get cell properties
  const tcPr = tc['w:tcPr'];
  let gridSpan: number | undefined;
  let vMerge: string | undefined;

  if (tcPr) {
    const gridSpanElement = tcPr['w:gridSpan'];
    if (gridSpanElement && gridSpanElement['@_w:val']) {
      gridSpan = parseInt(gridSpanElement['@_w:val']) || undefined;
    }

    const vMergeElement = tcPr['w:vMerge'];
    if (vMergeElement) {
      vMerge = vMergeElement['@_w:val'] || 'continue';
    }
  }

  // Extract paragraphs in cell
  const paragraphs = ensureArray(tc['w:p']);

  // Extract text from all paragraphs
  const text = paragraphs
    .map((p: any) => extractParagraphText(p))
    .filter((t: string) => t.trim())
    .join('\n');

  // For simplicity, we're not extracting full paragraph structure for cells
  // Just returning the text and basic cell info
  return {
    text,
    paragraphs: [], // Could extract full paragraphs if needed
    gridSpan,
    vMerge
  };
}

/**
 * Extract plain text from a paragraph (for table cells)
 */
function extractParagraphText(para: any): string {
  const runs = ensureArray(para['w:r']);

  let text = '';
  for (const run of runs) {
    text += extractTextValue(run['w:t']) || '';
  }

  // Also check hyperlinks
  const hyperlinks = ensureArray(para['w:hyperlink']);
  for (const hyperlink of hyperlinks) {
    const linkRuns = ensureArray(hyperlink['w:r']);
    for (const run of linkRuns) {
      text += extractTextValue(run['w:t']) || '';
    }
  }

  return text;
}

