/**
 * Markdown conversion utilities
 */

/**
 * Convert a 2D array of string data to markdown table format
 * @param rows - 2D array where first row is header, subsequent rows are data
 * @returns Markdown table string
 */
export function convertTableToMarkdown(rows: string[][]): string {
  if (rows.length === 0) return '';

  const lines: string[] = [];

  // Normalize column count
  const colCount = Math.max(...rows.map(r => r.length));

  // Header row
  const header = rows[0].map(cell => cell || ' ').join(' | ');
  lines.push('| ' + header + ' |');
  lines.push('| ' + rows[0].map(() => '---').join(' | ') + ' |');

  // Data rows
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i].map(cell => cell || ' ').join(' | ');
    lines.push('| ' + row + ' |');
  }

  return lines.join('\n');
}
