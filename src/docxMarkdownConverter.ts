import { Page, Paragraph } from './docxTypes';

/**
 * Convert a page to markdown format
 */
export function convertPageToMarkdown(page: Page): string {
  const lines: string[] = [];

  // Add page header
  lines.push(`# Page ${page.pageNumber}`);
  lines.push('');

  // Convert paragraphs
  for (const para of page.paragraphs) {
    const markdown = convertParagraphToMarkdown(para);
    if (markdown) {
      lines.push(markdown);
      lines.push('');
    }
  }

  // Convert tables
  for (const table of page.tables) {
    lines.push(table.markdown);
    lines.push('');
  }

  // Add formulas
  for (const formula of page.formulas) {
    if (formula.text) {
      lines.push(`**Formula**: ${formula.text}`);
    } else {
      lines.push('**Formula**: (OMML - see raw data)');
    }
    lines.push('');
  }

  // Add image references
  for (const image of page.images) {
    const altText = image.description || image.fileName;
    lines.push(`![${altText}](${image.fileName})`);
    lines.push('');
  }

  return lines.join('\n').trim();
}

/**
 * Convert a paragraph to markdown
 */
function convertParagraphToMarkdown(para: Paragraph): string {
  // Check if it's a heading
  if (para.style) {
    const headingLevel = getHeadingLevel(para.style);
    if (headingLevel > 0) {
      const hashes = '#'.repeat(headingLevel);
      return `${hashes} ${para.text}`;
    }
  }

  // Check if it's a list item
  if (para.isList) {
    const indent = '  '.repeat(para.listLevel || 0);
    // For now, treat all lists as bullet lists
    // Could check listId to determine numbered vs bullet
    return `${indent}- ${para.text}`;
  }

  // Regular paragraph with inline formatting
  if (para.runs && para.runs.length > 0) {
    let formatted = '';

    for (const run of para.runs) {
      let text = run.text;

      // Apply formatting
      if (run.bold && run.italic) {
        text = `***${text}***`;
      } else if (run.bold) {
        text = `**${text}**`;
      } else if (run.italic) {
        text = `*${text}*`;
      }

      formatted += text;
    }

    return formatted;
  }

  // Plain text paragraph
  return para.text;
}

/**
 * Get heading level from style name
 */
function getHeadingLevel(style: string): number {
  const match = style.match(/Heading(\d+)/i);
  if (match) {
    return parseInt(match[1]) || 0;
  }
  return 0;
}
