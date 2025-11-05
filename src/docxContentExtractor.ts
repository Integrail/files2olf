import { Paragraph, TextRun } from './docxTypes';
import { ensureArray } from './utils/array';
import { extractTextValue } from './utils/xml';

/**
 * Extract paragraphs and text content from document elements
 */
export function extractContent(elements: any[]): Paragraph[] {
  const paragraphs: Paragraph[] = [];

  for (const element of elements) {
    if (element.type === 'paragraph') {
      const paragraph = extractParagraph(element.data);
      if (paragraph) {
        paragraphs.push(paragraph);
      }
    }
  }

  return paragraphs;
}

/**
 * Extract a single paragraph
 */
function extractParagraph(para: any): Paragraph | null {
  // Extract paragraph properties
  const pPr = para['w:pPr'];
  let style: string | undefined;
  let isList = false;
  let listLevel: number | undefined;
  let listId: string | undefined;

  if (pPr) {
    // Get paragraph style (Heading1, Heading2, etc.)
    const pStyle = pPr['w:pStyle'];
    if (pStyle && pStyle['@_w:val']) {
      style = pStyle['@_w:val'];
    }

    // Check if it's a list
    const numPr = pPr['w:numPr'];
    if (numPr) {
      isList = true;
      const ilvl = numPr['w:ilvl'];
      const numId = numPr['w:numId'];

      if (ilvl && ilvl['@_w:val'] !== undefined) {
        listLevel = parseInt(ilvl['@_w:val']) || 0;
      }
      if (numId && numId['@_w:val']) {
        listId = numId['@_w:val'];
      }
    }
  }

  // Check for math formula (skip these, they're handled separately)
  if (para['m:oMathPara'] || para['m:oMath']) {
    return null;
  }

  // Extract text runs
  const runs = extractTextRuns(para);

  // Build paragraph text
  const text = runs.map(r => r.text).join('');

  // Skip empty paragraphs
  if (!text.trim() && runs.length === 0) {
    return null;
  }

  return {
    text,
    style,
    isList,
    listLevel,
    listId,
    runs
  };
}

/**
 * Extract text runs from a paragraph
 */
function extractTextRuns(para: any): TextRun[] {
  const runs: TextRun[] = [];

  const wRuns = ensureArray(para['w:r']);

  for (const run of wRuns) {
    // Skip if it's a page break marker
    if (run['w:lastRenderedPageBreak']) {
      continue;
    }

    // Extract text
    const text = extractTextValue(run['w:t']) || '';
    if (!text) continue;

    // Extract formatting
    const rPr = run['w:rPr'];
    let bold = false;
    let italic = false;
    let underline = false;
    let color: string | undefined;

    if (rPr) {
      bold = rPr['w:b'] !== undefined || rPr['w:bCs'] !== undefined;
      italic = rPr['w:i'] !== undefined || rPr['w:iCs'] !== undefined;
      underline = rPr['w:u'] !== undefined;

      const colorElement = rPr['w:color'];
      if (colorElement && colorElement['@_w:val']) {
        color = colorElement['@_w:val'];
      }
    }

    runs.push({
      text,
      bold: bold || undefined,
      italic: italic || undefined,
      underline: underline || undefined,
      color
    });
  }

  // Also check for hyperlinks
  const hyperlinks = ensureArray(para['w:hyperlink']);
  for (const hyperlink of hyperlinks) {
    const linkRuns = ensureArray(hyperlink['w:r']);
    for (const run of linkRuns) {
      const text = extractTextValue(run['w:t']) || '';
      if (text) {
        runs.push({
          text,
          underline: true // Hyperlinks are typically underlined
        });
      }
    }
  }

  return runs;
}
