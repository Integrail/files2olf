import { XMLParser } from 'fast-xml-parser';

interface ParsedText {
  text: string;
  isTitle: boolean;
  isSubtitle: boolean;
  listLevel?: number;
  isNumbered?: boolean;
}

/**
 * Convert a slide XML to markdown format
 * @param slideXml - The XML content of the slide
 * @returns Markdown representation of the slide
 */
export function convertSlideToMarkdown(slideXml: string): string {
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    textNodeName: '#text',
    ignoreDeclaration: true,
    parseAttributeValue: false
  });

  const parsed = parser.parse(slideXml);
  const textElements: ParsedText[] = [];

  // Navigate through the slide structure
  const slide = parsed['p:sld'];
  if (!slide) return '';

  const cSld = slide['p:cSld'];
  if (!cSld) return '';

  const spTree = cSld['p:spTree'];
  if (!spTree) return '';

  // Process all shapes in the slide
  const shapes = Array.isArray(spTree['p:sp']) ? spTree['p:sp'] : (spTree['p:sp'] ? [spTree['p:sp']] : []);

  for (const shape of shapes) {
    const shapeTexts = extractTextFromShape(shape);
    if (shapeTexts.length > 0) {
      textElements.push(...shapeTexts);
    }
  }

  // Process graphic frames (for tables, diagrams)
  const graphicFrames = Array.isArray(spTree['p:graphicFrame'])
    ? spTree['p:graphicFrame']
    : (spTree['p:graphicFrame'] ? [spTree['p:graphicFrame']] : []);

  for (const frame of graphicFrames) {
    const text = extractTextFromGraphicFrame(frame);
    if (text) {
      textElements.push({ text, isTitle: false, isSubtitle: false });
    }
  }

  // Convert to markdown
  return convertToMarkdown(textElements);
}

/**
 * Extract text content from a shape
 */
function extractTextFromShape(shape: any): ParsedText[] {
  const results: ParsedText[] = [];

  // Check if this is a title, subtitle, or content placeholder
  const nvSpPr = shape['p:nvSpPr'];
  let isTitle = false;
  let isSubtitle = false;
  let isContentPlaceholder = false;

  if (nvSpPr && nvSpPr['p:nvPr']) {
    const ph = nvSpPr['p:nvPr']['p:ph'];
    if (ph) {
      const phType = ph['@_type'];
      isTitle = phType === 'title' || phType === 'ctrTitle';
      isSubtitle = phType === 'subTitle';
      // Content placeholders have no type or are explicitly marked as 'body' or 'obj'
      isContentPlaceholder = !phType || phType === 'body' || phType === 'obj';
    }
  }

  // Extract text from txBody
  const txBody = shape['p:txBody'];
  if (!txBody) return results;

  // Get all paragraphs
  const paragraphs = Array.isArray(txBody['a:p'])
    ? txBody['a:p']
    : (txBody['a:p'] ? [txBody['a:p']] : []);

  for (const para of paragraphs) {
    const paraText = extractParagraphText(para, isContentPlaceholder);
    if (paraText.text.trim()) {
      results.push({
        text: paraText.text,
        isTitle,
        isSubtitle,
        listLevel: paraText.listLevel,
        isNumbered: paraText.isNumbered
      });
    }
  }

  return results;
}

/**
 * Extract text from a paragraph with list detection
 */
function extractParagraphText(para: any, isContentPlaceholder: boolean = false): { text: string; listLevel?: number; isNumbered?: boolean } {
  let text = '';
  let listLevel: number | undefined;
  let isNumbered: boolean | undefined;

  // Check for list properties
  const pPr = para['a:pPr'];
  if (pPr) {
    // Check list level
    if (pPr['@_lvl'] !== undefined) {
      listLevel = parseInt(pPr['@_lvl']) || 0;
    }

    // Check if it's a numbered list
    if (pPr['a:buAutoNum']) {
      isNumbered = true;
      listLevel = listLevel ?? 0; // Default to level 0 if not set
    } else if (pPr['a:buFont'] || pPr['a:buChar'] || pPr['a:buBlip']) {
      isNumbered = false; // It's a bullet list
      listLevel = listLevel ?? 0; // Default to level 0 if not set
    } else if (pPr['a:buNone']) {
      // Explicitly no bullet - don't treat as list
      isNumbered = undefined;
      listLevel = undefined;
    }
  }

  // Extract text runs
  const runs = Array.isArray(para['a:r'])
    ? para['a:r']
    : (para['a:r'] ? [para['a:r']] : []);

  for (const run of runs) {
    const t = run['a:t'];
    if (t !== undefined && t !== null) {
      // Handle string, number, boolean, or object with #text property
      if (typeof t === 'string') {
        text += t;
      } else if (typeof t === 'number' || typeof t === 'boolean') {
        text += String(t);
      } else if (typeof t === 'object' && t['#text'] !== undefined) {
        text += String(t['#text']);
      }
    }
  }

  // For content placeholders without explicit properties, default to bullet list
  if (isContentPlaceholder && listLevel === undefined && isNumbered === undefined && !pPr?.['a:buNone']) {
    // Content placeholders default to bullet lists at level 0
    listLevel = 0;
    isNumbered = false;
  }

  // Also check for direct text nodes
  if (para['a:t'] !== undefined && para['a:t'] !== null) {
    const t = para['a:t'];
    if (typeof t === 'string') {
      text += t;
    } else if (typeof t === 'number' || typeof t === 'boolean') {
      text += String(t);
    } else if (typeof t === 'object' && t['#text'] !== undefined) {
      text += String(t['#text']);
    }
  }

  return { text, listLevel, isNumbered };
}

/**
 * Extract text from graphic frames (tables, diagrams)
 */
function extractTextFromGraphicFrame(frame: any): string | null {
  const graphic = frame['a:graphic'];
  if (!graphic) return null;

  const graphicData = graphic['a:graphicData'];
  if (!graphicData) return null;

  const uri = graphicData['@_uri'];

  // Handle tables
  if (uri && uri.includes('table')) {
    return extractTableText(graphicData);
  }

  // For diagrams, we'll need the separate data XML file
  // which should be handled separately via the diagram data
  // For now, just note that a diagram exists
  if (uri && uri.includes('diagram')) {
    return '[Diagram]';
  }

  return null;
}

/**
 * Extract text from table structure
 */
function extractTableText(graphicData: any): string | null {
  const tbl = graphicData['a:tbl'];
  if (!tbl) return null;

  const rows: string[][] = [];

  // Get table rows
  const trs = Array.isArray(tbl['a:tr'])
    ? tbl['a:tr']
    : (tbl['a:tr'] ? [tbl['a:tr']] : []);

  for (const tr of trs) {
    const row: string[] = [];

    // Get cells in row
    const tcs = Array.isArray(tr['a:tc'])
      ? tr['a:tc']
      : (tr['a:tc'] ? [tr['a:tc']] : []);

    for (const tc of tcs) {
      const txBody = tc['a:txBody'];
      if (txBody) {
        const paragraphs = Array.isArray(txBody['a:p'])
          ? txBody['a:p']
          : (txBody['a:p'] ? [txBody['a:p']] : []);

        const cellText = paragraphs
          .map(p => extractParagraphText(p).text)
          .filter(t => t.trim())
          .join(' ');

        row.push(cellText || '');
      } else {
        row.push('');
      }
    }

    if (row.length > 0) {
      rows.push(row);
    }
  }

  if (rows.length === 0) return null;

  // Convert to markdown table
  return convertTableToMarkdown(rows);
}

/**
 * Convert table rows to markdown format
 */
function convertTableToMarkdown(rows: string[][]): string {
  if (rows.length === 0) return '';

  const colCount = Math.max(...rows.map(r => r.length));
  const lines: string[] = [];

  // Add header row
  if (rows.length > 0) {
    const header = rows[0].map(cell => cell || ' ').join(' | ');
    lines.push('| ' + header + ' |');
    lines.push('|' + ' --- |'.repeat(colCount));
  }

  // Add data rows
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i].map(cell => cell || ' ').join(' | ');
    lines.push('| ' + row + ' |');
  }

  return lines.join('\n');
}

/**
 * Convert parsed text elements to markdown
 */
function convertToMarkdown(elements: ParsedText[]): string {
  const lines: string[] = [];
  let currentListLevel = -1;

  for (const element of elements) {
    const text = element.text.trim();
    if (!text) continue;

    // Handle titles
    if (element.isTitle) {
      lines.push(`# ${text}`);
      lines.push('');
      currentListLevel = -1;
    }
    // Handle subtitles
    else if (element.isSubtitle) {
      lines.push(`## ${text}`);
      lines.push('');
      currentListLevel = -1;
    }
    // Handle lists
    else if (element.listLevel !== undefined) {
      const indent = '  '.repeat(element.listLevel);
      const marker = element.isNumbered ? '1. ' : '- ';
      lines.push(`${indent}${marker}${text}`);
      currentListLevel = element.listLevel;
    }
    // Handle regular text
    else {
      if (currentListLevel >= 0) {
        lines.push(''); // Add spacing after lists
      }
      lines.push(text);
      lines.push('');
      currentListLevel = -1;
    }
  }

  return lines.join('\n').trim();
}

/**
 * Extract text from diagram data XML
 * @param diagramXml - The XML content of the diagram data file
 * @returns Text content from the diagram
 */
export function extractDiagramText(diagramXml: string): string {
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    textNodeName: '#text',
    ignoreDeclaration: true
  });

  try {
    const parsed = parser.parse(diagramXml);
    const textElements: string[] = [];

    // Navigate through diagram structure to find text
    // Diagrams typically have text in various locations depending on type
    const extractText = (obj: any): void => {
      if (!obj || typeof obj !== 'object') return;

      // Look for text nodes
      if (obj['dgm:t'] || obj['t']) {
        const t = obj['dgm:t'] || obj['t'];
        const text = typeof t === 'string' ? t : t['#text'];
        if (text && text.trim()) {
          textElements.push(text.trim());
        }
      }

      // Recursively search all properties
      for (const key in obj) {
        if (obj.hasOwnProperty(key)) {
          const value = obj[key];
          if (Array.isArray(value)) {
            value.forEach(item => extractText(item));
          } else if (typeof value === 'object') {
            extractText(value);
          }
        }
      }
    };

    extractText(parsed);

    return textElements.join('\n');
  } catch (error) {
    console.error('Error parsing diagram XML:', error);
    return '';
  }
}
