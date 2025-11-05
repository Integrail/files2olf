import { xmlParser, extractTextValue } from './utils/xml';
import { ensureArray } from './utils/array';
import { XML_PATHS, PLACEHOLDER_TYPES, GRAPHIC_URIS } from './utils/constants';
import { convertTableToMarkdown as convertToMarkdownTable } from './utils/markdown';
import type {
  PptxShape,
  Paragraph,
  ParagraphResult,
  ParagraphProperties,
  GraphicFrame,
  GraphicData
} from './utils/xmlTypes';

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
 * @throws Error if slideXml is invalid or cannot be parsed
 */
export function convertSlideToMarkdown(slideXml: string): string {
  // Input validation
  if (typeof slideXml !== 'string') {
    throw new TypeError('slideXml must be a string');
  }
  if (!slideXml.trim()) {
    return '';
  }

  try {
    const parsed = xmlParser.parse(slideXml);
    const textElements: ParsedText[] = [];

  // Navigate through the slide structure
  const slide = parsed[XML_PATHS.SLIDE];
  if (!slide) return '';

  const cSld = slide[XML_PATHS.COMMON_SLIDE_DATA];
  if (!cSld) return '';

  const spTree = cSld[XML_PATHS.SHAPE_TREE];
  if (!spTree) return '';

  // Process all shapes in the slide
  const shapes = ensureArray(spTree[XML_PATHS.SHAPE]);

  for (const shape of shapes) {
    const shapeTexts = extractTextFromShape(shape);
    if (shapeTexts.length > 0) {
      textElements.push(...shapeTexts);
    }
  }

  // Process graphic frames (for tables, diagrams)
  const graphicFrames = ensureArray(spTree[XML_PATHS.GRAPHIC_FRAME]);

  for (const frame of graphicFrames) {
    const text = extractTextFromGraphicFrame(frame);
    if (text) {
      textElements.push({ text, isTitle: false, isSubtitle: false });
    }
  }

    // Convert to markdown
    return convertToMarkdown(textElements);
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    throw new Error(`Failed to convert slide XML to markdown: ${errorMessage}`);
  }
}

/**
 * Extract text content from a shape
 */
function extractTextFromShape(shape: any): ParsedText[] {
  const results: ParsedText[] = [];

  // Check if this is a title, subtitle, or content placeholder
  const nvSpPr = shape[XML_PATHS.NON_VISUAL_SHAPE_PROPS];
  let isTitle = false;
  let isSubtitle = false;
  let isContentPlaceholder = false;

  if (nvSpPr && nvSpPr[XML_PATHS.NON_VISUAL_PROPS]) {
    const ph = nvSpPr[XML_PATHS.NON_VISUAL_PROPS][XML_PATHS.PLACEHOLDER];
    if (ph) {
      const phType = ph[XML_PATHS.ATTR_TYPE];
      isTitle = phType === PLACEHOLDER_TYPES.TITLE || phType === PLACEHOLDER_TYPES.CENTER_TITLE;
      isSubtitle = phType === PLACEHOLDER_TYPES.SUBTITLE;
      // Content placeholders have no type or are explicitly marked as 'body' or 'obj'
      isContentPlaceholder = !phType || phType === PLACEHOLDER_TYPES.BODY || phType === PLACEHOLDER_TYPES.OBJECT;
    }
  }

  // Extract text from txBody
  const txBody = shape[XML_PATHS.TEXT_BODY];
  if (!txBody) return results;

  // Get all paragraphs
  const paragraphs = ensureArray(txBody[XML_PATHS.PARAGRAPH]);

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
 * Detect list properties from paragraph properties
 */
function detectListProperties(pPr: any): { listLevel?: number; isNumbered?: boolean } {
  if (!pPr) return {};

  let listLevel: number | undefined;
  let isNumbered: boolean | undefined;

  // Check list level
  if (pPr[XML_PATHS.ATTR_LEVEL] !== undefined) {
    listLevel = parseInt(pPr[XML_PATHS.ATTR_LEVEL]) || 0;
  }

  // Check if it's a numbered list
  if (pPr[XML_PATHS.BULLET_AUTO_NUM]) {
    isNumbered = true;
    listLevel = listLevel ?? 0;
  } else if (pPr[XML_PATHS.BULLET_FONT] || pPr[XML_PATHS.BULLET_CHAR] || pPr[XML_PATHS.BULLET_BLIP]) {
    isNumbered = false; // It's a bullet list
    listLevel = listLevel ?? 0;
  } else if (pPr[XML_PATHS.BULLET_NONE]) {
    // Explicitly no bullet
    return { listLevel: undefined, isNumbered: undefined };
  }

  return { listLevel, isNumbered };
}

/**
 * Extract text from text runs in a paragraph
 */
function extractTextFromRuns(para: any): string {
  const runs = ensureArray(para[XML_PATHS.TEXT_RUN]);
  let text = '';

  for (const run of runs) {
    text += extractTextValue(run[XML_PATHS.TEXT]);
  }

  // Also check for direct text nodes
  text += extractTextValue(para[XML_PATHS.TEXT]);

  return text;
}

/**
 * Extract text from a paragraph with list detection
 */
function extractParagraphText(para: any, isContentPlaceholder: boolean = false): { text: string; listLevel?: number; isNumbered?: boolean } {
  const pPr = para[XML_PATHS.PARAGRAPH_PROPS];
  let { listLevel, isNumbered } = detectListProperties(pPr);

  // Extract text from runs
  const text = extractTextFromRuns(para);

  // For content placeholders without explicit properties, default to bullet list
  if (isContentPlaceholder && listLevel === undefined && isNumbered === undefined && !pPr?.[XML_PATHS.BULLET_NONE]) {
    listLevel = 0;
    isNumbered = false;
  }

  return { text, listLevel, isNumbered };
}

/**
 * Extract text from graphic frames (tables, diagrams)
 */
function extractTextFromGraphicFrame(frame: any): string | null {
  const graphic = frame[XML_PATHS.GRAPHIC];
  if (!graphic) return null;

  const graphicData = graphic[XML_PATHS.GRAPHIC_DATA];
  if (!graphicData) return null;

  const uri = graphicData[XML_PATHS.ATTR_URI];

  // Handle tables
  if (uri && uri.includes(GRAPHIC_URIS.TABLE)) {
    return extractTableText(graphicData);
  }

  // For diagrams, we'll need the separate data XML file
  // which should be handled separately via the diagram data
  // For now, just note that a diagram exists
  if (uri && uri.includes(GRAPHIC_URIS.DIAGRAM)) {
    return '[Diagram]';
  }

  return null;
}

/**
 * Extract text from table structure
 */
function extractTableText(graphicData: any): string | null {
  const tbl = graphicData[XML_PATHS.TABLE];
  if (!tbl) return null;

  const rows: string[][] = [];

  // Get table rows
  const trs = ensureArray(tbl[XML_PATHS.TABLE_ROW]);

  for (const tr of trs) {
    const row: string[] = [];

    // Get cells in row
    const tcs = ensureArray(tr[XML_PATHS.TABLE_CELL]);

    for (const tc of tcs) {
      const txBody = tc[XML_PATHS.TEXT_BODY_DRAWING];
      if (txBody) {
        const paragraphs = ensureArray(txBody[XML_PATHS.PARAGRAPH]);

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
  return convertToMarkdownTable(rows);
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
 * @throws Error if diagramXml is invalid or cannot be parsed
 */
export function extractDiagramText(diagramXml: string): string {
  // Input validation
  if (typeof diagramXml !== 'string') {
    throw new TypeError('diagramXml must be a string');
  }
  if (!diagramXml.trim()) {
    return '';
  }

  try {
    const parsed = xmlParser.parse(diagramXml);
    const textElements: string[] = [];

    // Navigate through diagram structure to find text
    // Diagrams typically have text in various locations depending on type
    const MAX_RECURSION_DEPTH = 50;

    const extractText = (obj: any, depth: number = 0): void => {
      if (!obj || typeof obj !== 'object') return;

      // Prevent stack overflow from deeply nested or circular structures
      if (depth > MAX_RECURSION_DEPTH) {
        console.warn(`Maximum diagram recursion depth (${MAX_RECURSION_DEPTH}) reached, truncating extraction`);
        return;
      }

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
            value.forEach(item => extractText(item, depth + 1));
          } else if (typeof value === 'object') {
            extractText(value, depth + 1);
          }
        }
      }
    };

    extractText(parsed, 0);

    return textElements.join('\n');
  } catch (error) {
    console.error('Error parsing diagram XML:', error);
    return '';
  }
}
