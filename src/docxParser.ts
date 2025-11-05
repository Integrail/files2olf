import JSZip from 'jszip';
import { Page, DocxParseOptions, DocxParseResult } from './docxTypes';
import { xmlParser } from './utils/xml';
import { extractContent } from './docxContentExtractor';
import { extractTables } from './docxTableExtractor';
import { extractFormulas } from './docxFormulaExtractor';
import { extractImages } from './docxImageExtractor';
import { convertPageToMarkdown } from './docxMarkdownConverter';
import { cleanupZip } from './utils/zipCleanup';
import { parseRelationshipsFromFile } from './utils/relationships';

// Maximum file size: 100MB
const MAX_DOCX_FILE_SIZE = 100 * 1024 * 1024;
// Maximum number of pages
const MAX_PAGES = 1000;

/**
 * Parse a DOCX file and extract pages with their content, tables, formulas, and images
 * @param docxBuffer - Buffer containing the DOCX file content
 * @param options - Optional parsing options
 * @returns Promise resolving to the parsed document data
 * @throws TypeError if docxBuffer is not a Buffer
 * @throws Error if the DOCX file is invalid or cannot be parsed
 */
export async function parseDocx(docxBuffer: Buffer, options?: DocxParseOptions): Promise<DocxParseResult> {
  // Input validation
  if (!Buffer.isBuffer(docxBuffer)) {
    throw new TypeError('docxBuffer must be a Buffer');
  }
  if (docxBuffer.length === 0) {
    throw new Error('docxBuffer is empty');
  }
  if (docxBuffer.length > MAX_DOCX_FILE_SIZE) {
    throw new Error(
      `File size ${docxBuffer.length} bytes exceeds maximum ${MAX_DOCX_FILE_SIZE} bytes (100MB)`
    );
  }

  let zip: JSZip | undefined;
  let relationships: Map<string, string> | undefined;

  try {
    zip = await JSZip.loadAsync(docxBuffer);

    // Parse main document XML
    const documentFile = zip.file('word/document.xml');
    if (!documentFile) {
      throw new Error('Invalid DOCX file: word/document.xml not found');
    }
    const documentXml = await documentFile.async('text');

    // Parse relationships
    const relsFile = zip.file('word/_rels/document.xml.rels');
    relationships = relsFile ? await parseRelationshipsFromFile(relsFile) : new Map();

    // Parse document and split into pages
    const pages = await parsePages(documentXml, zip, relationships, options);

    // Validate page count
    if (pages.length > MAX_PAGES) {
      throw new Error(
        `Document has ${pages.length} pages, maximum is ${MAX_PAGES}`
      );
    }

    return { pages };
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    throw new Error(`Failed to parse DOCX file: ${errorMessage}`);
  } finally {
    // CRITICAL: Clean up JSZip to prevent memory leaks
    cleanupZip(zip);
    zip = undefined;

    // Clear relationships map
    if (relationships) {
      relationships.clear();
      relationships = undefined;
    }
  }
}


/**
 * Parse document XML and split into pages
 */
async function parsePages(
  documentXml: string,
  zip: JSZip,
  relationships: Map<string, string>,
  options?: DocxParseOptions
): Promise<Page[]> {
  const parsed = xmlParser.parse(documentXml);

  const document = parsed['w:document'];
  if (!document) {
    throw new Error('Invalid document structure');
  }

  const body = document['w:body'];
  if (!body) {
    throw new Error('Document body not found');
  }

  // Get all body children (paragraphs, tables, etc.)
  const bodyElements = getBodyElements(body);

  // Split into pages by lastRenderedPageBreak markers
  const pages = splitIntoPages(bodyElements);

  // Process each page
  const result: Page[] = [];

  for (let i = 0; i < pages.length; i++) {
    const pageElements = pages[i];
    const page = await processPage(pageElements, i + 1, zip, relationships, options);
    result.push(page);
  }

  return result;
}

/**
 * Get all elements from document body
 */
function getBodyElements(body: any): any[] {
  const elements: any[] = [];

  // Paragraphs
  const paragraphs = body['w:p'];
  if (paragraphs) {
    if (Array.isArray(paragraphs)) {
      elements.push(...paragraphs.map(p => ({ type: 'paragraph', data: p })));
    } else {
      elements.push({ type: 'paragraph', data: paragraphs });
    }
  }

  // Tables
  const tables = body['w:tbl'];
  if (tables) {
    if (Array.isArray(tables)) {
      elements.push(...tables.map(t => ({ type: 'table', data: t })));
    } else {
      elements.push({ type: 'table', data: tables });
    }
  }

  // Sort by document order (they should be interleaved correctly already)
  return elements;
}

/**
 * Split elements into pages based on page break markers
 */
function splitIntoPages(elements: any[]): any[][] {
  const pages: any[][] = [];
  let currentPage: any[] = [];

  for (const element of elements) {
    // Check if this element contains a page break
    if (element.type === 'paragraph' && containsPageBreak(element.data)) {
      // Split the paragraph at the break
      const { before, after } = splitParagraphAtPageBreak(element.data);

      if (before) {
        currentPage.push({ type: 'paragraph', data: before });
      }

      // Start new page
      if (currentPage.length > 0) {
        pages.push(currentPage);
      }
      currentPage = [];

      if (after) {
        currentPage.push({ type: 'paragraph', data: after });
      }
    } else {
      currentPage.push(element);
    }
  }

  // Add last page
  if (currentPage.length > 0) {
    pages.push(currentPage);
  }

  // If no pages (no content), return one empty page
  if (pages.length === 0) {
    pages.push([]);
  }

  return pages;
}

/**
 * Check if paragraph contains a page break marker
 */
function containsPageBreak(paragraph: any): boolean {
  const runs = paragraph['w:r'];
  if (!runs) return false;

  const runsArray = Array.isArray(runs) ? runs : [runs];

  return runsArray.some(run => run['w:lastRenderedPageBreak'] !== undefined);
}

/**
 * Split paragraph at page break into before and after
 */
function splitParagraphAtPageBreak(paragraph: any): { before: any | null; after: any | null } {
  const runs = paragraph['w:r'];
  if (!runs) return { before: paragraph, after: null };

  const runsArray = Array.isArray(runs) ? runs : [runs];

  // Find the index of the run with page break
  const breakIndex = runsArray.findIndex(run => run['w:lastRenderedPageBreak'] !== undefined);

  if (breakIndex === -1) {
    return { before: paragraph, after: null };
  }

  // Split runs into before (including break run) and after
  const beforeRuns = runsArray.slice(0, breakIndex + 1);
  const afterRuns = runsArray.slice(breakIndex + 1);

  // Create before paragraph (clone paragraph structure)
  const beforePara = beforeRuns.length > 0
    ? {
        ...paragraph,
        'w:r': beforeRuns.length === 1 ? beforeRuns[0] : beforeRuns
      }
    : null;

  // Create after paragraph (clone paragraph structure, preserve properties)
  const afterPara = afterRuns.length > 0
    ? {
        ...paragraph,
        'w:r': afterRuns.length === 1 ? afterRuns[0] : afterRuns
      }
    : null;

  return { before: beforePara, after: afterPara };
}

/**
 * Process a single page
 */
async function processPage(
  elements: any[],
  pageNumber: number,
  zip: JSZip,
  relationships: Map<string, string>,
  options?: DocxParseOptions
): Promise<Page> {
  // Extract content
  const paragraphs = extractContent(elements);

  // Extract tables
  const tables = extractTables(elements);

  // Extract formulas
  const formulas = options?.includeFormulas !== false ? extractFormulas(elements) : [];

  // Extract images
  const images = options?.includeImages !== false
    ? await extractImages(elements, zip, relationships)
    : [];

  // Convert to markdown
  const markdown = options?.convertToMarkdown !== false
    ? convertPageToMarkdown({ pageNumber, paragraphs, tables, formulas, images })
    : undefined;

  return {
    pageNumber,
    paragraphs,
    tables,
    formulas,
    images,
    markdown
  };
}
