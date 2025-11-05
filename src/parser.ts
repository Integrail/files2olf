import JSZip from 'jszip';
import { Slide, SlideImage, DiagramData, PptxParseResult, PptxParseOptions } from './types';
import { resolveRelativePath, getRelationshipPath } from './utils/path';
import { REGEX_PATTERNS } from './utils/constants';

/**
 * Parse a PPTX file and extract slides with their content, images, and diagram data
 * @param pptxBuffer - Buffer containing the PPTX file content
 * @param options - Optional parsing options
 * @returns Promise resolving to the parsed presentation data
 * @throws TypeError if pptxBuffer is not a Buffer
 * @throws Error if the PPTX file is invalid or cannot be parsed
 */
export async function parsePptx(pptxBuffer: Buffer, options?: PptxParseOptions): Promise<PptxParseResult> {
  // Input validation
  if (!Buffer.isBuffer(pptxBuffer)) {
    throw new TypeError('pptxBuffer must be a Buffer');
  }
  if (pptxBuffer.length === 0) {
    throw new Error('pptxBuffer is empty');
  }

  try {
    const zip = await JSZip.loadAsync(pptxBuffer);

  // Find all slide files
  const slideFiles = Object.keys(zip.files)
    .filter(path => REGEX_PATTERNS.SLIDE_FILE.test(path))
    .sort((a, b) => {
      const numA = parseInt(a.match(REGEX_PATTERNS.SLIDE_NUMBER)?.[1] || '0');
      const numB = parseInt(b.match(REGEX_PATTERNS.SLIDE_NUMBER)?.[1] || '0');
      return numA - numB;
    });

  // Load content types to map extensions to MIME types
  const contentTypes = await loadContentTypes(zip);

    // Process slides either sequentially or in parallel based on options
    const slides = options?.parallel
      ? await processSlidesParallel(zip, slideFiles, contentTypes)
      : await processSlidesSequential(zip, slideFiles, contentTypes);

    return { slides };
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    throw new Error(`Failed to parse PPTX file: ${errorMessage}`);
  }
}

/**
 * Process slides sequentially (default behavior)
 */
async function processSlidesSequential(
  zip: JSZip,
  slideFiles: string[],
  contentTypes: Map<string, string>
): Promise<Slide[]> {
  const slides: Slide[] = [];

  for (let i = 0; i < slideFiles.length; i++) {
    const slidePath = slideFiles[i];
    const slideNumber = i + 1;

    // Get slide XML
    const slideFile = zip.file(slidePath);
    if (!slideFile) {
      console.warn(`Warning: Slide file not found: ${slidePath}`);
      continue;
    }
    const slideXml = await slideFile.async('text');
    if (!slideXml) {
      console.warn(`Warning: Slide XML is empty: ${slidePath}`);
      continue;
    }

    // Get slide relationships
    const relsPath = getRelationshipPath(slidePath);
    const relationships = await parseRelationships(zip, relsPath);

    // Extract images
    const images = await extractImages(zip, relationships, slidePath, contentTypes);

    // Extract diagram data
    const diagrams = await extractDiagrams(zip, relationships, slidePath);

    slides.push({
      slideNumber,
      xml: slideXml,
      images,
      diagrams
    });
  }

  return slides;
}

/**
 * Process slides in parallel for better performance
 */
async function processSlidesParallel(
  zip: JSZip,
  slideFiles: string[],
  contentTypes: Map<string, string>
): Promise<Slide[]> {
  const slidePromises = slideFiles.map(async (slidePath, i) => {
    const slideNumber = i + 1;

    // Get slide XML
    const slideFile = zip.file(slidePath);
    if (!slideFile) {
      console.warn(`Warning: Slide file not found: ${slidePath}`);
      return null;
    }
    const slideXml = await slideFile.async('text');
    if (!slideXml) {
      console.warn(`Warning: Slide XML is empty: ${slidePath}`);
      return null;
    }

    // Get slide relationships
    const relsPath = getRelationshipPath(slidePath);
    const relationships = await parseRelationships(zip, relsPath);

    // Extract images and diagrams in parallel
    const [images, diagrams] = await Promise.all([
      extractImages(zip, relationships, slidePath, contentTypes),
      extractDiagrams(zip, relationships, slidePath)
    ]);

    return {
      slideNumber,
      xml: slideXml,
      images,
      diagrams
    };
  });

  const slideResults = await Promise.all(slidePromises);
  return slideResults.filter((slide): slide is Slide => slide !== null);
}

/**
 * Load content types mapping from [Content_Types].xml
 */
async function loadContentTypes(zip: JSZip): Promise<Map<string, string>> {
  const contentTypesMap = new Map<string, string>();

  const contentTypesFile = zip.file('[Content_Types].xml');
  if (!contentTypesFile) return contentTypesMap;

  const contentTypesXml = await contentTypesFile.async('text');

  // Parse extension to content type mappings
  const defaultMatches = contentTypesXml.matchAll(/\<Default[^>]+Extension="([^"]+)"[^>]+ContentType="([^"]+)"/g);
  for (const match of defaultMatches) {
    contentTypesMap.set(match[1], match[2]);
  }

  return contentTypesMap;
}

/**
 * Parse relationship file to get references to images, diagrams, etc.
 */
async function parseRelationships(zip: JSZip, relsPath: string): Promise<Map<string, string>> {
  const relationships = new Map<string, string>();

  const relsFile = zip.file(relsPath);
  if (!relsFile) return relationships;

  const relsXml = await relsFile.async('text');

  // Extract relationship mappings: rId -> Target path
  const relMatches = relsXml.matchAll(/\<Relationship[^>]+Id="([^"]+)"[^>]+Target="([^"]+)"/g);
  for (const match of relMatches) {
    relationships.set(match[1], match[2]);
  }

  return relationships;
}

/**
 * Extract images referenced in a slide
 */
async function extractImages(
  zip: JSZip,
  relationships: Map<string, string>,
  slidePath: string,
  contentTypes: Map<string, string>
): Promise<SlideImage[]> {
  const images: SlideImage[] = [];
  const slideDir = slidePath.substring(0, slidePath.lastIndexOf('/'));

  for (const [rId, target] of relationships.entries()) {
    // Check if this is an image reference (typically in ../media/ folder)
    if (target.includes('../media/') || REGEX_PATTERNS.IMAGE_EXTENSION.test(target)) {
      // Resolve relative path
      const imagePath = resolveRelativePath(slideDir, target);

      const imageFile = zip.file(imagePath);
      if (!imageFile) {
        console.warn(`Warning: Image file not found: ${imagePath} (referenced as ${rId} in ${slidePath})`);
        continue;
      }

      const content = await imageFile.async('nodebuffer');
      const fileName = imagePath.split('/').pop() || '';
      const extension = fileName.split('.').pop()?.toLowerCase() || '';
      const contentType = contentTypes.get(extension) || `image/${extension}`;

      images.push({
        rId,
        path: imagePath,
        fileName,
        content,
        contentType
      });
    }
  }

  return images;
}

/**
 * Extract diagram data files referenced in a slide
 */
async function extractDiagrams(
  zip: JSZip,
  relationships: Map<string, string>,
  slidePath: string
): Promise<DiagramData[]> {
  const diagrams: DiagramData[] = [];
  const slideDir = slidePath.substring(0, slidePath.lastIndexOf('/'));

  for (const [rId, target] of relationships.entries()) {
    // Check if this is a diagram data file reference
    if (target.includes('../diagrams/') && REGEX_PATTERNS.DIAGRAM_DATA.test(target)) {
      // Resolve relative path to get the diagram data file
      const dataPath = resolveRelativePath(slideDir, target);

      // Read the diagram data file directly
      const dataFile = zip.file(dataPath);
      if (dataFile) {
        const xml = await dataFile.async('text');
        diagrams.push({
          rId,
          path: dataPath,
          xml
        });
      }
    }
  }

  return diagrams;
}
