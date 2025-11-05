import JSZip from 'jszip';
import { Slide, SlideImage, DiagramData, PptxParseResult, PptxParseOptions } from './types';
import { resolveRelativePath, getRelationshipPath } from './utils/path';
import { REGEX_PATTERNS } from './utils/constants';
import { cleanupZip } from './utils/zipCleanup';
import { parseRelationshipsFromPath } from './utils/relationships';

/**
 * Parse a PPTX file and extract slides with their content, images, and diagram data
 * @param pptxBuffer - Buffer containing the PPTX file content
 * @param options - Optional parsing options
 * @returns Promise resolving to the parsed presentation data
 * @throws TypeError if pptxBuffer is not a Buffer
 * @throws Error if the PPTX file is invalid or cannot be parsed
 */
// Maximum file size: 100MB
const MAX_PPTX_FILE_SIZE = 100 * 1024 * 1024;
// Maximum number of slides
const MAX_SLIDES = 1000;
// Maximum image size: 50MB
const MAX_IMAGE_SIZE = 50 * 1024 * 1024;

export async function parsePptx(pptxBuffer: Buffer, options?: PptxParseOptions): Promise<PptxParseResult> {
  // Input validation
  if (!Buffer.isBuffer(pptxBuffer)) {
    throw new TypeError('pptxBuffer must be a Buffer');
  }
  if (pptxBuffer.length === 0) {
    throw new Error('pptxBuffer is empty');
  }
  if (pptxBuffer.length > MAX_PPTX_FILE_SIZE) {
    throw new Error(
      `File size ${pptxBuffer.length} bytes exceeds maximum ${MAX_PPTX_FILE_SIZE} bytes (100MB)`
    );
  }

  let zip: JSZip | undefined;

  try {
    zip = await JSZip.loadAsync(pptxBuffer);

    // Find all slide files
    const slideFiles = Object.keys(zip.files)
      .filter(path => REGEX_PATTERNS.SLIDE_FILE.test(path))
      .sort((a, b) => {
        const numA = parseInt(a.match(REGEX_PATTERNS.SLIDE_NUMBER)?.[1] || '0');
        const numB = parseInt(b.match(REGEX_PATTERNS.SLIDE_NUMBER)?.[1] || '0');
        return numA - numB;
      });

    // Validate slide count
    if (slideFiles.length > MAX_SLIDES) {
      throw new Error(
        `Presentation has ${slideFiles.length} slides, maximum is ${MAX_SLIDES}`
      );
    }

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
  } finally {
    // CRITICAL: Clean up JSZip to prevent memory leaks
    cleanupZip(zip);
    zip = undefined;
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
    const relationships = await parseRelationshipsFromPath(zip, relsPath);

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
    const relationships = await parseRelationshipsFromPath(zip, relsPath);

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

      // Validate image size
      if (content.length > MAX_IMAGE_SIZE) {
        console.warn(
          `Warning: Image ${imagePath} size ${content.length} bytes exceeds maximum ${MAX_IMAGE_SIZE} bytes, skipping`
        );
        continue;
      }

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
