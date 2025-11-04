import JSZip from 'jszip';
import { Slide, SlideImage, DiagramData, PptxParseResult } from './types';

/**
 * Parse a PPTX file and extract slides with their content, images, and diagram data
 * @param pptxBuffer - Buffer containing the PPTX file content
 * @returns Promise resolving to the parsed presentation data
 */
export async function parsePptx(pptxBuffer: Buffer): Promise<PptxParseResult> {
  const zip = await JSZip.loadAsync(pptxBuffer);
  const slides: Slide[] = [];

  // Find all slide files
  const slideFiles = Object.keys(zip.files)
    .filter(path => path.match(/^ppt\/slides\/slide\d+\.xml$/))
    .sort((a, b) => {
      const numA = parseInt(a.match(/slide(\d+)\.xml$/)?.[1] || '0');
      const numB = parseInt(b.match(/slide(\d+)\.xml$/)?.[1] || '0');
      return numA - numB;
    });

  // Load content types to map extensions to MIME types
  const contentTypes = await loadContentTypes(zip);

  // Process each slide
  for (let i = 0; i < slideFiles.length; i++) {
    const slidePath = slideFiles[i];
    const slideNumber = i + 1;

    // Get slide XML
    const slideXml = await zip.file(slidePath)?.async('text');
    if (!slideXml) continue;

    // Get slide relationships
    const relsPath = slidePath.replace('ppt/slides/', 'ppt/slides/_rels/').replace('.xml', '.xml.rels');
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

  return { slides };
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
    if (target.includes('../media/') || target.match(/\.(png|jpg|jpeg|gif|bmp|svg|tiff?)$/i)) {
      // Resolve relative path
      const imagePath = resolveRelativePath(slideDir, target);

      const imageFile = zip.file(imagePath);
      if (!imageFile) continue;

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
    if (target.includes('../diagrams/') && target.match(/data\d*\.xml$/i)) {
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

/**
 * Resolve a relative path from a base directory
 */
function resolveRelativePath(baseDir: string, relativePath: string): string {
  const parts = baseDir.split('/');
  const relParts = relativePath.split('/');

  for (const part of relParts) {
    if (part === '..') {
      parts.pop();
    } else if (part !== '.') {
      parts.push(part);
    }
  }

  return parts.join('/');
}
