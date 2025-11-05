import JSZip from 'jszip';
import { DocumentImage } from './docxTypes';
import { ensureArray } from './utils/array';

// Maximum image size: 50MB
const MAX_IMAGE_SIZE = 50 * 1024 * 1024;

/**
 * Extract images from document elements
 */
export async function extractImages(
  elements: any[],
  zip: JSZip,
  relationships: Map<string, string>
): Promise<DocumentImage[]> {
  const images: DocumentImage[] = [];
  const processedRIds = new Set<string>();

  for (const element of elements) {
    if (element.type === 'paragraph') {
      const paraImages = await extractImagesFromParagraph(element.data, zip, relationships, processedRIds);
      images.push(...paraImages);
    }
  }

  return images;
}

/**
 * Extract images from a paragraph
 */
async function extractImagesFromParagraph(
  para: any,
  zip: JSZip,
  relationships: Map<string, string>,
  processedRIds: Set<string>
): Promise<DocumentImage[]> {
  const images: DocumentImage[] = [];

  // Look for drawing elements in runs
  const runs = ensureArray(para['w:r']);

  for (const run of runs) {
    const drawing = run['w:drawing'];
    if (!drawing) continue;

    // Extract image from drawing
    const imageRIds = extractImageRIdsFromDrawing(drawing);

    for (const rId of imageRIds) {
      // Wrap each image extraction in try-catch for robustness
      try {
        // Skip if already processed
        if (processedRIds.has(rId)) continue;
        processedRIds.add(rId);

        // Get image path from relationships
        const imagePath = relationships.get(rId);
        if (!imagePath) {
          console.warn(`Warning: Image relationship ${rId} not found`);
          continue;
        }

        // Resolve path (relationships use relative paths)
        const fullPath = `word/${imagePath}`;

        // Extract image from ZIP
        const imageFile = zip.file(fullPath);
        if (!imageFile) {
          console.warn(`Warning: Image file not found: ${fullPath}`);
          continue;
        }

        const content = await imageFile.async('nodebuffer');

        // Validate image size
        if (content.length > MAX_IMAGE_SIZE) {
          console.warn(
            `Warning: Image ${fullPath} size ${content.length} bytes exceeds maximum ${MAX_IMAGE_SIZE} bytes, skipping`
          );
          continue;
        }

        const fileName = imagePath.split('/').pop() || `image_${rId}`;
        const extension = fileName.split('.').pop()?.toLowerCase() || 'png';
        const contentType = getContentType(extension);

        // Try to get description from drawing
        const description = extractImageDescription(drawing);

        images.push({
          rId,
          path: fullPath,
          fileName,
          content,
          contentType,
          description
        });
      } catch (error) {
        // Log error but continue processing other images - graceful degradation
        console.error(`Error extracting image ${rId}:`, error instanceof Error ? error.message : String(error));
      }
    }
  }

  return images;
}

/**
 * Extract image relationship IDs from drawing element
 */
function extractImageRIdsFromDrawing(drawing: any): string[] {
  const rIds: string[] = [];

  // Drawing can be inline or anchor
  const inline = drawing['wp:inline'];
  const anchor = drawing['wp:anchor'];

  const drawingElement = inline || anchor;
  if (!drawingElement) return rIds;

  // Navigate to blip element: graphic -> graphicData -> pic -> blipFill -> blip
  const graphic = drawingElement['a:graphic'];
  if (!graphic) return rIds;

  const graphicData = graphic['a:graphicData'];
  if (!graphicData) return rIds;

  const pic = graphicData['pic:pic'];
  if (!pic) return rIds;

  const blipFill = pic['pic:blipFill'];
  if (!blipFill) return rIds;

  const blip = blipFill['a:blip'];
  if (!blip) return rIds;

  // Get r:embed attribute
  const rEmbed = blip['@_r:embed'];
  if (rEmbed) {
    rIds.push(rEmbed);
  }

  return rIds;
}

/**
 * Extract image description/alt text from drawing
 */
function extractImageDescription(drawing: any): string | undefined {
  const inline = drawing['wp:inline'];
  const anchor = drawing['wp:anchor'];

  const drawingElement = inline || anchor;
  if (!drawingElement) return undefined;

  const docPr = drawingElement['wp:docPr'];
  if (docPr && docPr['@_descr']) {
    return docPr['@_descr'];
  }

  return undefined;
}

/**
 * Get content type from file extension
 */
function getContentType(extension: string): string {
  const ext = extension.toLowerCase();
  const contentTypes: Record<string, string> = {
    'png': 'image/png',
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg',
    'gif': 'image/gif',
    'bmp': 'image/bmp',
    'svg': 'image/svg+xml',
    'tiff': 'image/tiff',
    'tif': 'image/tiff',
    'webp': 'image/webp',
    'emf': 'image/x-emf',
    'wmf': 'image/x-wmf'
  };

  return contentTypes[ext] || `image/${ext}`;
}
