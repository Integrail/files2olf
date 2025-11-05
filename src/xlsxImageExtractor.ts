import ExcelJS from 'exceljs';
import { SheetImage } from './xlsxTypes';

/**
 * Extract images from a worksheet
 */
export function extractImages(
  worksheet: ExcelJS.Worksheet,
  workbook: ExcelJS.Workbook
): SheetImage[] {
  const images: SheetImage[] = [];

  // ExcelJS stores images in worksheet.getImages()
  if (worksheet.getImages) {
    const imageRefs = worksheet.getImages();

    for (const imageRef of imageRefs) {
      try {
        // Get image data from workbook media
        const workbookModel = workbook.model as any;
        const image = workbookModel.media?.find((media: any) => media.name === imageRef.imageId);

        if (image && image.buffer) {
          // Determine content type from extension
          const extension = image.extension || 'png';
          const contentType = getContentType(extension);

          // Extract position if available
          const position = imageRef.range
            ? {
                row: imageRef.range.tl?.nativeRow || imageRef.range.tl?.row || 0,
                col: imageRef.range.tl?.nativeCol || imageRef.range.tl?.col || 0
              }
            : undefined;

          // Ensure buffer is a Node.js Buffer
          const buffer = Buffer.isBuffer(image.buffer)
            ? image.buffer
            : Buffer.from(image.buffer as ArrayBuffer);

          images.push({
            fileName: `image${imageRef.imageId}.${extension}`,
            content: buffer,
            contentType,
            position
          });
        }
      } catch (error) {
        console.warn(`Warning: Failed to extract image ${imageRef.imageId}:`, error);
      }
    }
  }

  return images;
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
    'webp': 'image/webp'
  };

  return contentTypes[ext] || `image/${ext}`;
}
