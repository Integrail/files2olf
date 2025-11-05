import JSZip from 'jszip';

/**
 * Parse relationships from a ZIP file at specified path
 * @param zip - JSZip instance
 * @param relsPath - Path to relationships file
 * @returns Map of relationship ID to target path
 */
export async function parseRelationshipsFromPath(
  zip: JSZip,
  relsPath: string
): Promise<Map<string, string>> {
  const relsFile = zip.file(relsPath);
  if (!relsFile) return new Map();

  return parseRelationshipsFromFile(relsFile);
}

/**
 * Parse relationships from a JSZip file object
 * @param relsFile - JSZip file object containing relationships XML
 * @returns Map of relationship ID to target path
 */
export async function parseRelationshipsFromFile(
  relsFile: JSZip.JSZipObject
): Promise<Map<string, string>> {
  const relationships = new Map<string, string>();

  const relsXml = await relsFile.async('text');

  // Extract relationship mappings: rId -> Target path
  const relMatches = relsXml.matchAll(/<Relationship[^>]+Id="([^"]+)"[^>]+Target="([^"]+)"/g);
  for (const match of relMatches) {
    if (match[1] && match[2]) {
      relationships.set(match[1], match[2]);
    }
  }

  return relationships;
}
