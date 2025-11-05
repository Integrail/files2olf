/**
 * Resolve a relative path from a base directory
 * @param baseDir - Base directory path
 * @param relativePath - Relative path to resolve
 * @returns Resolved absolute path
 */
export function resolveRelativePath(baseDir: string, relativePath: string): string {
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

/**
 * Get the directory path from a file path
 * @param filePath - Full file path
 * @returns Directory path
 */
export function getDirectoryPath(filePath: string): string {
  const lastSlashIndex = filePath.lastIndexOf('/');
  return lastSlashIndex > -1 ? filePath.substring(0, lastSlashIndex) : '';
}

/**
 * Convert a slide path to its relationship file path
 * @param slidePath - Path to slide XML file
 * @returns Path to corresponding relationship file
 */
export function getRelationshipPath(slidePath: string): string {
  return slidePath
    .replace('ppt/slides/', 'ppt/slides/_rels/')
    .replace('.xml', '.xml.rels');
}
