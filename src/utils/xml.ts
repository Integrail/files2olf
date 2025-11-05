import { XMLParser } from 'fast-xml-parser';

/**
 * Reusable XML parser instance configured for PowerPoint XML parsing
 * Shared across all parsing operations to avoid recreating parser instances
 */
export const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  ignoreDeclaration: true,
  parseAttributeValue: false
});

/**
 * Utility function to extract text from various XML node types
 * @param node - XML node that may contain text (string, number, boolean, or object with #text)
 * @returns Extracted text as string, or empty string if no text found
 */
export function extractTextValue(node: any): string {
  if (node === undefined || node === null) return '';
  if (typeof node === 'string') return node;
  if (typeof node === 'number' || typeof node === 'boolean') return String(node);
  if (typeof node === 'object' && node['#text'] !== undefined) return String(node['#text']);
  return '';
}
