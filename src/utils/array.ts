/**
 * Utility function to ensure a value is always an array
 * @param value - A value that may be undefined, a single item, or an array
 * @returns An array containing the value(s)
 */
export function ensureArray<T>(value: T | T[] | undefined): T[] {
  if (Array.isArray(value)) return value;
  if (value !== undefined && value !== null) return [value];
  return [];
}
