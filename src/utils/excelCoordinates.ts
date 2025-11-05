/**
 * Excel coordinate conversion utilities
 */

/**
 * Convert row and column numbers to Excel cell address (e.g., (1, 1) -> "A1")
 * @param row - Row number (1-based)
 * @param col - Column number (1-based)
 * @returns Excel cell address (e.g., "A1", "Z99", "AA1")
 */
export function coordsToAddress(row: number, col: number): string {
  let colLetter = '';
  let tempCol = col;

  while (tempCol > 0) {
    const remainder = (tempCol - 1) % 26;
    colLetter = String.fromCharCode(65 + remainder) + colLetter;
    tempCol = Math.floor((tempCol - 1) / 26);
  }

  return `${colLetter}${row}`;
}

/**
 * Convert Excel cell address to row and column coordinates
 * @param address - Excel cell address (e.g., "A1", "Z99", "AA1")
 * @returns Object with row and column numbers (1-based)
 * @throws Error if address format is invalid
 */
export function cellAddressToCoords(address: string): { row: number; col: number } {
  const match = address.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Invalid cell address: ${address}`);
  }

  const colLetters = match[1];
  const row = parseInt(match[2]);

  let col = 0;
  for (let i = 0; i < colLetters.length; i++) {
    col = col * 26 + (colLetters.charCodeAt(i) - 64);
  }

  return { row, col };
}

/**
 * Parse a cell range reference (e.g., "A1:C3") to coordinates
 * @param ref - Range reference string
 * @returns Range coordinates (1-based)
 * @throws Error if ref format is invalid
 */
export function parseCellRange(ref: string): {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
} {
  const parts = ref.split(':');
  const start = cellAddressToCoords(parts[0]);
  const end = parts.length > 1 ? cellAddressToCoords(parts[1]) : start;

  return {
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col
  };
}
