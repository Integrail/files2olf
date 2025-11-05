// PPTX Parser
export { parsePptx } from './parser';
export { convertSlideToMarkdown, extractDiagramText } from './converter';
export {
  Slide,
  SlideImage,
  DiagramData,
  PptxParseResult,
  PptxParseOptions
} from './types';

// XLSX Parser
export { parseXlsx } from './xlsxParser';

// XLS Parser (Legacy Excel 97-2003)
export { parseXls } from './xlsParser';

// Shared Excel types (used by both XLSX and XLS)
export {
  Sheet,
  Table,
  CellData,
  CellType,
  MergedCellRange,
  SheetImage as XlsxSheetImage,
  TableJson,
  XlsxParseOptions,
  XlsxParseResult
} from './xlsxTypes';

// Format Detection Utilities
export {
  detectOfficeFormat,
  isExcelFile,
  isPowerPointFile,
  type OfficeFormat
} from './utils/formatDetector';
