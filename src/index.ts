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
