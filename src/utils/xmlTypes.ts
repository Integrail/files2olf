/**
 * TypeScript interfaces for PowerPoint XML structures
 * These provide type safety when working with parsed XML nodes
 */

/**
 * Base type for any XML node
 */
export interface XmlNode {
  [key: string]: any;
}

/**
 * Text value that can be a string, number, boolean, or object with #text property
 */
export type TextValue = string | number | boolean | { '#text': string | number };

/**
 * PowerPoint shape structure
 */
export interface PptxShape extends XmlNode {
  'p:nvSpPr'?: NonVisualShapeProps;
  'p:txBody'?: TextBody;
}

/**
 * Non-visual shape properties
 */
export interface NonVisualShapeProps extends XmlNode {
  'p:nvPr'?: NonVisualProps;
}

/**
 * Non-visual properties containing placeholder information
 */
export interface NonVisualProps extends XmlNode {
  'p:ph'?: Placeholder;
}

/**
 * Placeholder information
 */
export interface Placeholder extends XmlNode {
  '@_type'?: string;
  '@_idx'?: string;
}

/**
 * Text body containing paragraphs
 */
export interface TextBody extends XmlNode {
  'a:p'?: Paragraph | Paragraph[];
}

/**
 * Paragraph structure
 */
export interface Paragraph extends XmlNode {
  'a:pPr'?: ParagraphProperties;
  'a:r'?: TextRun | TextRun[];
  'a:t'?: TextValue;
}

/**
 * Paragraph properties including list information
 */
export interface ParagraphProperties extends XmlNode {
  '@_lvl'?: string | number;
  'a:buAutoNum'?: any;
  'a:buFont'?: any;
  'a:buChar'?: any;
  'a:buBlip'?: any;
  'a:buNone'?: any;
}

/**
 * Text run containing formatted text
 */
export interface TextRun extends XmlNode {
  'a:rPr'?: RunProperties;
  'a:t'?: TextValue;
}

/**
 * Text run properties
 */
export interface RunProperties extends XmlNode {
  'a:sym'?: any;
}

/**
 * Result of paragraph text extraction
 */
export interface ParagraphResult {
  text: string;
  listLevel?: number;
  isNumbered?: boolean;
}

/**
 * Graphic frame containing tables or diagrams
 */
export interface GraphicFrame extends XmlNode {
  'a:graphic'?: Graphic;
}

/**
 * Graphic element
 */
export interface Graphic extends XmlNode {
  'a:graphicData'?: GraphicData;
}

/**
 * Graphic data element
 */
export interface GraphicData extends XmlNode {
  '@_uri'?: string;
  'a:tbl'?: Table;
}

/**
 * Table structure
 */
export interface Table extends XmlNode {
  'a:tr'?: TableRow | TableRow[];
}

/**
 * Table row
 */
export interface TableRow extends XmlNode {
  'a:tc'?: TableCell | TableCell[];
}

/**
 * Table cell
 */
export interface TableCell extends XmlNode {
  'a:txBody'?: TextBody;
}
