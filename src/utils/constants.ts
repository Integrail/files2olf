/**
 * PowerPoint XML element paths
 */
export const XML_PATHS = {
  // Slide structure
  SLIDE: 'p:sld',
  COMMON_SLIDE_DATA: 'p:cSld',
  SHAPE_TREE: 'p:spTree',

  // Shapes and frames
  SHAPE: 'p:sp',
  GRAPHIC_FRAME: 'p:graphicFrame',

  // Shape properties
  NON_VISUAL_SHAPE_PROPS: 'p:nvSpPr',
  NON_VISUAL_PROPS: 'p:nvPr',
  PLACEHOLDER: 'p:ph',
  TEXT_BODY: 'p:txBody',

  // Text elements
  PARAGRAPH: 'a:p',
  PARAGRAPH_PROPS: 'a:pPr',
  TEXT_RUN: 'a:r',
  RUN_PROPS: 'a:rPr',
  TEXT: 'a:t',

  // List properties
  BULLET_AUTO_NUM: 'a:buAutoNum',
  BULLET_FONT: 'a:buFont',
  BULLET_CHAR: 'a:buChar',
  BULLET_BLIP: 'a:buBlip',
  BULLET_NONE: 'a:buNone',
  SYMBOL_FONT: 'a:sym',

  // Table elements
  TABLE: 'a:tbl',
  TABLE_ROW: 'a:tr',
  TABLE_CELL: 'a:tc',

  // Graphic data
  GRAPHIC: 'a:graphic',
  GRAPHIC_DATA: 'a:graphicData',

  // Diagram elements
  DIAGRAM_REL_IDS: 'dgm:relIds',
  DIAGRAM_TEXT: 'dgm:t',

  // Attributes
  ATTR_TYPE: '@_type',
  ATTR_LEVEL: '@_lvl',
  ATTR_URI: '@_uri'
} as const;

/**
 * PowerPoint placeholder types
 */
export const PLACEHOLDER_TYPES = {
  TITLE: 'title',
  CENTER_TITLE: 'ctrTitle',
  SUBTITLE: 'subTitle',
  BODY: 'body',
  OBJECT: 'obj'
} as const;

/**
 * Precompiled regular expressions for better performance
 */
export const REGEX_PATTERNS = {
  // Slide file pattern: ppt/slides/slide1.xml, ppt/slides/slide2.xml, etc.
  SLIDE_FILE: /^ppt\/slides\/slide\d+\.xml$/,

  // Extract slide number from filename
  SLIDE_NUMBER: /slide(\d+)\.xml$/,

  // Image file extensions
  IMAGE_EXTENSION: /\.(png|jpg|jpeg|gif|bmp|svg|tiff?)$/i,

  // Diagram data file pattern
  DIAGRAM_DATA: /data\d*\.xml$/i
} as const;

/**
 * Content type URIs for graphic data
 */
export const GRAPHIC_URIS = {
  TABLE: 'table',
  DIAGRAM: 'diagram'
} as const;
