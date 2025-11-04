/**
 * Represents an image referenced in a slide
 */
export interface SlideImage {
  /** The relationship ID used to reference this image in the slide XML */
  rId: string;
  /** The path within the PPTX archive (e.g., "ppt/media/image1.png") */
  path: string;
  /** The image file name */
  fileName: string;
  /** The binary content of the image */
  content: Buffer;
  /** The content type/MIME type of the image */
  contentType: string;
}

/**
 * Represents diagram data referenced in a slide
 */
export interface DiagramData {
  /** The relationship ID used to reference this diagram in the slide XML */
  rId: string;
  /** The path to the data XML file (e.g., "ppt/diagrams/data1.xml") */
  path: string;
  /** The XML content of the diagram data file */
  xml: string;
}

/**
 * Represents a single slide with its content and references
 */
export interface Slide {
  /** The slide number (1-indexed) */
  slideNumber: number;
  /** The XML content of the slide */
  xml: string;
  /** Array of images referenced in this slide */
  images: SlideImage[];
  /** Array of diagram data files referenced in this slide */
  diagrams: DiagramData[];
}

/**
 * Result of parsing a PPTX document
 */
export interface PptxParseResult {
  /** Array of slides extracted from the presentation */
  slides: Slide[];
}
