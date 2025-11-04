# OLFT - PowerPoint Parser

A TypeScript library for parsing PowerPoint (.pptx) files and extracting slide content, images, and diagram data.

## Features

- Extract slide XML content
- Extract all referenced images with metadata
- Extract diagram data files (data*.xml)
- Full TypeScript support with type definitions
- Zero dependencies except JSZip

## Installation

```bash
npm install
```

## Usage

```typescript
import * as fs from 'fs';
import { parsePptx, convertSlideToMarkdown, extractDiagramText } from 'olft';

async function main() {
  // Read PPTX file
  const pptxBuffer = fs.readFileSync('presentation.pptx');

  // Parse the file
  const result = await parsePptx(pptxBuffer);

  // Access slides
  result.slides.forEach((slide) => {
    console.log(`Slide ${slide.slideNumber}`);
    console.log(`XML: ${slide.xml}`);

    // Access images
    slide.images.forEach((image) => {
      console.log(`Image: ${image.fileName}`);
      console.log(`Content Type: ${image.contentType}`);
      // image.content is a Buffer with the image data
      fs.writeFileSync(`output/${image.fileName}`, image.content);
    });

    // Access diagram data
    slide.diagrams.forEach((diagram) => {
      console.log(`Diagram Data: ${diagram.path}`);
      console.log(`XML: ${diagram.xml}`);
    });

    // Convert slide to markdown
    const markdown = convertSlideToMarkdown(slide.xml);
    console.log(markdown);

    // Extract text from diagrams
    slide.diagrams.forEach((diagram) => {
      const diagramText = extractDiagramText(diagram.xml);
      console.log(diagramText);
    });
  });
}
```

## Markdown Conversion

The library includes a markdown converter that extracts text content from slides and formats it as markdown.

```typescript
import { parsePptx, convertSlideToMarkdown, extractDiagramText } from 'olft';

const result = await parsePptx(pptxBuffer);

result.slides.forEach((slide) => {
  // Convert slide content to markdown
  const markdown = convertSlideToMarkdown(slide.xml);
  console.log(markdown);
  // Output:
  // # Slide Title
  //
  // Body text content
  //
  // - Bullet point 1
  // - Bullet point 2
  //   - Nested bullet
  //
  // | Header 1 | Header 2 |
  // | --- | --- |
  // | Cell 1 | Cell 2 |

  // Extract text from diagram data
  slide.diagrams.forEach((diagram) => {
    const text = extractDiagramText(diagram.xml);
    console.log(text);
  });
});
```

**Supported Features:**
- Headings (titles and subtitles)
- Bullet lists (with nesting)
- Numbered lists
- Tables (converted to markdown format)
- Diagram text extraction

## API

### `parsePptx(pptxBuffer: Buffer): Promise<PptxParseResult>`

Parses a PPTX file and returns structured data.

**Parameters:**
- `pptxBuffer`: Buffer containing the PPTX file content

**Returns:** Promise resolving to `PptxParseResult`

### `convertSlideToMarkdown(slideXml: string): string`

Converts slide XML content to markdown format.

**Parameters:**
- `slideXml`: The XML content of a slide

**Returns:** Markdown-formatted string with the slide's text content

**Features:**
- Detects and converts titles to `# Heading`
- Detects and converts subtitles to `## Heading`
- Preserves bullet lists with proper nesting
- Preserves numbered lists
- Converts tables to markdown table format

### `extractDiagramText(diagramXml: string): string`

Extracts text content from diagram data XML.

**Parameters:**
- `diagramXml`: The XML content of a diagram data file

**Returns:** Extracted text content from the diagram

### Types

#### `PptxParseResult`
```typescript
interface PptxParseResult {
  slides: Slide[];
}
```

#### `Slide`
```typescript
interface Slide {
  slideNumber: number;      // 1-indexed
  xml: string;              // Slide XML content
  images: SlideImage[];     // Referenced images
  diagrams: DiagramData[];  // Referenced diagram data
}
```

#### `SlideImage`
```typescript
interface SlideImage {
  rId: string;           // Relationship ID
  path: string;          // Path in PPTX archive
  fileName: string;      // Image file name
  content: Buffer;       // Binary image data
  contentType: string;   // MIME type
}
```

#### `DiagramData`
```typescript
interface DiagramData {
  rId: string;   // Relationship ID
  path: string;  // Path to data XML file
  xml: string;   // XML content
}
```

## Example

Run the example script:

```bash
npm run build
npx ts-node example.ts path/to/your/presentation.pptx
```

## Building

```bash
npm run build
```

This compiles TypeScript to JavaScript in the `dist/` directory.

## License

ISC
