# OLFT - Office File Parser

A TypeScript library for parsing PowerPoint (.pptx), Excel (.xlsx/.xls), and Word (.docx) files, extracting content, images, tables, and formulas with intelligent conversion to Markdown and JSON.

## Features

### Word
- **DOCX (2007+)**: ✅ Full support
  - Extract content page-by-page
  - Extract paragraphs with formatting (bold, italic, headings)
  - Extract tables with markdown conversion
  - Extract mathematical formulas (OMML format)
  - Extract embedded images
  - Convert to Markdown
- **DOC (97-2003)**: ❌ Not supported (convert to DOCX first)

### PowerPoint
- **PPTX (2007+)**: ✅ Full support
  - Extract slide XML content
  - Extract all referenced images with metadata
  - Extract diagram data files (data*.xml)
  - Convert slides to Markdown
- **PPT (97-2003)**: ❌ Not supported (convert to PPTX first)

### Excel
- **XLSX (2007+)**: ✅ Full support
  - Extract all sheets with cell data
  - Detect and extract tables
  - Convert tables with nested headers to hierarchical JSON
  - Extract embedded images
  - Support for merged cells and complex table structures
  - Pivot tables treated as regular cells

- **XLS (97-2003)**: ✅ Supported with limitations
  - Extract all sheets with cell data
  - Detect and extract tables
  - Convert tables to JSON
  - Support for merged cells and formulas
  - ⚠️ **Image extraction not supported** (SheetJS limitation)

### General
- Full TypeScript support with type definitions
- Enterprise-grade error handling
- Production-ready with no memory leaks
- Automatic format detection utility

## Installation

```bash
npm install
```

## Usage

### PowerPoint (.pptx)

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

### Excel (.xlsx)

```typescript
import * as fs from 'fs';
import { parseXlsx } from 'olft';

async function main() {
  // Read XLSX file
  const xlsxBuffer = fs.readFileSync('data.xlsx');

  // Parse the file with JSON conversion enabled
  const result = await parseXlsx(xlsxBuffer, {
    convertToJson: true,
    includeImages: true
  });

  // Access sheets
  result.sheets.forEach((sheet) => {
    console.log(`Sheet: ${sheet.name}`);

    // Access tables
    sheet.tables.forEach((table) => {
      console.log(`Table: ${table.name}`);
      console.log(`Range: ${table.range}`);

      // Access markdown representation
      console.log(table.markdown);

      // Access JSON representation (if convertToJson: true)
      if (table.json) {
        console.log(JSON.stringify(table.json, null, 2));
      }

      // Check for nested headers
      if (table.hasHierarchicalHeaders) {
        console.log('This table has nested headers!');
        console.log('Merged headers:', table.mergedHeaders);
      }
    });

    // Access images
    sheet.images.forEach((image) => {
      console.log(`Image: ${image.fileName}`);
      fs.writeFileSync(`output/${image.fileName}`, image.content);
    });
  });
}
```

### Excel Legacy (.xls)

```typescript
import * as fs from 'fs';
import { parseXls } from 'olft';

async function main() {
  // Read XLS file (Excel 97-2003)
  const xlsBuffer = fs.readFileSync('legacy-data.xls');

  // Parse the file (same options as XLSX)
  const result = await parseXls(xlsBuffer, {
    convertToJson: true
  });

  // Note: Images are not extracted from XLS files
  // Data access is identical to XLSX
  result.sheets.forEach((sheet) => {
    console.log(`Sheet: ${sheet.name}`);
    sheet.tables.forEach((table) => {
      console.log(table.json);
    });
  });
}
```

### Auto-detect Format

```typescript
import { detectOfficeFormat, parseXlsx, parseXls } from 'olft';

async function main() {
  const buffer = fs.readFileSync('unknown-file.xls');

  // Detect format
  const format = await detectOfficeFormat(buffer);
  console.log(`Detected format: ${format}`);

  // Parse based on format
  if (format === 'xlsx') {
    const result = await parseXlsx(buffer);
  } else if (format === 'xls') {
    const result = await parseXls(buffer);
  } else if (format === 'pptx') {
    const result = await parsePptx(buffer);
  } else if (format === 'ppt') {
    throw new Error('PPT format not supported. Please convert to PPTX first.');
  }
}
```

## Markdown Conversion

### PowerPoint

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

### `parsePptx(pptxBuffer: Buffer, options?: PptxParseOptions): Promise<PptxParseResult>`

Parses a PPTX file and returns structured data.

**Parameters:**
- `pptxBuffer`: Buffer containing the PPTX file content
- `options` (optional): Parsing options
  - `parallel` (boolean, default: `false`): Process slides in parallel for better performance

**Returns:** Promise resolving to `PptxParseResult`

**Example:**
```typescript
// Default: Sequential processing (backward compatible)
const result = await parsePptx(pptxBuffer);

// Parallel processing for better performance
const result = await parsePptx(pptxBuffer, { parallel: true });
```

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

## Memory Considerations

This library loads the entire PPTX file into memory for processing. For typical presentations (< 50MB), this is not an issue. For very large files:

**Best Practices:**
- Set reasonable file size limits in your application (e.g., 50MB max)
- Process files sequentially rather than in parallel for memory efficiency
- Dereference result objects when finished to allow garbage collection:
  ```typescript
  let result = await parsePptx(buffer);
  // Use result...
  result = null; // Allow GC to clean up
  ```

**Parallel Mode:**
- Use `{ parallel: true }` for better performance on multi-core systems
- For very large presentations (100+ slides), sequential mode uses less memory

## Production Deployment

The library is production-ready and enterprise-grade:
- ✅ Comprehensive error handling with context
- ✅ Input validation on all public APIs
- ✅ TypeScript for type safety
- ✅ No memory leaks or resource leaks
- ✅ Warning logs for recoverable issues

See `PRODUCTION_READINESS_REVIEW.md` for detailed analysis.

## License

ISC
