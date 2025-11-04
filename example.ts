import * as fs from 'fs';
import { parsePptx, convertSlideToMarkdown, extractDiagramText } from './src/index';

async function main() {
  // Read a PPTX file
  const pptxPath = process.argv[2] || 'example.pptx';

  if (!fs.existsSync(pptxPath)) {
    console.error(`Error: File '${pptxPath}' not found`);
    console.log('Usage: ts-node example.ts <path-to-pptx-file>');
    process.exit(1);
  }

  console.log(`Parsing PPTX file: ${pptxPath}\n`);

  const pptxBuffer = fs.readFileSync(pptxPath);
  const result = await parsePptx(pptxBuffer);

  console.log(`Total slides: ${result.slides.length}\n`);

  // Display information about each slide
  result.slides.forEach((slide) => {
    console.log(`--- Slide ${slide.slideNumber} ---`);
    console.log(`XML length: ${slide.xml.length} characters`);
    console.log(`Images: ${slide.images.length}`);

    slide.images.forEach((img, idx) => {
      console.log(`  [${idx + 1}] ${img.fileName} (rId: ${img.rId}, type: ${img.contentType}, size: ${img.content.length} bytes)`);
    });

    console.log(`Diagrams: ${slide.diagrams.length}`);
    slide.diagrams.forEach((diagram, idx) => {
      console.log(`  [${idx + 1}] ${diagram.path} (rId: ${diagram.rId}, XML length: ${diagram.xml.length} characters)`);
    });

    // Convert slide to markdown
    console.log('\nMarkdown Content:');
    console.log('---');
    const markdown = convertSlideToMarkdown(slide.xml);
    console.log(markdown || '(No text content found)');
    console.log('---');

    // Extract diagram text if available
    if (slide.diagrams.length > 0) {
      console.log('\nDiagram Text:');
      slide.diagrams.forEach((diagram, idx) => {
        const diagramText = extractDiagramText(diagram.xml);
        if (diagramText) {
          console.log(`  Diagram ${idx + 1}: ${diagramText}`);
        }
      });
    }

    console.log('');
  });

  // Example: Save first image from first slide (if exists)
  if (result.slides.length > 0 && result.slides[0].images.length > 0) {
    const firstImage = result.slides[0].images[0];
    const outputPath = `extracted_${firstImage.fileName}`;
    fs.writeFileSync(outputPath, firstImage.content);
    console.log(`Saved first image to: ${outputPath}`);
  }

  // Example: Save first diagram data (if exists)
  if (result.slides.length > 0 && result.slides[0].diagrams.length > 0) {
    const firstDiagram = result.slides[0].diagrams[0];
    const outputPath = 'extracted_diagram_data.xml';
    fs.writeFileSync(outputPath, firstDiagram.xml);
    console.log(`Saved first diagram data to: ${outputPath}`);
  }
}

main().catch(console.error);
