import * as fs from 'fs';
import { parseDocx } from './src/index';

async function main() {
  // Read a DOCX file
  const docxPath = process.argv[2] || 'example.docx';

  if (!fs.existsSync(docxPath)) {
    console.error(`Error: File '${docxPath}' not found`);
    console.log('Usage: ts-node docxExample.ts <path-to-docx-file>');
    process.exit(1);
  }

  console.log(`Parsing DOCX file: ${docxPath}\n`);

  const docxBuffer = fs.readFileSync(docxPath);

  // Parse with markdown conversion enabled
  const result = await parseDocx(docxBuffer, {
    convertToMarkdown: true,
    includeImages: true,
    includeFormulas: true
  });

  console.log(`Total pages: ${result.pages.length}\n`);

  // Display information about each page
  result.pages.forEach((page) => {
    console.log(`=== Page ${page.pageNumber} ===`);
    console.log(`Paragraphs: ${page.paragraphs.length}`);
    console.log(`Tables: ${page.tables.length}`);
    console.log(`Formulas: ${page.formulas.length}`);
    console.log(`Images: ${page.images.length}`);

    // Display paragraph styles
    const headings = page.paragraphs.filter(p => p.style?.startsWith('Heading'));
    if (headings.length > 0) {
      console.log(`\nHeadings:`);
      headings.forEach(h => {
        console.log(`  [${h.style}] ${h.text}`);
      });
    }

    // Display tables
    page.tables.forEach((table, idx) => {
      console.log(`\n--- Table ${idx + 1} ---`);
      console.log(`Rows: ${table.rows.length}`);
      console.log(`\nMarkdown:`);
      console.log(table.markdown);
    });

    // Display formulas
    if (page.formulas.length > 0) {
      console.log(`\nFormulas:`);
      page.formulas.forEach((formula, idx) => {
        console.log(`  [${idx + 1}] ${formula.text || '(OMML formula)'}`);
      });
    }

    // Display images
    if (page.images.length > 0) {
      console.log(`\nImages:`);
      page.images.forEach((image, idx) => {
        console.log(`  [${idx + 1}] ${image.fileName} (${image.contentType}, ${image.content.length} bytes)`);
        if (image.description) {
          console.log(`      Description: ${image.description}`);
        }

        // Save image to file
        const outputPath = `extracted_${image.fileName}`;
        fs.writeFileSync(outputPath, image.content);
        console.log(`      Saved to: ${outputPath}`);
      });
    }

    // Display markdown
    if (page.markdown) {
      console.log(`\n--- Markdown ---`);
      console.log(page.markdown);
      console.log(`---`);
    }

    console.log('\n');
  });
}

main().catch(console.error);
