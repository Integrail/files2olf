import * as fs from 'fs';
import { parseXlsx } from './src/index';

async function main() {
  // Read an XLSX file
  const xlsxPath = process.argv[2] || 'example.xlsx';

  if (!fs.existsSync(xlsxPath)) {
    console.error(`Error: File '${xlsxPath}' not found`);
    console.log('Usage: ts-node xlsxExample.ts <path-to-xlsx-file>');
    process.exit(1);
  }

  console.log(`Parsing XLSX file: ${xlsxPath}\n`);

  const xlsxBuffer = fs.readFileSync(xlsxPath);

  // Parse with JSON conversion enabled
  const result = await parseXlsx(xlsxBuffer, {
    convertToJson: true,
    includeImages: true
  });

  console.log(`Total sheets: ${result.sheets.length}\n`);

  // Display information about each sheet
  result.sheets.forEach((sheet) => {
    console.log(`=== Sheet ${sheet.index + 1}: ${sheet.name} ===`);
    console.log(`Tables: ${sheet.tables.length}`);
    console.log(`Merged cells: ${sheet.mergedCells.length}`);
    console.log(`Images: ${sheet.images.length}`);

    // Display merged cell information
    if (sheet.mergedCells.length > 0) {
      console.log('\nMerged Cells:');
      sheet.mergedCells.forEach((merge) => {
        console.log(`  ${merge.ref}: "${merge.value}" (${merge.colSpan}x${merge.rowSpan})`);
      });
    }

    // Display table information
    sheet.tables.forEach((table, idx) => {
      console.log(`\n--- Table ${idx + 1}: ${table.name} ---`);
      console.log(`Range: ${table.range}`);
      console.log(`Columns: ${table.columns.join(', ')}`);
      console.log(`Rows: ${table.data.length}`);
      console.log(`Has hierarchical headers: ${table.hasHierarchicalHeaders}`);

      // Show markdown representation
      console.log('\nMarkdown:');
      console.log(table.markdown);

      // Show JSON representation if available
      if (table.json) {
        console.log('\nJSON:');
        console.log(JSON.stringify(table.json, null, 2));
      }

      console.log('');
    });

    // Display images
    if (sheet.images.length > 0) {
      console.log('\nImages:');
      sheet.images.forEach((image, idx) => {
        console.log(`  [${idx + 1}] ${image.fileName} (${image.contentType}, ${image.content.length} bytes)`);
        if (image.position) {
          console.log(`      Position: Row ${image.position.row}, Col ${image.position.col}`);
        }

        // Save image to file
        const outputPath = `extracted_${image.fileName}`;
        fs.writeFileSync(outputPath, image.content);
        console.log(`      Saved to: ${outputPath}`);
      });
    }

    console.log('\n');
  });
}

main().catch(console.error);
