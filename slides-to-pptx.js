#!/usr/bin/env node

const { execSync } = require("child_process");
const { PDFDocument } = require("pdf-lib");
const fs = require("fs");
const path = require("path");

// Parse command line arguments
const args = process.argv.slice(2);
if (args.length === 0 || args.includes("--help") || args.includes("-h")) {
  console.log(`
Slides JSON to PPTX Converter
============================

Usage: node slides-to-pptx.js <slides.json> [options]

Arguments:
  <slides.json>   Path to the slides.json file

Options:
  --pdf-only      Generate only merged PDF output
  --keep-pdfs     Keep individual PDF files after merging
  --output <dir>  Output directory (default: same as slides.json)
  --help, -h      Show this help message

Examples:
  node slides-to-pptx.js aws-costs/slides.json
  node slides-to-pptx.js slides.json --pdf-only
  node slides-to-pptx.js slides.json --output ./output/
`);
  process.exit(0);
}

const slidesJsonPath = args[0];
const pdfOnly = args.includes("--pdf-only");
const keepPdfs = args.includes("--keep-pdfs");
const outputIndex = args.indexOf("--output");
const outputDir = outputIndex !== -1 ? args[outputIndex + 1] : path.dirname(path.resolve(slidesJsonPath));

// Validate input file
if (!fs.existsSync(slidesJsonPath)) {
  console.error(`❌ File not found: ${slidesJsonPath}`);
  process.exit(1);
}

console.log(`🚀 Starting slides.json to PPTX conversion: ${slidesJsonPath}`);
console.log(`📁 Output directory: ${outputDir}`);

async function convertHtmlToPdf(htmlFile, targetDir) {
  console.log(`🔄 Converting ${path.basename(htmlFile)} to PDF...`);
  try {
    const htmlToSlidesPath = path.join(__dirname, "html-to-slides.js");
    execSync(`node "${htmlToSlidesPath}" "${htmlFile}" --pdf-only --output "${targetDir}"`, {
      stdio: "inherit",
    });
    return path.join(targetDir, path.basename(htmlFile, ".html") + ".pdf");
  } catch (error) {
    console.error(`❌ Failed to convert ${htmlFile}:`, error.message);
    throw error;
  }
}

async function mergePdfs(pdfPaths, outputPath) {
  console.log("📄 Merging PDFs in order...");

  const mergedPdf = await PDFDocument.create();

  for (const pdfPath of pdfPaths) {
    if (!fs.existsSync(pdfPath)) {
      console.error(`❌ PDF not found: ${pdfPath}`);
      continue;
    }

    console.log(`📎 Adding ${path.basename(pdfPath)}`);
    const pdfBytes = fs.readFileSync(pdfPath);
    const pdf = await PDFDocument.load(pdfBytes);
    const pages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
    pages.forEach((page) => mergedPdf.addPage(page));
  }

  const pdfBytes = await mergedPdf.save();
  fs.writeFileSync(outputPath, pdfBytes);
  console.log(`✅ Merged PDF created: ${outputPath}`);
}

function convertToPptx(pdfFile) {
  console.log("📊 Converting merged PDF to PPTX...");

  const pptxPath = pdfFile.replace(".pdf", ".pptx");
  const outputDirPath = path.dirname(pdfFile);

  try {
    execSync(
      `soffice --infilter=impress_pdf_import --convert-to pptx "${pdfFile}" --outdir "${outputDirPath}"`,
      {
        stdio: "pipe",
      }
    );
    console.log(`✅ PPTX created: ${pptxPath}`);
    return pptxPath;
  } catch (error) {
    console.error("❌ Error converting to PPTX:", error.message);
    console.error("Make sure LibreOffice is installed and available in PATH");
    throw error;
  }
}

async function main() {
  try {
    // Read and parse slides.json
    const slidesData = JSON.parse(fs.readFileSync(slidesJsonPath, 'utf8'));
    const { slide_ids, files } = slidesData;
    
    if (!slide_ids || !files) {
      console.error("❌ Invalid slides.json format. Expected 'slide_ids' and 'files' properties.");
      process.exit(1);
    }

    console.log(`📋 Found ${slide_ids.length} slides in order:`);
    slide_ids.forEach((id, index) => {
      console.log(`   ${index + 1}. ${id}`);
    });
    console.log();

    // Create output directory if it doesn't exist
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    // Find HTML files for each slide in order
    const htmlFiles = [];
    const slidesDir = path.dirname(slidesJsonPath);
    
    for (const slideId of slide_ids) {
      // Find the file with matching id
      const fileData = files.find(f => f.id === slideId);
      if (fileData) {
        // Check if HTML file exists in slides directory
        const htmlFile = path.join(slidesDir, `${slideId}.html`);
        if (fs.existsSync(htmlFile)) {
          htmlFiles.push(htmlFile);
          console.log(`✅ Found: ${slideId}.html`);
        } else {
          console.warn(`⚠️  Missing: ${slideId}.html`);
        }
      } else {
        console.warn(`⚠️  No file data found for slide: ${slideId}`);
      }
    }

    if (htmlFiles.length === 0) {
      console.error("❌ No HTML files found. Check that HTML files exist in the slides directory.");
      process.exit(1);
    }

    console.log(`\n🔄 Converting ${htmlFiles.length} HTML files to PDFs...\n`);

    // Convert each HTML to PDF
    const pdfPaths = [];
    for (const htmlFile of htmlFiles) {
      const pdfPath = await convertHtmlToPdf(htmlFile, outputDir);
      pdfPaths.push(pdfPath);
    }

    console.log(`\n📊 Generated ${pdfPaths.length} PDFs\n`);

    // Determine output filename based on slides directory name
    const slidesBaseName = path.basename(slidesDir);
    const mergedPdfPath = path.join(outputDir, `${slidesBaseName}-presentation.pdf`);
    
    // Merge PDFs
    await mergePdfs(pdfPaths, mergedPdfPath);

    // Convert to PPTX if requested
    let pptxPath;
    if (!pdfOnly) {
      pptxPath = convertToPptx(mergedPdfPath);
    }

    console.log("\n🎉 Conversion completed successfully!");
    console.log(`📄 PDF: ${mergedPdfPath}`);
    if (pptxPath) {
      console.log(`📊 PPTX: ${pptxPath}`);
    }

    // Cleanup individual PDFs unless requested to keep them
    if (!keepPdfs) {
      console.log("\n🧹 Cleaning up individual PDFs...");
      pdfPaths.forEach((pdfPath) => {
        if (fs.existsSync(pdfPath)) {
          fs.unlinkSync(pdfPath);
          console.log(`🗑️  Deleted ${path.basename(pdfPath)}`);
        }
      });
    }

  } catch (error) {
    console.error("❌ Process failed:", error.message);
    process.exit(1);
  }
}

main();