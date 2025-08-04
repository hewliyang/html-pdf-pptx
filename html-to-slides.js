#!/usr/bin/env node

const puppeteer = require("puppeteer");
const { execSync } = require("child_process");
const path = require("path");
const fs = require("fs");
const { JSDOM } = require("jsdom");
const { library, icon } = require("@fortawesome/fontawesome-svg-core");
const fas = require("@fortawesome/free-solid-svg-icons");
const far = require("@fortawesome/free-regular-svg-icons");
const fab = require("@fortawesome/free-brands-svg-icons");

// Add all FontAwesome icons to the library
library.add(fas.fas, far.far, fab.fab);

// Parse command line arguments
const args = process.argv.slice(2);
if (args.length === 0 || args.includes("--help") || args.includes("-h")) {
  console.log(`
HTML to Slides Converter
========================

Usage: node html-to-slides.js <html-file> [options]

Arguments:
  <html-file>     Path to the HTML file to convert

Options:
  --pdf-only      Generate only PDF output
  --pptx-only     Generate only PPTX output
  --width <px>    Set viewport width (default: 1280)
  --height <px>   Set viewport height (default: 720)
  --output <dir>  Output directory (default: same as input file)
  --help, -h      Show this help message

Examples:
  node html-to-slides.js presentation.html
  node html-to-slides.js slides.html --pdf-only
  node html-to-slides.js deck.html --width 1920 --height 1080
  node html-to-slides.js slides.html --output ./output/
`);
  process.exit(0);
}

// Parse arguments
const htmlFile = args[0];
const pdfOnly = args.includes("--pdf-only");
const pptxOnly = args.includes("--pptx-only");
const widthIndex = args.indexOf("--width");
const heightIndex = args.indexOf("--height");
const outputIndex = args.indexOf("--output");

const width = widthIndex !== -1 ? parseInt(args[widthIndex + 1]) : 1280;
const height = heightIndex !== -1 ? parseInt(args[heightIndex + 1]) : 720;
const outputDir =
  outputIndex !== -1
    ? args[outputIndex + 1]
    : path.dirname(path.resolve(htmlFile));

// Validate input file
if (!fs.existsSync(htmlFile)) {
  console.error(`❌ File not found: ${htmlFile}`);
  process.exit(1);
}

console.log(`🚀 Starting conversion for: ${htmlFile}`);
console.log(`📐 Dimensions: ${width}x${height}px`);
console.log(`📁 Output directory: ${outputDir}`);

function getIconInfo(classList) {
  // Find the icon name class (starts with fa- but not just "fa")
  const iconClass = classList.find(
    (cls) => cls.startsWith("fa-") && cls !== "fa"
  );
  if (!iconClass) return null;

  // Determine the icon style (solid, regular, brands)
  let iconStyle = "fas"; // default to solid
  if (classList.includes("far") || classList.includes("fa-regular"))
    iconStyle = "far";
  if (classList.includes("fab") || classList.includes("fa-brands"))
    iconStyle = "fab";

  return {
    prefix: iconStyle,
    iconName: iconClass.substring(3), // remove 'fa-' prefix
  };
}

function preprocessHtml(htmlContent) {
  console.log("🔄 Preprocessing HTML to rasterize FontAwesome icons...");

  const jsdom = new JSDOM(htmlContent);
  const { document } = jsdom.window;

  // First, update CSS rules that target <i> elements to also work with <svg> elements
  const styleElements = document.querySelectorAll("style");
  let cssRulesUpdated = 0;

  styleElements.forEach((styleElement) => {
    let cssText = styleElement.textContent;
    const originalCss = cssText;

    // Find CSS rules that target 'i' elements and replace them to target both 'i' and 'svg'
    // Pattern matches things like ".trend-item i {" or "div i," etc.
    cssText = cssText.replace(
      /([^{}]*?\s)(i)(\s*{)/g,
      (match, before, element, after) => {
        // Check if this is actually an element selector (not part of a word)
        if (before.match(/[a-zA-Z0-9_-]$/)) return match; // Don't replace if 'i' is part of a word

        // Replace "i {" with "i, svg {" to target both elements
        return before + "i, " + before.trim() + " svg" + after;
      }
    );

    if (cssText !== originalCss) {
      styleElement.textContent = cssText;
      cssRulesUpdated++;
    }
  });

  if (cssRulesUpdated > 0) {
    console.log(
      `🎨 Updated ${cssRulesUpdated} CSS style blocks to support SVG elements`
    );
  }

  // Find all FontAwesome icons
  const faIcons = document.querySelectorAll('i[class*="fa-"]');
  let replacedCount = 0;
  let skippedCount = 0;

  faIcons.forEach((iconElement) => {
    const classList = Array.from(iconElement.classList);
    const iconInfo = getIconInfo(classList);

    if (iconInfo) {
      try {
        // Generate SVG using FontAwesome library
        const faIcon = icon(iconInfo);

        if (faIcon && faIcon.html) {
          // Create SVG element from FontAwesome-generated HTML
          const svgHtml = faIcon.html[0];
          const tempDiv = document.createElement("div");
          tempDiv.innerHTML = svgHtml;
          const svg = tempDiv.firstElementChild;

          // PRESERVE ALL ORIGINAL CLASSES - this is the key fix!
          // Combine FontAwesome's SVG classes with ALL original <i> classes
          const originalClasses = classList.join(" ");
          const faClasses = svg.getAttribute("class") || "";
          svg.setAttribute("class", `${faClasses} ${originalClasses}`.trim());

          // Apply FontAwesome's default SVG sizing
          svg.style.height = "1em";
          svg.style.width = "1.25em";
          svg.style.verticalAlign = "-0.125em";
          svg.style.overflow = "visible";

          // Copy ALL inline styles from the original icon
          if (iconElement.style.cssText) {
            const existingStyles = svg.style.cssText;
            svg.style.cssText =
              existingStyles + "; " + iconElement.style.cssText;
          }

          // Copy any other attributes (id, data-*, etc.)
          Array.from(iconElement.attributes).forEach((attr) => {
            if (attr.name !== "class" && attr.name !== "style") {
              svg.setAttribute(attr.name, attr.value);
            }
          });

          // Replace the icon element
          iconElement.parentNode.replaceChild(svg, iconElement);
          replacedCount++;
        } else {
          skippedCount++;
        }
      } catch (error) {
        console.warn(
          `⚠️  Could not find icon: ${iconInfo.prefix} ${iconInfo.iconName}`
        );
        skippedCount++;
      }
    } else {
      skippedCount++;
    }
  });

  if (replacedCount > 0) {
    console.log(`✅ Replaced ${replacedCount} FontAwesome icons with SVG`);
  }
  if (skippedCount > 0) {
    console.log(`⚠️  Skipped ${skippedCount} unknown FontAwesome icons`);
  }

  // Add CSS to prevent page breaks and ensure content fits on one page
  const head = document.querySelector("head");
  if (head) {
    const pageBreakStyle = document.createElement("style");
    pageBreakStyle.textContent = `
      @media print {
        * {
          page-break-inside: avoid !important;
          page-break-after: avoid !important;
        }
        
        body, html {
          overflow: hidden !important;
          page-break-inside: avoid !important;
        }
        
        .slide-container {
          page-break-inside: avoid !important;
          page-break-after: avoid !important;
          max-height: 100vh !important;
          overflow: hidden !important;
        }
        
        /* Scale content to fit if needed */
        body {
          transform-origin: top left;
          width: 100vw;
          height: 100vh;
        }
      }
    `;
    head.appendChild(pageBreakStyle);
    console.log(
      "📐 Added CSS to prevent page breaks and ensure single-page output"
    );
  }

  return "<!DOCTYPE html>" + jsdom.serialize();
}

async function convertHtmlToPdf(inputPath, outputPath, maxRetries = 3) {
  console.log(`\n📄 Converting ${path.basename(inputPath)} to PDF...`);

  // Read and preprocess HTML
  const htmlContent = fs.readFileSync(inputPath, "utf8");
  const processedHtml = preprocessHtml(htmlContent);

  // Create temporary file with processed HTML
  const tempHtmlPath = path.join(
    path.dirname(inputPath),
    `temp_${Date.now()}_${path.basename(inputPath)}`
  );
  fs.writeFileSync(tempHtmlPath, processedHtml);

  let lastError;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    console.log(`🔄 Attempt ${attempt}/${maxRetries}`);

    const browser = await puppeteer.launch({
      headless: true,
      args: [
        "--no-sandbox",
        "--disable-setuid-sandbox",
        "--disable-dev-shm-usage",
      ],
    });

    try {
      const page = await browser.newPage();
      await page.setViewport({ width, height });

      const fileUrl = `file://${path.resolve(tempHtmlPath)}`;

      // Use shorter timeout and faster wait condition
      await page.goto(fileUrl, {
        waitUntil: "domcontentloaded",
        timeout: 10000,
      });

      // Wait for any dynamic content (images, webfonts) to settle
      await new Promise((resolve) => setTimeout(resolve, 1000));

      /*
       * Dynamically measure the rendered slide height. Some slides can be
       * taller than the default 720 px (e.g. long bullet lists).  If we keep
       * the PDF page height fixed, the bottom will be cut off.  Instead, we
       * query the `.slide-container` (fallback to <body>) for its scrollHeight
       * and grow the viewport as well as the PDF page size accordingly.
       */
      const contentHeight = await page.evaluate(() => {
        const el = document.querySelector(".slide-container") || document.body;
        return Math.ceil(el.scrollHeight || el.getBoundingClientRect().height);
      });

      if (contentHeight > height) {
        await page.setViewport({ width, height: contentHeight });
      }

      const widthInches = width / 96; // 1 CSS px = 1/96 in
      const heightInches = contentHeight / 96;

      await page.pdf({
        path: outputPath,
        width: `${widthInches}in`,
        height: `${heightInches}in`,
        printBackground: true,
        margin: { top: 0, right: 0, bottom: 0, left: 0 },
        preferCSSPageSize: false,
        pageRanges: "1",
        scale: 1,
      });

      console.log(`✅ PDF created: ${outputPath}`);

      // Success - clean up and return
      await browser.close();
      if (fs.existsSync(tempHtmlPath)) {
        fs.unlinkSync(tempHtmlPath);
      }
      return;
    } catch (error) {
      lastError = error;
      console.warn(`⚠️  Attempt ${attempt} failed: ${error.message}`);

      try {
        await browser.close();
      } catch (closeError) {
        console.warn(`⚠️  Error closing browser: ${closeError.message}`);
      }

      if (attempt < maxRetries) {
        console.log(`🔄 Retrying in 2 seconds...`);
        await new Promise((resolve) => setTimeout(resolve, 2000));
      }
    }
  }

  // All retries failed - clean up and throw error
  if (fs.existsSync(tempHtmlPath)) {
    fs.unlinkSync(tempHtmlPath);
  }

  throw new Error(
    `Failed to convert after ${maxRetries} attempts. Last error: ${lastError.message}`
  );
}

function convertPdfToPptx(pdfPath) {
  console.log(`\n📊 Converting PDF to PPTX...`);

  const pptxPath = pdfPath.replace(".pdf", ".pptx");
  const outputDirPath = path.dirname(pdfPath);

  try {
    execSync(
      `soffice --infilter=impress_pdf_import --convert-to pptx "${pdfPath}" --outdir "${outputDirPath}"`,
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
    const baseName = path.basename(htmlFile, ".html");
    const pdfPath = path.join(outputDir, `${baseName}.pdf`);

    // Ensure output directory exists
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    // Convert HTML to PDF
    if (!pptxOnly) {
      await convertHtmlToPdf(htmlFile, pdfPath);
    }

    // Convert PDF to PPTX if requested
    if (!pdfOnly && !pptxOnly) {
      convertPdfToPptx(pdfPath);
    } else if (pptxOnly) {
      // First create PDF, then convert to PPTX
      await convertHtmlToPdf(htmlFile, pdfPath);
      convertPdfToPptx(pdfPath);
      // Optionally delete intermediate PDF
      fs.unlinkSync(pdfPath);
    }

    console.log("\n🎉 Conversion completed successfully!");
    if (!pptxOnly) {
      console.log(`📄 PDF: ${pdfPath}`);
    }
    if (!pdfOnly) {
      const pptxPath = pdfPath.replace(".pdf", ".pptx");
      console.log(`📊 PPTX: ${pptxPath}`);
    }
  } catch (error) {
    console.error("❌ Error during conversion:", error.message);
    process.exit(1);
  }
}

main();
