#!/usr/bin/env node

const { execSync } = require("child_process");
const { PDFDocument } = require("pdf-lib");
const puppeteer = require("puppeteer");
const fs = require("fs");
const path = require("path");
const { JSDOM } = require("jsdom");
const { library, icon } = require("@fortawesome/fontawesome-svg-core");
const fas = require("@fortawesome/free-solid-svg-icons");
const far = require("@fortawesome/free-regular-svg-icons");
const fab = require("@fortawesome/free-brands-svg-icons");

library.add(fas.fas, far.far, fab.fab);

function printHelpAndExit() {
  console.log(`
Manus Slides CLI
================

Usage: manus-slides <slides.json> [options]

Arguments:
  <slides.json>   Path to the Manus slides.json file

Options:
  --pdf-only          Generate only merged PDF output
  --keep-pdfs         Keep individual slide PDF files after merging
  --width <px>        Viewport width for HTML render (default: 1280)
  --height <px>       Viewport height for HTML render (default: 720)
  --concurrency <n>   Number of concurrent HTML->PDF conversions (default: 4)
  --output <dir>      Output directory (default: slides.json directory)
  --help, -h          Show this help message

Examples:
  manus-slides aws-costs/slides.json
  manus-slides aws-costs/slides.json --pdf-only
  manus-slides aws-costs/slides.json --width 1920 --height 1080
  manus-slides aws-costs/slides.json --output ./dist
`);
  process.exit(0);
}

function parseArgs(argv) {
  const args = argv.slice(2);
  if (args.length === 0 || args.includes("--help") || args.includes("-h")) {
    printHelpAndExit();
  }

  const slidesJsonPath = args[0];
  const pdfOnly = args.includes("--pdf-only");
  const keepPdfs = args.includes("--keep-pdfs");
  const widthIndex = args.indexOf("--width");
  const heightIndex = args.indexOf("--height");
  const concurrencyIndex = args.indexOf("--concurrency");
  const outputIndex = args.indexOf("--output");

  const width = widthIndex !== -1 ? parseInt(args[widthIndex + 1], 10) : 1280;
  const height = heightIndex !== -1 ? parseInt(args[heightIndex + 1], 10) : 720;
  const concurrency =
    concurrencyIndex !== -1
      ? Math.max(1, parseInt(args[concurrencyIndex + 1], 10) || 4)
      : 4;
  const outputDir =
    outputIndex !== -1
      ? args[outputIndex + 1]
      : path.dirname(path.resolve(slidesJsonPath));

  return {
    slidesJsonPath,
    pdfOnly,
    keepPdfs,
    width,
    height,
    concurrency,
    outputDir,
  };
}

function getIconInfo(classList) {
  const iconClass = classList.find(
    (cls) => cls.startsWith("fa-") && cls !== "fa"
  );
  if (!iconClass) return null;

  let iconStyle = "fas";
  if (classList.includes("far") || classList.includes("fa-regular"))
    iconStyle = "far";
  if (classList.includes("fab") || classList.includes("fa-brands"))
    iconStyle = "fab";

  return { prefix: iconStyle, iconName: iconClass.substring(3) };
}

function preprocessHtml(htmlContent) {
  const jsdom = new JSDOM(htmlContent);
  const { document } = jsdom.window;

  const styleElements = document.querySelectorAll("style");
  styleElements.forEach((styleElement) => {
    let cssText = styleElement.textContent;
    const originalCss = cssText;
    cssText = cssText.replace(
      /([^{}]*?\s)(i)(\s*{)/g,
      (match, before, element, after) => {
        if (before.match(/[a-zA-Z0-9_-]$/)) return match;
        return before + "i, " + before.trim() + " svg" + after;
      }
    );
    if (cssText !== originalCss) {
      styleElement.textContent = cssText;
    }
  });

  const faIcons = document.querySelectorAll('i[class*="fa-"]');
  faIcons.forEach((iconElement) => {
    const classList = Array.from(iconElement.classList);
    const iconInfo = getIconInfo(classList);
    if (!iconInfo) return;
    try {
      const faIcon = icon(iconInfo);
      if (!faIcon || !faIcon.html) return;
      const svgHtml = faIcon.html[0];
      const tempDiv = document.createElement("div");
      tempDiv.innerHTML = svgHtml;
      const svg = tempDiv.firstElementChild;
      const originalClasses = classList.join(" ");
      const faClasses = svg.getAttribute("class") || "";
      svg.setAttribute("class", `${faClasses} ${originalClasses}`.trim());
      svg.style.height = "1em";
      svg.style.width = "1.25em";
      svg.style.verticalAlign = "-0.125em";
      svg.style.overflow = "visible";
      if (iconElement.style.cssText) {
        const existingStyles = svg.style.cssText;
        svg.style.cssText = existingStyles + "; " + iconElement.style.cssText;
      }
      Array.from(iconElement.attributes).forEach((attr) => {
        if (attr.name !== "class" && attr.name !== "style") {
          svg.setAttribute(attr.name, attr.value);
        }
      });
      iconElement.parentNode.replaceChild(svg, iconElement);
    } catch (_) {
      /* ignore unknown icon */
    }
  });

  const head = document.querySelector("head");
  if (head) {
    const pageBreakStyle = document.createElement("style");
    pageBreakStyle.textContent = `
      @media print {
        * { page-break-inside: avoid !important; page-break-after: avoid !important; }
        body, html { overflow: hidden !important; page-break-inside: avoid !important; }
        .slide-container { page-break-inside: avoid !important; page-break-after: avoid !important; max-height: 100vh !important; overflow: hidden !important; }
        body { transform-origin: top left; width: 100vw; height: 100vh; }
      }
    `;
    head.appendChild(pageBreakStyle);
  }

  return "<!DOCTYPE html>" + jsdom.serialize();
}

async function convertHtmlToPdf(
  htmlPath,
  outputPdfPath,
  width,
  height,
  maxRetries = 3
) {
  const htmlContent = fs.readFileSync(htmlPath, "utf8");
  const processedHtml = preprocessHtml(htmlContent);

  const tempHtmlPath = path.join(
    path.dirname(htmlPath),
    `temp_${Date.now()}_${path.basename(htmlPath)}`
  );
  fs.writeFileSync(tempHtmlPath, processedHtml);

  let lastError;
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
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
      await page.goto(fileUrl, {
        waitUntil: "domcontentloaded",
        timeout: 10000,
      });
      await new Promise((resolve) => setTimeout(resolve, 1000));
      const contentHeight = await page.evaluate(() => {
        const el = document.querySelector(".slide-container") || document.body;
        return Math.ceil(el.scrollHeight || el.getBoundingClientRect().height);
      });
      if (contentHeight > height) {
        await page.setViewport({ width, height: contentHeight });
      }
      const widthInches = width / 96;
      const heightInches = (contentHeight || height) / 96;
      await page.pdf({
        path: outputPdfPath,
        width: `${widthInches}in`,
        height: `${heightInches}in`,
        printBackground: true,
        margin: { top: 0, right: 0, bottom: 0, left: 0 },
        preferCSSPageSize: false,
        pageRanges: "1",
        scale: 1,
      });
      await browser.close();
      fs.existsSync(tempHtmlPath) && fs.unlinkSync(tempHtmlPath);
      return;
    } catch (error) {
      lastError = error;
      try {
        await browser.close();
      } catch (_) {}
      if (attempt < maxRetries) {
        await new Promise((r) => setTimeout(r, 2000));
      }
    }
  }
  fs.existsSync(tempHtmlPath) && fs.unlinkSync(tempHtmlPath);
  throw new Error(
    `Failed to render HTML to PDF: ${
      lastError ? lastError.message : "unknown error"
    }`
  );
}

async function mergePdfs(pdfPaths, outputPath) {
  const mergedPdf = await PDFDocument.create();
  for (const pdfPath of pdfPaths) {
    if (!fs.existsSync(pdfPath)) continue;
    const pdfBytes = fs.readFileSync(pdfPath);
    const pdf = await PDFDocument.load(pdfBytes);
    const pages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
    pages.forEach((page) => mergedPdf.addPage(page));
  }
  const pdfBytes = await mergedPdf.save();
  fs.writeFileSync(outputPath, pdfBytes);
}

function convertToPptx(pdfFile) {
  const pptxPath = pdfFile.replace(/\.pdf$/, ".pptx");
  const outputDirPath = path.dirname(pdfFile);
  execSync(
    `soffice --infilter=impress_pdf_import --convert-to pptx "${pdfFile}" --outdir "${outputDirPath}"`,
    {
      stdio: "pipe",
    }
  );
  return pptxPath;
}

async function concurrentMap(items, concurrency, mapFn) {
  const results = new Array(items.length);
  let nextIndex = 0;
  let active = 0;
  return new Promise((resolve, reject) => {
    const schedule = () => {
      if (nextIndex >= items.length && active === 0) return resolve(results);
      while (active < concurrency && nextIndex < items.length) {
        const current = nextIndex++;
        active++;
        Promise.resolve(mapFn(items[current], current))
          .then((res) => {
            results[current] = res;
            active--;
            schedule();
          })
          .catch(reject);
      }
    };
    schedule();
  });
}

async function main() {
  const {
    slidesJsonPath,
    pdfOnly,
    keepPdfs,
    width,
    height,
    concurrency,
    outputDir,
  } = parseArgs(process.argv);
  const overallStartMs = Date.now();

  if (!fs.existsSync(slidesJsonPath)) {
    console.error(`❌ File not found: ${slidesJsonPath}`);
    process.exit(1);
  }

  console.log(`🚀 Starting slides.json to PPTX conversion: ${slidesJsonPath}`);
  console.log(`📐 Dimensions: ${width}x${height}px`);
  console.log(`⚙️  Concurrency: ${concurrency}`);
  console.log(`📁 Output directory: ${outputDir}`);

  const slidesData = JSON.parse(fs.readFileSync(slidesJsonPath, "utf8"));
  const { slide_ids, files } = slidesData;
  if (!slide_ids || !files) {
    console.error(
      "❌ Invalid slides.json format. Expected 'slide_ids' and 'files' properties."
    );
    process.exit(1);
  }

  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const slidesDir = path.dirname(slidesJsonPath);
  const htmlFiles = [];
  for (const slideId of slide_ids) {
    const fileData = files.find((f) => f.id === slideId);
    if (!fileData) continue;
    const htmlFile = path.join(slidesDir, `${slideId}.html`);
    if (fs.existsSync(htmlFile)) htmlFiles.push(htmlFile);
  }
  if (htmlFiles.length === 0) {
    console.error(
      "❌ No HTML files found. Check that HTML files exist next to slides.json."
    );
    process.exit(1);
  }

  const renderStartMs = Date.now();
  const individualPdfPaths = await concurrentMap(
    htmlFiles,
    concurrency,
    async (htmlFile) => {
      const pdfPath = path.join(
        outputDir,
        path.basename(htmlFile, ".html") + ".pdf"
      );
      console.log(
        `🔄 Converting ${path.basename(htmlFile)} → ${path.basename(pdfPath)}`
      );
      await convertHtmlToPdf(htmlFile, pdfPath, width, height);
      return pdfPath;
    }
  );
  const renderElapsedMs = Date.now() - renderStartMs;

  const mergedPdfPath = path.join(outputDir, "presentation.pdf");
  console.log("📄 Merging PDFs → presentation.pdf");
  const mergeStartMs = Date.now();
  await mergePdfs(individualPdfPaths, mergedPdfPath);
  const mergeElapsedMs = Date.now() - mergeStartMs;

  let pptxPath;
  let pptxElapsedMs = 0;
  if (!pdfOnly) {
    console.log("📊 Converting merged PDF to PPTX → presentation.pptx");
    try {
      const pptxStartMs = Date.now();
      pptxPath = convertToPptx(mergedPdfPath);
      pptxElapsedMs = Date.now() - pptxStartMs;
    } catch (error) {
      console.error("❌ Error converting to PPTX:", error.message);
      console.error("Make sure LibreOffice is installed and available in PATH");
      process.exit(1);
    }
  }

  if (!keepPdfs) {
    console.log("🧹 Cleaning up individual slide PDFs...");
    individualPdfPaths.forEach((p) => {
      if (fs.existsSync(p)) fs.unlinkSync(p);
    });
  }

  console.log("\n🎉 Conversion completed successfully!");
  console.log(`📄 PDF: ${mergedPdfPath}`);
  if (pptxPath) {
    console.log(`📊 PPTX: ${pptxPath}`);
  }
  const totalElapsedMs = Date.now() - overallStartMs;
  const baseTiming = `⏱️  Timings — HTML→PDF: ${(
    renderElapsedMs / 1000
  ).toFixed(1)}s, Merge: ${(mergeElapsedMs / 1000).toFixed(1)}s`;
  const pptxTiming = pptxPath
    ? `, PPTX: ${(pptxElapsedMs / 1000).toFixed(1)}s`
    : "";
  console.log(
    `${baseTiming}${pptxTiming}, Total: ${(totalElapsedMs / 1000).toFixed(1)}s`
  );
}

main();
