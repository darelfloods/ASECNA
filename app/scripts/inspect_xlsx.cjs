const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");

function pickFirst(arr, predicate) {
  for (const x of arr) if (predicate(x)) return x;
  return null;
}

async function readZip(filePath) {
  const buf = fs.readFileSync(filePath);
  return JSZip.loadAsync(buf);
}

function extractTagBlock(xml, tagName) {
  const re = new RegExp(`<${tagName}\\b[^>]*>[\\s\\S]*?<\\/${tagName}>`, "i");
  const m = xml.match(re);
  return m ? m[0] : null;
}

function extractSelfClosing(xml, tagName) {
  const re = new RegExp(`<${tagName}\\b[^>]*/>`, "i");
  const m = xml.match(re);
  return m ? m[0] : null;
}

function hasTag(xml, tagName) {
  return new RegExp(`<${tagName}\\b`, "i").test(xml);
}

async function inspectWorkbook(zip, label) {
  const files = Object.keys(zip.files);
  const sheetFiles = files
    .filter((f) => /^xl\/worksheets\/sheet\d+\.xml$/.test(f))
    .sort((a, b) => {
      const an = Number(a.match(/sheet(\d+)\.xml$/)[1]);
      const bn = Number(b.match(/sheet(\d+)\.xml$/)[1]);
      return an - bn;
    });

  console.log(`\n=== ${label} ===`);
  console.log(`Worksheets: ${sheetFiles.length}`);

  for (const sf of sheetFiles) {
    const xml = await zip.file(sf).async("string");
    const rel = `xl/worksheets/_rels/${path.basename(sf)}.rels`;
    const relXml = zip.file(rel) ? await zip.file(rel).async("string") : null;

    console.log(`\n-- ${sf} --`);
    console.log(`has <headerFooter>: ${hasTag(xml, "headerFooter")}`);
    console.log(`has <legacyDrawingHF>: ${hasTag(xml, "legacyDrawingHF")}`);
    console.log(`has <drawing>: ${hasTag(xml, "drawing")}`);
    console.log(`has <pageSetup>: ${hasTag(xml, "pageSetup")}`);
    console.log(`has <pageMargins>: ${hasTag(xml, "pageMargins")}`);
    console.log(`has <rowBreaks>: ${hasTag(xml, "rowBreaks")}`);
    console.log(`has <colBreaks>: ${hasTag(xml, "colBreaks")}`);

    const legacy = extractSelfClosing(xml, "legacyDrawingHF");
    const drawing = extractSelfClosing(xml, "drawing");
    const headerFooter = extractTagBlock(xml, "headerFooter");
    const pageSetup = extractSelfClosing(xml, "pageSetup") || extractTagBlock(xml, "pageSetup");

    if (legacy) console.log(`legacyDrawingHF: ${legacy}`);
    if (drawing) console.log(`drawing: ${drawing}`);
    if (pageSetup) console.log(`pageSetup: ${pageSetup.slice(0, 160)}...`);
    if (headerFooter) console.log(`headerFooter: ${headerFooter.slice(0, 160)}...`);

    if (relXml) {
      const relLines = relXml
        .split(/\r?\n/)
        .filter((l) => l.includes("Relationship"))
        .slice(0, 30);
      console.log(`rels present: yes (${relLines.length} relationship lines shown)`);
      for (const l of relLines) console.log(`  ${l.trim()}`);
    } else {
      console.log(`rels present: no`);
    }
  }

  const media = files.filter((f) => f.startsWith("xl/media/"));
  const drawings = files.filter((f) => f.startsWith("xl/drawings/"));
  const vml = files.filter((f) => f.includes("vmlDrawing"));
  console.log(`\nassets: media=${media.length}, drawings=${drawings.length}, vmlDrawing=${vml.length}`);
}

async function main() {
  const template = process.argv[2];
  const generated = process.argv[3];
  if (!template || !generated) {
    console.error("Usage: node scripts/inspect_xlsx.cjs <template.xlsx> <generated.xlsx>");
    process.exit(1);
  }

  const templateZip = await readZip(template);
  const genZip = await readZip(generated);

  await inspectWorkbook(templateZip, "TEMPLATE");
  await inspectWorkbook(genZip, "GENERATED");

  // Heuristic: find which template sheet contains the invoice header/footer/drawings
  const templateSheets = Object.keys(templateZip.files).filter((f) => /^xl\/worksheets\/sheet\d+\.xml$/.test(f));
  const likely = [];
  for (const sf of templateSheets) {
    const xml = await templateZip.file(sf).async("string");
    if (hasTag(xml, "legacyDrawingHF") || hasTag(xml, "drawing") || hasTag(xml, "headerFooter")) {
      likely.push(sf);
    }
  }
  console.log("\nTemplate sheets with header/footer/drawing:", likely);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});

