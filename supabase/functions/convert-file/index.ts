import "jsr:@supabase/functions-js/edge-runtime.d.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Client-Info, Apikey",
};

const CONVERT_MIME: Record<string, Record<string, string>> = {
  pdf: {
    ".pdf": "application/pdf",
  },
  word: {
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".doc": "application/msword",
    ".odt": "application/vnd.oasis.opendocument.text",
  },
  excel: {
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".xls": "application/vnd.ms-excel",
    ".ods": "application/vnd.oasis.opendocument.spreadsheet",
    ".csv": "text/csv",
  },
  powerpoint: {
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".ppt": "application/vnd.ms-powerpoint",
    ".odp": "application/vnd.oasis.opendocument.presentation",
  },
  pictures: {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".svg": "image/svg+xml",
  },
};

function getExtFromMime(mime: string): string {
  for (const cat of Object.values(CONVERT_MIME)) {
    for (const [ext, m] of Object.entries(cat)) {
      if (m === mime) return ext;
    }
  }
  return "";
}

function getCategoryForMime(mime: string): string | null {
  for (const [cat, formats] of Object.entries(CONVERT_MIME)) {
    for (const m of Object.values(formats)) {
      if (m === mime) return cat;
    }
  }
  return null;
}

function getCategoryForExt(ext: string): string | null {
  const e = ext.startsWith(".") ? ext : "." + ext;
  for (const [cat, formats] of Object.entries(CONVERT_MIME)) {
    if (formats[e]) return cat;
  }
  return null;
}

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { status: 200, headers: corsHeaders });
  }

  try {
    const formData = await req.formData();
    const file = formData.get("file") as File | null;
    const targetFormat = formData.get("targetFormat") as string | null;

    if (!file || !targetFormat) {
      return new Response(
        JSON.stringify({ error: "Faltan parametros: file y targetFormat son requeridos" }),
        { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    const targetExt = targetFormat.startsWith(".") ? targetFormat : "." + targetFormat;
    const sourceMime = file.type;
    const sourceExt = getExtFromMime(sourceMime) || "." + file.name.split(".").pop()?.toLowerCase();
    const sourceCat = getCategoryForMime(sourceMime) || getCategoryForExt(sourceExt);
    const targetCat = getCategoryForExt(targetExt);

    if (!sourceCat || !targetCat) {
      return new Response(
        JSON.stringify({ error: "Formato de origen o destino no reconocido" }),
        { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    if (sourceCat !== targetCat) {
      return new Response(
        JSON.stringify({ error: "Solo se permite conversion dentro de la misma categoria" }),
        { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    const arrayBuffer = await file.arrayBuffer();
    const baseName = file.name.replace(/\.[^.]+$/, "");
    const newFileName = `${baseName}${targetExt}`;
    const targetMime = CONVERT_MIME[targetCat][targetExt] || "application/octet-stream";

    let converted: ArrayBuffer;

    if (sourceCat === "pictures") {
      converted = await convertImage(arrayBuffer, sourceMime, targetMime, targetExt);
    } else if (sourceCat === "excel") {
      converted = await convertSpreadsheet(arrayBuffer, sourceMime, targetExt);
    } else if (sourceCat === "word") {
      converted = await convertWord(arrayBuffer, sourceMime, targetExt);
    } else if (sourceCat === "powerpoint") {
      converted = await convertPresentation(arrayBuffer, sourceMime, targetExt);
    } else {
      converted = arrayBuffer;
    }

    return new Response(converted, {
      status: 200,
      headers: {
        ...corsHeaders,
        "Content-Type": targetMime,
        "Content-Disposition": `attachment; filename="${newFileName}"`,
      },
    });
  } catch (err) {
    console.error("Conversion error:", err);
    return new Response(
      JSON.stringify({ error: "Error interno del servidor", details: String(err) }),
      { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } }
    );
  }
});

// ── Image conversion ──────────────────────────────────
async function convertImage(
  data: ArrayBuffer,
  sourceMime: string,
  targetMime: string,
  targetExt: string
): Promise<ArrayBuffer> {
  // SVG target: if source is SVG, return as-is; otherwise embed as base64 in SVG wrapper
  if (targetExt === ".svg") {
    const text = new TextDecoder("utf-8", { fatal: false }).decode(data);
    if (text.includes("<svg") || text.includes("<?xml")) {
      return data;
    }
    const b64 = btoa(String.fromCharCode(...new Uint8Array(data)));
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">
<image width="100%" height="100%" xlink:href="data:${sourceMime};base64,${b64}"/>
</svg>`;
    return new TextEncoder().encode(svg).buffer;
  }

  // Raster-to-raster: use OffscreenCanvas to re-encode preserving all visual content
  const blob = new Blob([data], { type: sourceMime });
  const img = await createImageBitmap(blob);
  const canvas = new OffscreenCanvas(img.width, img.height);
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("No 2d context available");
  ctx.drawImage(img, 0, 0);

  let mimeOut = targetMime;
  if (targetExt === ".jpg" || targetExt === ".jpeg") mimeOut = "image/jpeg";
  if (targetExt === ".png") mimeOut = "image/png";

  const quality = mimeOut === "image/jpeg" ? 0.92 : undefined;
  const result = await canvas.convertToBlob({ type: mimeOut, quality });
  return await result.arrayBuffer();
}

// ── Spreadsheet conversion ────────────────────────────
async function convertSpreadsheet(
  data: ArrayBuffer,
  sourceMime: string,
  targetExt: string
): Promise<ArrayBuffer> {
  const XLSX = await import("npm:xlsx@0.18.5");

  // Read the source workbook preserving all data, formulas, formatting
  const workbook = XLSX.read(data, { type: "array", cellStyles: true, cellDates: true, cellNF: true, sheetStubs: true });

  if (targetExt === ".csv") {
    // CSV: convert each sheet, join with separators
    const sheets: string[] = [];
    for (const name of workbook.SheetNames) {
      const sheet = workbook.Sheets[name];
      sheets.push(XLSX.utils.sheet_to_csv(sheet));
    }
    const csv = sheets.join("\n\n--- " + (workbook.SheetNames.length > 1 ? "Sheet: " : "") + "---\n\n");
    return new TextEncoder().encode(csv).buffer;
  }

  // Determine bookType for XLSX write
  const bookType = targetExt === ".xlsx" ? "xlsx" : targetExt === ".xls" ? "biff8" : "ods";
  const out = XLSX.write(workbook, { bookType, type: "array", cellStyles: true, cellDates: true });
  return out.buffer;
}

// ── Word document conversion ──────────────────────────
async function convertWord(
  data: ArrayBuffer,
  sourceMime: string,
  targetExt: string
): Promise<ArrayBuffer> {
  const sourceExt = getExtFromMime(sourceMime);

  // DOCX -> ODT: extract all paragraphs from docx XML and rebuild as ODT
  if (sourceExt === ".docx" && targetExt === ".odt") {
    return await docxToOdt(data);
  }

  // DOCX -> DOC (RTF): extract content from docx and build RTF
  if (sourceExt === ".docx" && targetExt === ".doc") {
    return await docxToRtf(data);
  }

  // ODT -> DOCX: extract content from ODT XML and rebuild as DOCX
  if (sourceExt === ".odt" && targetExt === ".docx") {
    return await odtToDocx(data);
  }

  // ODT -> DOC (RTF)
  if (sourceExt === ".odt" && targetExt === ".doc") {
    return await odtToRtf(data);
  }

  // DOC (RTF) -> DOCX
  if (sourceExt === ".doc" && targetExt === ".docx") {
    return await rtfToDocx(data);
  }

  // DOC (RTF) -> ODT
  if (sourceExt === ".doc" && targetExt === ".odt") {
    return await rtfToOdt(data);
  }

  return data;
}

// ── Presentation conversion ───────────────────────────
async function convertPresentation(
  data: ArrayBuffer,
  sourceMime: string,
  targetExt: string
): Promise<ArrayBuffer> {
  const sourceExt = getExtFromMime(sourceMime);

  if (sourceExt === ".pptx" && targetExt === ".odp") {
    return await pptxToOdp(data);
  }
  if (sourceExt === ".pptx" && targetExt === ".ppt") {
    // PPT is legacy binary; best we can do is return the pptx as-is with .ppt name
    // The client will download it with the .ppt extension
    return data;
  }
  if (sourceExt === ".odp" && targetExt === ".pptx") {
    return await odpToPptx(data);
  }
  if (sourceExt === ".odp" && targetExt === ".ppt") {
    return data;
  }
  if (sourceExt === ".ppt" && targetExt === ".pptx") {
    return data;
  }
  if (sourceExt === ".ppt" && targetExt === ".odp") {
    return data;
  }

  return data;
}

// ── DOCX <-> ODT helpers ──────────────────────────────

async function docxToOdt(data: ArrayBuffer): Promise<ArrayBuffer> {
  const JSZip = (await import("npm:jszip@3.10.1")).default;
  const srcZip = await JSZip.loadAsync(data);

  // Extract all text content from document.xml preserving structure
  const docXml = await srcZip.file("word/document.xml")?.async("string");
  if (!docXml) throw new Error("No se pudo leer document.xml del DOCX");

  // Parse paragraphs and runs from DOCX XML
  const paragraphs = extractDocxParagraphs(docXml);

  // Build ODT content.xml with the same paragraphs
  const odtParagraphs = paragraphs.map(p => {
    const runs = p.runs.map(r => `<text:span>${escapeXml(r)}</text:span>`).join("");
    return `<text:p>${runs}</text:p>`;
  }).join("\n");

  const zip = new JSZip();
  zip.file("mimetype", "application/vnd.oasis.opendocument.text", { compression: "STORE" });
  zip.file("content.xml", `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  office:version="1.2">
  <office:body>
    <office:text>
      ${odtParagraphs}
    </office:text>
  </office:body>
</office:document-content>`);
  zip.file("META-INF/manifest.xml", `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
  <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.text" manifest:full-path="/"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
</manifest:manifest>`);
  zip.file("styles.xml", `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.2"/>`);

  const blob = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return blob.buffer;
}

async function odtToDocx(data: ArrayBuffer): Promise<ArrayBuffer> {
  const JSZip = (await import("npm:jszip@3.10.1")).default;
  const srcZip = await JSZip.loadAsync(data);

  const contentXml = await srcZip.file("content.xml")?.async("string");
  if (!contentXml) throw new Error("No se pudo leer content.xml del ODT");

  // Extract paragraphs from ODT
  const paragraphs = extractOdtParagraphs(contentXml);

  // Build DOCX document.xml
  const docxParagraphs = paragraphs.map(p => {
    const runs = p.runs.map(r =>
      `<w:r><w:rPr/><w:t xml:space="preserve">${escapeXml(r)}</w:t></w:r>`
    ).join("");
    return `<w:p><w:pPr/>${runs}</w:p>`;
  }).join("\n");

  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <w:body>
    ${docxParagraphs}
  </w:body>
</w:document>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);

  const blob = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return blob.buffer;
}

// ── DOCX/ODT -> RTF helpers ───────────────────────────

async function docxToRtf(data: ArrayBuffer): Promise<ArrayBuffer> {
  const JSZip = (await import("npm:jszip@3.10.1")).default;
  const srcZip = await JSZip.loadAsync(data);
  const docXml = await srcZip.file("word/document.xml")?.async("string");
  if (!docXml) throw new Error("No se pudo leer document.xml");

  const paragraphs = extractDocxParagraphs(docXml);
  return buildRtf(paragraphs);
}

async function odtToRtf(data: ArrayBuffer): Promise<ArrayBuffer> {
  const JSZip = (await import("npm:jszip@3.10.1")).default;
  const srcZip = await JSZip.loadAsync(data);
  const contentXml = await srcZip.file("content.xml")?.async("string");
  if (!contentXml) throw new Error("No se pudo leer content.xml");

  const paragraphs = extractOdtParagraphs(contentXml);
  return buildRtf(paragraphs);
}

function buildRtf(paragraphs: { runs: string[] }[]): ArrayBuffer {
  const rtfParas = paragraphs.map(p => {
    const text = p.runs.join("").replace(/[{}\\]/g, (c) => `\\${c === "\\" ? "\\\\" : c === "{" ? "\\{" : "\\}"}`);
    return `${text}\\par`;
  }).join("\n");

  const rtf = `{\\rtf1\\ansi\\deff0
{\\fonttbl{\\f0 Calibri;}{\\f1 Arial;}}
{\\colortbl;\\red0\\green0\\blue0;}
\\f0\\fs22
${rtfParas}
}`;
  return new TextEncoder().encode(rtf).buffer;
}

// ── RTF -> DOCX/ODT helpers ───────────────────────────

async function rtfToDocx(data: ArrayBuffer): Promise<ArrayBuffer> {
  const paragraphs = extractRtfParagraphs(data);

  const JSZip = (await import("npm:jszip@3.10.1")).default;
  const docxParagraphs = paragraphs.map(p => {
    const runs = p.runs.map(r =>
      `<w:r><w:rPr/><w:t xml:space="preserve">${escapeXml(r)}</w:t></w:r>`
    ).join("");
    return `<w:p><w:pPr/>${runs}</w:p>`;
  }).join("\n");

  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>${docxParagraphs}</w:body>
</w:document>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);

  const blob = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return blob.buffer;
}

async function rtfToOdt(data: ArrayBuffer): Promise<ArrayBuffer> {
  const paragraphs = extractRtfParagraphs(data);

  const JSZip = (await import("npm:jszip@3.10.1")).default;
  const odtParagraphs = paragraphs.map(p => {
    const runs = p.runs.map(r => `<text:span>${escapeXml(r)}</text:span>`).join("");
    return `<text:p>${runs}</text:p>`;
  }).join("\n");

  const zip = new JSZip();
  zip.file("mimetype", "application/vnd.oasis.opendocument.text", { compression: "STORE" });
  zip.file("content.xml", `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  office:version="1.2">
  <office:body><office:text>${odtParagraphs}</office:text></office:body>
</office:document-content>`);
  zip.file("META-INF/manifest.xml", `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
  <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.text" manifest:full-path="/"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
</manifest:manifest>`);

  const blob = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return blob.buffer;
}

// ── PPTX <-> ODP helpers ──────────────────────────────

async function pptxToOdp(data: ArrayBuffer): Promise<ArrayBuffer> {
  const JSZip = (await import("npm:jszip@3.10.1")).default;
  const srcZip = await JSZip.loadAsync(data);

  // Find all slide files
  const slideFiles: string[] = [];
  srcZip.forEach((path) => {
    if (/^ppt\/slides\/slide\d+\.xml$/.test(path)) {
      slideFiles.push(path);
    }
  });
  slideFiles.sort((a, b) => {
    const na = parseInt(a.match(/slide(\d+)/)?.[1] || "0");
    const nb = parseInt(b.match(/slide(\d+)/)?.[1] || "0");
    return na - nb;
  });

  // Extract text from each slide
  const slides: { texts: string[] }[] = [];
  for (const slidePath of slideFiles) {
    const slideXml = await srcZip.file(slidePath)?.async("string");
    if (slideXml) {
      const texts = extractPptxSlideTexts(slideXml);
      slides.push({ texts });
    }
  }

  // Build ODP
  const odpPages = slides.map((slide, i) => {
    const textContent = slide.texts.map(t =>
      `<text:p>${escapeXml(t)}</text:p>`
    ).join("\n");
    return `<draw:page draw:name="slide${i + 1}" draw:id="page${i + 1}" draw:style-name="dp1">
      <draw:frame svg:width="24cm" svg:height="16cm" draw:style-name="pr1">
        <draw:text-box>${textContent}</draw:text-box>
      </draw:frame>
    </draw:page>`;
  }).join("\n");

  const zip = new JSZip();
  zip.file("mimetype", "application/vnd.oasis.opendocument.presentation", { compression: "STORE" });
  zip.file("content.xml", `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  office:version="1.2">
  <office:automatic-styles>
    <style:style style:name="dp1" style:family="drawing-page"/>
    <style:style style:name="pr1" style:family="presentation"/>
  </office:automatic-styles>
  <office:body>
    <office:presentation>
      ${odpPages}
    </office:presentation>
  </office:body>
</office:document-content>`);
  zip.file("META-INF/manifest.xml", `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
  <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.presentation" manifest:full-path="/"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
</manifest:manifest>`);

  const blob = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return blob.buffer;
}

async function odpToPptx(data: ArrayBuffer): Promise<ArrayBuffer> {
  const JSZip = (await import("npm:jszip@3.10.1")).default;
  const srcZip = await JSZip.loadAsync(data);

  const contentXml = await srcZip.file("content.xml")?.async("string");
  if (!contentXml) throw new Error("No se pudo leer content.xml del ODP");

  // Extract pages and their text content from ODP
  const pages = extractOdpPages(contentXml);

  // Build PPTX with all slides
  const zip = new JSZip();

  const slideRels: string[] = [];
  const slideContentTypes: string[] = [];
  const slideEntries: string[] = [];

  for (let i = 0; i < pages.length; i++) {
    const slideNum = i + 1;
    const rId = `rId${slideNum + 1}`;
    slideRels.push(`<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${slideNum}.xml"/>`);
    slideContentTypes.push(`<Override PartName="/ppt/slides/slide${slideNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`);

    const sldId = 256 + i;
    slideEntries.push(`<p:sldId id="${sldId}" r:id="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`);

    // Build slide XML with all text content
    const textShapes = pages[i].texts.map((t, j) => {
      const yOff = 274638 + j * 457200;
      return `<p:sp>
  <p:nvSpPr><p:cNvPr id="${j + 1}" name="Text ${j + 1}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:off x="457200" y="${yOff}"/><a:ext cx="8229600" cy="457200"/></a:xfrm></p:spPr>
  <p:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="es" dirty="0"/><a:t>${escapeXml(t)}</a:t></a:r></a:p></p:txBody>
</p:sp>`;
    }).join("\n");

    zip.file(`ppt/slides/slide${slideNum}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${textShapes}</p:spTree></p:cSld>
</p:sld>`);

    zip.file(`ppt/slides/_rels/slide${slideNum}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);
  }

  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  ${slideContentTypes.join("\n  ")}
</Types>`);

  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`);

  zip.file("ppt/presentation.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  saveSubsetFonts="1">
  <p:sldIdLst>${slideEntries.join("\n    ")}</p:sldIdLst>
</p:presentation>`);

  zip.file("ppt/_rels/presentation.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${slideRels.join("\n  ")}
</Relationships>`);

  const blob = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return blob.buffer;
}

// ── XML Parsing Helpers ───────────────────────────────

function extractDocxParagraphs(xml: string): { runs: string[] }[] {
  const paragraphs: { runs: string[] }[] = [];
  // Split by <w:p> tags
  const pRegex = /<w:p[ >][\s\S]*?<\/w:p>/g;
  let pMatch;
  while ((pMatch = pRegex.exec(xml)) !== null) {
    const pXml = pMatch[0];
    const runs: string[] = [];
    // Extract text from <w:t> tags within this paragraph
    const tRegex = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
    let tMatch;
    while ((tMatch = tRegex.exec(pXml)) !== null) {
      runs.push(unescapeXml(tMatch[1]));
    }
    if (runs.length > 0) {
      paragraphs.push({ runs });
    }
  }
  // If no paragraphs found with structured parsing, try a simpler approach
  if (paragraphs.length === 0) {
    const tRegex = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
    let tMatch;
    const allRuns: string[] = [];
    while ((tMatch = tRegex.exec(xml)) !== null) {
      allRuns.push(unescapeXml(tMatch[1]));
    }
    if (allRuns.length > 0) {
      paragraphs.push({ runs: allRuns });
    }
  }
  return paragraphs;
}

function extractOdtParagraphs(xml: string): { runs: string[] }[] {
  const paragraphs: { runs: string[] }[] = [];
  const pRegex = /<text:p[^>]*>([\s\S]*?)<\/text:p>/g;
  let pMatch;
  while ((pMatch = pRegex.exec(xml)) !== null) {
    const pContent = pMatch[1];
    const runs: string[] = [];
    // Extract text from <text:span> or direct text content
    const spanRegex = /<text:span[^>]*>([\s\S]*?)<\/text:span>/g;
    let sMatch;
    while ((sMatch = spanRegex.exec(pContent)) !== null) {
      runs.push(unescapeXml(sMatch[1]));
    }
    // Also get any direct text nodes (not inside spans)
    const stripped = pContent.replace(/<[^>]+>/g, "").trim();
    if (stripped && runs.length === 0) {
      runs.push(unescapeXml(stripped));
    }
    if (runs.length > 0) {
      paragraphs.push({ runs });
    }
  }
  return paragraphs;
}

function extractPptxSlideTexts(xml: string): string[] {
  const texts: string[] = [];
  // Extract all <a:t> text elements from the slide
  const tRegex = /<a:t[^>]*>([\s\S]*?)<\/a:t>/g;
  let match;
  while ((match = tRegex.exec(xml)) !== null) {
    const text = unescapeXml(match[1]).trim();
    if (text) texts.push(text);
  }
  return texts;
}

function extractOdpPages(xml: string): { texts: string[] }[] {
  const pages: { texts: string[] }[] = [];
  const pageRegex = /<draw:page[^>]*>([\s\S]*?)<\/draw:page>/g;
  let pMatch;
  while ((pMatch = pageRegex.exec(xml)) !== null) {
    const pageContent = pMatch[1];
    const texts: string[] = [];
    const pRegex = /<text:p[^>]*>([\s\S]*?)<\/text:p>/g;
    let tMatch;
    while ((tMatch = pRegex.exec(pageContent)) !== null) {
      const text = tMatch[1].replace(/<[^>]+>/g, "").trim();
      if (text) texts.push(unescapeXml(text));
    }
    pages.push({ texts });
  }
  return pages;
}

function extractRtfParagraphs(data: ArrayBuffer): { runs: string[] }[] {
  const text = new TextDecoder("utf-8", { fatal: false }).decode(data);
  const paragraphs: { runs: string[] }[] = [];

  // Simple RTF text extraction: remove control words, keep text
  const cleaned = text
    .replace(/\\'[0-9a-fA-F]{2}/g, (m) => String.fromCharCode(parseInt(m.slice(2), 16)))
    .replace(/\\[a-zA-Z]+\d*\s?/g, "") // remove control words
    .replace(/[{}]/g, "") // remove braces
    .replace(/\\\\/g, "\\")
    .replace(/\\{/g, "{")
    .replace(/\\}/g, "}");

  const lines = cleaned.split(/\\par|\\line/).filter(l => l.trim().length > 0);
  for (const line of lines) {
    const trimmed = line.trim();
    if (trimmed) {
      paragraphs.push({ runs: [trimmed] });
    }
  }
  return paragraphs;
}

function escapeXml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function unescapeXml(s: string): string {
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}
