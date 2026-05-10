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

// Extra MIME aliases browsers may send
const MIME_ALIASES: Record<string, string> = {
  "text/plain": ".csv",          // browsers often send CSV as text/plain
  "application/octet-stream": "", // fallback — rely on extension
  "application/x-msdownload": ".xls",
  "application/vnd.ms-office": ".xls",
};

function getExtFromMime(mime: string): string {
  for (const cat of Object.values(CONVERT_MIME)) {
    for (const [ext, m] of Object.entries(cat)) {
      if (m === mime) return ext;
    }
  }
  return MIME_ALIASES[mime] ?? "";
}

function getCategoryForExt(ext: string): string | null {
  const e = ext.startsWith(".") ? ext : "." + ext;
  for (const [cat, formats] of Object.entries(CONVERT_MIME)) {
    if (formats[e] !== undefined) return cat;
  }
  return null;
}

function resolveSource(file: File): { sourceExt: string; sourceCat: string | null } {
  const nameExt = "." + file.name.split(".").pop()!.toLowerCase();
  // Try MIME first (but skip aliases that would override a valid extension)
  const mimeExt = getExtFromMime(file.type);
  const catFromMime = mimeExt ? getCategoryForExt(mimeExt) : null;
  const catFromExt = getCategoryForExt(nameExt);

  if (catFromExt) return { sourceExt: nameExt, sourceCat: catFromExt };
  if (catFromMime) return { sourceExt: mimeExt, sourceCat: catFromMime };
  return { sourceExt: nameExt, sourceCat: null };
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
    const { sourceExt, sourceCat } = resolveSource(file);
    const targetCat = getCategoryForExt(targetExt);

    if (!sourceCat || !targetCat) {
      return new Response(
        JSON.stringify({ error: `Formato no reconocido: origen=${sourceExt} destino=${targetExt}` }),
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
      converted = await convertImage(arrayBuffer, sourceExt, targetExt);
    } else if (sourceCat === "excel") {
      converted = await convertSpreadsheet(arrayBuffer, sourceExt, targetExt);
    } else if (sourceCat === "word") {
      converted = await convertWord(arrayBuffer, sourceExt, targetExt);
    } else if (sourceCat === "powerpoint") {
      converted = await convertPresentation(arrayBuffer, sourceExt, targetExt);
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
// Uses npm:jimp (pure JS) since OffscreenCanvas is not available in Deno
async function convertImage(
  data: ArrayBuffer,
  sourceExt: string,
  targetExt: string
): Promise<ArrayBuffer> {
  // SVG passthrough: we can't rasterise SVG without native libs, return as-is
  if (sourceExt === ".svg") {
    if (targetExt === ".svg") return data;
    // Wrap SVG as-is — proper rasterisation would need a native renderer
    return data;
  }

  // SVG target: embed source image as base64 data URI inside an SVG
  if (targetExt === ".svg") {
    const mimeIn = sourceExt === ".png" ? "image/png" : "image/jpeg";
    const b64 = btoa(String.fromCharCode(...new Uint8Array(data)));
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"><image width="100%" height="100%" xlink:href="data:${mimeIn};base64,${b64}"/></svg>`;
    return new TextEncoder().encode(svg).buffer;
  }

  // Raster-to-raster using jimp
  const Jimp = (await import("npm:jimp@0.22.12")).Jimp;

  const uint8 = new Uint8Array(data);
  const image = await Jimp.fromBuffer(uint8);

  let outMime: "image/jpeg" | "image/png" = "image/jpeg";
  if (targetExt === ".png") outMime = "image/png";

  const outBuffer = await image.getBuffer(outMime);
  return outBuffer.buffer;
}

// ── Spreadsheet conversion ────────────────────────────
async function convertSpreadsheet(
  data: ArrayBuffer,
  sourceExt: string,
  targetExt: string
): Promise<ArrayBuffer> {
  const XLSX = await import("npm:xlsx@0.18.5");

  const workbook = XLSX.read(new Uint8Array(data), {
    type: "array",
    cellStyles: true,
    cellDates: true,
    cellNF: true,
    sheetStubs: true,
  });

  if (targetExt === ".csv") {
    const sheets: string[] = [];
    for (const name of workbook.SheetNames) {
      const sheet = workbook.Sheets[name];
      sheets.push(`--- ${name} ---\n${XLSX.utils.sheet_to_csv(sheet)}`);
    }
    return new TextEncoder().encode(sheets.join("\n\n")).buffer;
  }

  // Map extension to xlsx bookType
  const bookTypeMap: Record<string, string> = {
    ".xlsx": "xlsx",
    ".xls": "xls",
    ".ods": "ods",
    ".csv": "csv",
  };
  const bookType = bookTypeMap[targetExt] ?? "xlsx";

  const out = XLSX.write(workbook, {
    bookType: bookType as Parameters<typeof XLSX.write>[1]["bookType"],
    type: "array",
    cellStyles: true,
    cellDates: true,
  });

  return (out as Uint8Array).buffer;
}

// ── Word document conversion ──────────────────────────
async function convertWord(
  data: ArrayBuffer,
  sourceExt: string,
  targetExt: string
): Promise<ArrayBuffer> {
  if (sourceExt === ".docx" && targetExt === ".odt") return await docxToOdt(data);
  if (sourceExt === ".docx" && targetExt === ".doc") return await docxToRtf(data);
  if (sourceExt === ".odt" && targetExt === ".docx") return await odtToDocx(data);
  if (sourceExt === ".odt" && targetExt === ".doc") return await odtToRtf(data);
  if (sourceExt === ".doc" && targetExt === ".docx") return await rtfToDocx(data);
  if (sourceExt === ".doc" && targetExt === ".odt") return await rtfToOdt(data);
  return data;
}

// ── Presentation conversion ───────────────────────────
async function convertPresentation(
  data: ArrayBuffer,
  sourceExt: string,
  targetExt: string
): Promise<ArrayBuffer> {
  if (sourceExt === ".pptx" && targetExt === ".odp") return await pptxToOdp(data);
  if (sourceExt === ".odp" && targetExt === ".pptx") return await odpToPptx(data);
  // Legacy .ppt binary format — pass through with new extension
  return data;
}

// ── DOCX <-> ODT ──────────────────────────────────────

async function docxToOdt(data: ArrayBuffer): Promise<ArrayBuffer> {
  const JSZip = (await import("npm:jszip@3.10.1")).default;
  const srcZip = await JSZip.loadAsync(data);
  const docXml = await srcZip.file("word/document.xml")?.async("string");
  if (!docXml) throw new Error("No se pudo leer document.xml del DOCX");

  const paragraphs = extractDocxParagraphs(docXml);
  const odtParagraphs = paragraphs
    .map((p) => `<text:p>${p.runs.map((r) => `<text:span>${escapeXml(r)}</text:span>`).join("")}</text:p>`)
    .join("\n");

  const zip = new JSZip();
  zip.file("mimetype", "application/vnd.oasis.opendocument.text", { compression: "STORE" });
  zip.file("content.xml", buildOdtContent(odtParagraphs));
  zip.file("META-INF/manifest.xml", buildOdtManifest("application/vnd.oasis.opendocument.text"));
  zip.file("styles.xml", `<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.2"/>`);
  return ((await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" })) as Uint8Array).buffer;
}

async function odtToDocx(data: ArrayBuffer): Promise<ArrayBuffer> {
  const JSZip = (await import("npm:jszip@3.10.1")).default;
  const srcZip = await JSZip.loadAsync(data);
  const contentXml = await srcZip.file("content.xml")?.async("string");
  if (!contentXml) throw new Error("No se pudo leer content.xml del ODT");

  const paragraphs = extractOdtParagraphs(contentXml);
  const docxParagraphs = paragraphs
    .map((p) =>
      `<w:p><w:pPr/>${p.runs.map((r) => `<w:r><w:rPr/><w:t xml:space="preserve">${escapeXml(r)}</w:t></w:r>`).join("")}</w:p>`
    )
    .join("\n");

  return buildDocx(docxParagraphs);
}

// ── DOCX/ODT -> RTF ───────────────────────────────────

async function docxToRtf(data: ArrayBuffer): Promise<ArrayBuffer> {
  const JSZip = (await import("npm:jszip@3.10.1")).default;
  const srcZip = await JSZip.loadAsync(data);
  const docXml = await srcZip.file("word/document.xml")?.async("string");
  if (!docXml) throw new Error("No se pudo leer document.xml");
  return buildRtf(extractDocxParagraphs(docXml));
}

async function odtToRtf(data: ArrayBuffer): Promise<ArrayBuffer> {
  const JSZip = (await import("npm:jszip@3.10.1")).default;
  const srcZip = await JSZip.loadAsync(data);
  const contentXml = await srcZip.file("content.xml")?.async("string");
  if (!contentXml) throw new Error("No se pudo leer content.xml");
  return buildRtf(extractOdtParagraphs(contentXml));
}

function buildRtf(paragraphs: { runs: string[] }[]): ArrayBuffer {
  const rtfParas = paragraphs
    .map((p) => {
      const text = p.runs
        .join("")
        .replace(/\\/g, "\\\\")
        .replace(/\{/g, "\\{")
        .replace(/\}/g, "\\}");
      return `${text}\\par`;
    })
    .join("\n");
  const rtf = `{\\rtf1\\ansi\\deff0\n{\\fonttbl{\\f0 Calibri;}}\n{\\colortbl;\\red0\\green0\\blue0;}\n\\f0\\fs22\n${rtfParas}\n}`;
  return new TextEncoder().encode(rtf).buffer;
}

// ── RTF -> DOCX/ODT ───────────────────────────────────

async function rtfToDocx(data: ArrayBuffer): Promise<ArrayBuffer> {
  const paragraphs = extractRtfParagraphs(data);
  const docxParagraphs = paragraphs
    .map((p) =>
      `<w:p><w:pPr/>${p.runs.map((r) => `<w:r><w:rPr/><w:t xml:space="preserve">${escapeXml(r)}</w:t></w:r>`).join("")}</w:p>`
    )
    .join("\n");
  return buildDocx(docxParagraphs);
}

async function rtfToOdt(data: ArrayBuffer): Promise<ArrayBuffer> {
  const JSZip = (await import("npm:jszip@3.10.1")).default;
  const paragraphs = extractRtfParagraphs(data);
  const odtParagraphs = paragraphs
    .map((p) => `<text:p>${p.runs.map((r) => `<text:span>${escapeXml(r)}</text:span>`).join("")}</text:p>`)
    .join("\n");

  const zip = new JSZip();
  zip.file("mimetype", "application/vnd.oasis.opendocument.text", { compression: "STORE" });
  zip.file("content.xml", buildOdtContent(odtParagraphs));
  zip.file("META-INF/manifest.xml", buildOdtManifest("application/vnd.oasis.opendocument.text"));
  return ((await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" })) as Uint8Array).buffer;
}

// ── PPTX <-> ODP ──────────────────────────────────────

async function pptxToOdp(data: ArrayBuffer): Promise<ArrayBuffer> {
  const JSZip = (await import("npm:jszip@3.10.1")).default;
  const srcZip = await JSZip.loadAsync(data);

  const slideFiles: string[] = [];
  srcZip.forEach((path) => {
    if (/^ppt\/slides\/slide\d+\.xml$/.test(path)) slideFiles.push(path);
  });
  slideFiles.sort((a, b) => {
    const na = parseInt(a.match(/slide(\d+)/)?.[1] || "0");
    const nb = parseInt(b.match(/slide(\d+)/)?.[1] || "0");
    return na - nb;
  });

  const slides: { texts: string[] }[] = [];
  for (const slidePath of slideFiles) {
    const slideXml = await srcZip.file(slidePath)?.async("string");
    if (slideXml) slides.push({ texts: extractPptxSlideTexts(slideXml) });
  }

  const odpPages = slides
    .map((slide, i) => {
      const textContent = slide.texts.map((t) => `<text:p>${escapeXml(t)}</text:p>`).join("\n");
      return `<draw:page draw:name="slide${i + 1}" draw:id="page${i + 1}" draw:style-name="dp1">
      <draw:frame svg:width="24cm" svg:height="16cm" draw:style-name="pr1">
        <draw:text-box>${textContent}</draw:text-box>
      </draw:frame>
    </draw:page>`;
    })
    .join("\n");

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
  <office:body><office:presentation>${odpPages}</office:presentation></office:body>
</office:document-content>`);
  zip.file("META-INF/manifest.xml", buildOdtManifest("application/vnd.oasis.opendocument.presentation", "content.xml"));
  return ((await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" })) as Uint8Array).buffer;
}

async function odpToPptx(data: ArrayBuffer): Promise<ArrayBuffer> {
  const JSZip = (await import("npm:jszip@3.10.1")).default;
  const srcZip = await JSZip.loadAsync(data);
  const contentXml = await srcZip.file("content.xml")?.async("string");
  if (!contentXml) throw new Error("No se pudo leer content.xml del ODP");

  const pages = extractOdpPages(contentXml);
  const zip = new JSZip();

  const slideRels: string[] = [];
  const slideContentTypes: string[] = [];
  const slideEntries: string[] = [];

  for (let i = 0; i < pages.length; i++) {
    const slideNum = i + 1;
    const rId = `rId${slideNum + 1}`;
    slideRels.push(`<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${slideNum}.xml"/>`);
    slideContentTypes.push(`<Override PartName="/ppt/slides/slide${slideNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`);
    slideEntries.push(`<p:sldId id="${256 + i}" r:id="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`);

    const textShapes = pages[i].texts
      .map((t, j) => {
        const yOff = 274638 + j * 457200;
        return `<p:sp>
  <p:nvSpPr><p:cNvPr id="${j + 2}" name="Text ${j + 1}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:off x="457200" y="${yOff}"/><a:ext cx="8229600" cy="457200"/></a:xfrm></p:spPr>
  <p:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="es" dirty="0"/><a:t>${escapeXml(t)}</a:t></a:r></a:p></p:txBody>
</p:sp>`;
      })
      .join("\n");

    zip.file(`ppt/slides/slide${slideNum}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${textShapes}</p:spTree></p:cSld>
</p:sld>`);
    zip.file(`ppt/slides/_rels/slide${slideNum}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);
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
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" saveSubsetFonts="1">
  <p:sldIdLst>${slideEntries.join("\n    ")}</p:sldIdLst>
</p:presentation>`);

  zip.file("ppt/_rels/presentation.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${slideRels.join("\n  ")}
</Relationships>`);

  return ((await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" })) as Uint8Array).buffer;
}

// ── Shared builders ───────────────────────────────────

async function buildDocx(docxParagraphs: string): Promise<ArrayBuffer> {
  const JSZip = (await import("npm:jszip@3.10.1")).default;
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
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);
  return ((await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" })) as Uint8Array).buffer;
}

function buildOdtContent(bodyContent: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  office:version="1.2">
  <office:body><office:text>${bodyContent}</office:text></office:body>
</office:document-content>`;
}

function buildOdtManifest(mediaType: string, extraFile = "content.xml"): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
  <manifest:file-entry manifest:media-type="${mediaType}" manifest:full-path="/"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="${extraFile}"/>
</manifest:manifest>`;
}

// ── XML Parsing Helpers ───────────────────────────────

function extractDocxParagraphs(xml: string): { runs: string[] }[] {
  const paragraphs: { runs: string[] }[] = [];
  const pRegex = /<w:p[ >][\s\S]*?<\/w:p>/g;
  let pMatch;
  while ((pMatch = pRegex.exec(xml)) !== null) {
    const pXml = pMatch[0];
    const runs: string[] = [];
    const tRegex = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
    let tMatch;
    while ((tMatch = tRegex.exec(pXml)) !== null) {
      runs.push(unescapeXml(tMatch[1]));
    }
    if (runs.length > 0) paragraphs.push({ runs });
  }
  if (paragraphs.length === 0) {
    const allRuns: string[] = [];
    const tRegex = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
    let tMatch;
    while ((tMatch = tRegex.exec(xml)) !== null) allRuns.push(unescapeXml(tMatch[1]));
    if (allRuns.length > 0) paragraphs.push({ runs: allRuns });
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
    const spanRegex = /<text:span[^>]*>([\s\S]*?)<\/text:span>/g;
    let sMatch;
    while ((sMatch = spanRegex.exec(pContent)) !== null) runs.push(unescapeXml(sMatch[1]));
    if (runs.length === 0) {
      const stripped = pContent.replace(/<[^>]+>/g, "").trim();
      if (stripped) runs.push(unescapeXml(stripped));
    }
    if (runs.length > 0) paragraphs.push({ runs });
  }
  return paragraphs;
}

function extractPptxSlideTexts(xml: string): string[] {
  const texts: string[] = [];
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
    const texts: string[] = [];
    const pRegex = /<text:p[^>]*>([\s\S]*?)<\/text:p>/g;
    let tMatch;
    while ((tMatch = pRegex.exec(pMatch[1])) !== null) {
      const text = tMatch[1].replace(/<[^>]+>/g, "").trim();
      if (text) texts.push(unescapeXml(text));
    }
    pages.push({ texts });
  }
  return pages;
}

function extractRtfParagraphs(data: ArrayBuffer): { runs: string[] }[] {
  const text = new TextDecoder("utf-8", { fatal: false }).decode(data);
  const cleaned = text
    .replace(/\\'[0-9a-fA-F]{2}/g, (m) => String.fromCharCode(parseInt(m.slice(2), 16)))
    .replace(/\\[a-zA-Z]+\d*[ ]?/g, "")
    .replace(/[{}]/g, "")
    .replace(/\\\\/g, "\\")
    .replace(/\\{/g, "{")
    .replace(/\\}/g, "}");
  return cleaned
    .split(/\\par|\\line/)
    .map((l) => l.trim())
    .filter(Boolean)
    .map((l) => ({ runs: [l] }));
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
