import "jsr:@supabase/functions-js/edge-runtime.d.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Client-Info, Apikey",
};

const CONVERT_MIME: Record<string, Record<string, string>> = {
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

const MIME_ALIASES: Record<string, string> = {
  "text/plain": ".csv",
  "application/octet-stream": "",
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

function getCategory(ext: string): string | null {
  const e = ext.startsWith(".") ? ext : "." + ext;
  for (const [cat, formats] of Object.entries(CONVERT_MIME)) {
    if (formats[e] !== undefined) return cat;
  }
  return null;
}

function resolveSource(file: File): { sourceExt: string; sourceCat: string | null } {
  const nameExt = "." + file.name.split(".").pop()!.toLowerCase();
  const mimeExt = getExtFromMime(file.type);
  const catFromExt = getCategory(nameExt);
  const catFromMime = mimeExt ? getCategory(mimeExt) : null;
  if (catFromExt) return { sourceExt: nameExt, sourceCat: catFromExt };
  if (catFromMime) return { sourceExt: mimeExt, sourceCat: catFromMime };
  return { sourceExt: nameExt, sourceCat: null };
}

// ── CloudConvert API ───────────────────────────────────
// Uses LibreOffice internally for full-fidelity conversions
// Preserves: images, tables, styles, headers/footers, footnotes, etc.

const CC_API = "https://api.cloudconvert.com/v2";

async function cloudConvert(data: ArrayBuffer, sourceExt: string, targetExt: string, fileName: string): Promise<ArrayBuffer> {
  const apiKey = Deno.env.get("CLOUDCONVERT_API_KEY");
  if (!apiKey) throw new Error("CLOUDCONVERT_API_KEY no configurada");

  const srcFormat = sourceExt.replace(".", "");
  const tgtFormat = targetExt.replace(".", "");

  // Step 1: Create upload+convert+export job
  const jobRes = await fetch(`${CC_API}/jobs`, {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      tasks: {
        "import-upload": { operation: "import/upload" },
        "convert-file": {
          operation: "convert",
          input: "import-upload",
          input_format: srcFormat,
          output_format: tgtFormat,
        },
        "export-url": {
          operation: "export/url",
          input: "convert-file",
        },
      },
    }),
  });

  if (!jobRes.ok) {
    const err = await jobRes.text();
    throw new Error(`CloudConvert job creation failed: ${err}`);
  }

  const job = await jobRes.json();
  const uploadTask = job.data.tasks.find((t: any) => t.name === "import-upload");
  if (!uploadTask) throw new Error("No upload task in CloudConvert job");

  // Step 2: Upload the file
  const uploadUrl = uploadTask.result.form.url;
  const uploadFields = uploadTask.result.form.parameters;

  const formData = new FormData();
  for (const [key, value] of Object.entries(uploadFields)) {
    formData.append(key, String(value));
  }
  formData.append("file", new Blob([data]), fileName);

  const uploadRes = await fetch(uploadUrl, { method: "POST", body: formData });
  if (!uploadRes.ok) {
    throw new Error(`CloudConvert upload failed: ${uploadRes.status}`);
  }

  // Step 3: Wait for job to complete
  const jobId = job.data.id;
  let attempts = 0;
  const maxAttempts = 60;

  while (attempts < maxAttempts) {
    await new Promise((r) => setTimeout(r, 1000));
    const statusRes = await fetch(`${CC_API}/jobs/${jobId}`, {
      headers: { "Authorization": `Bearer ${apiKey}` },
    });
    const status = await statusRes.json();
    const jobStatus = status.data.status;

    if (jobStatus === "finished") break;
    if (jobStatus === "error") {
      const failedTask = status.data.tasks.find((t: any) => t.status === "error");
      throw new Error(`CloudConvert conversion failed: ${failedTask?.result?.message || "unknown error"}`);
    }
    attempts++;
  }

  if (attempts >= maxAttempts) throw new Error("CloudConvert conversion timed out");

  // Step 4: Get the result
  const finalRes = await fetch(`${CC_API}/jobs/${jobId}`, {
    headers: { "Authorization": `Bearer ${apiKey}` },
  });
  const finalJob = await finalRes.json();
  const exportTask = finalJob.data.tasks.find((t: any) => t.name === "export-url");
  if (!exportTask?.result?.files?.[0]?.url) {
    throw new Error("CloudConvert: no export URL in result");
  }

  const downloadUrl = exportTask.result.files[0].url;
  const fileRes = await fetch(downloadUrl);
  if (!fileRes.ok) throw new Error("CloudConvert: failed to download result");

  return await fileRes.arrayBuffer();
}

// ── Main handler ───────────────────────────────────────

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
    const targetCat = getCategory(targetExt);

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

    if (sourceExt === targetExt) {
      return new Response(await file.arrayBuffer(), {
        status: 200,
        headers: { ...corsHeaders, "Content-Type": file.type || "application/octet-stream", "Content-Disposition": `attachment; filename="${file.name}"` },
      });
    }

    const arrayBuffer = await file.arrayBuffer();
    const baseName = file.name.replace(/\.[^.]+$/, "");
    const newFileName = `${baseName}${targetExt}`;
    const targetMime = CONVERT_MIME[targetCat]?.[targetExt] || "application/octet-stream";

    let converted: ArrayBuffer;

    if (sourceCat === "pictures") {
      converted = await convertImage(arrayBuffer, sourceExt, targetExt);
    } else if (sourceCat === "excel") {
      converted = await convertSpreadsheet(arrayBuffer, sourceExt, targetExt);
    } else {
      // Word and PowerPoint: use CloudConvert for full fidelity
      // Falls back to local conversion if CloudConvert is not configured
      converted = await convertDocument(arrayBuffer, sourceExt, targetExt, file.name);
    }

    return new Response(converted, {
      status: 200,
      headers: { ...corsHeaders, "Content-Type": targetMime, "Content-Disposition": `attachment; filename="${newFileName}"` },
    });
  } catch (err) {
    console.error("Conversion error:", err);
    const msg = err instanceof Error ? err.message : String(err);
    return new Response(
      JSON.stringify({ error: "Error interno del servidor", details: msg }),
      { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } }
    );
  }
});

// ════════════════════════════════════════════════════════
// DOCUMENT CONVERSIONS (Word + PowerPoint)
// CloudConvert primary, local fallback
// ════════════════════════════════════════════════════════

async function convertDocument(data: ArrayBuffer, sourceExt: string, targetExt: string, fileName: string): Promise<ArrayBuffer> {
  const hasCloudConvert = !!Deno.env.get("CLOUDCONVERT_API_KEY");

  if (hasCloudConvert) {
    try {
      return await cloudConvert(data, sourceExt, targetExt, fileName);
    } catch (ccErr) {
      console.error("CloudConvert failed, falling back to local:", ccErr);
    }
  }

  // Local fallback
  const cat = getCategory(sourceExt);
  if (cat === "word") return await convertWordLocal(data, sourceExt, targetExt);
  if (cat === "powerpoint") return await convertPresentationLocal(data, sourceExt, targetExt);
  return data;
}

// ── Local Word conversion (fallback) ───────────────────

async function convertWordLocal(data: ArrayBuffer, sourceExt: string, targetExt: string): Promise<ArrayBuffer> {
  // DOCX -> ODT: odf-kit preserves most structure
  if (sourceExt === ".docx" && targetExt === ".odt") {
    return await docxToOdtDirect(data);
  }
  // All other paths go through HTML hub (loses some formatting)
  const html = await wordToHtml(data, sourceExt);
  return await htmlToWord(html, targetExt);
}

async function wordToHtml(data: ArrayBuffer, sourceExt: string): Promise<string> {
  switch (sourceExt) {
    case ".docx": return await docxToHtml(data);
    case ".odt": return await odtToHtml(data);
    case ".doc": return await docToHtml(data);
    default: throw new Error(`Formato no soportado: ${sourceExt}`);
  }
}

async function htmlToWord(html: string, targetExt: string): Promise<ArrayBuffer> {
  switch (targetExt) {
    case ".docx": return await htmlToDocx(html);
    case ".odt": return await htmlToOdt(html);
    case ".doc": return await htmlToRtf(html);
    default: throw new Error(`Formato no soportado: ${targetExt}`);
  }
}

// ── DOCX -> ODT via odf-kit ────────────────────────────

async function docxToOdtDirect(data: ArrayBuffer): Promise<ArrayBuffer> {
  const odfKit = await import("npm:odf-kit@0.13.1");
  const { bytes } = await odfKit.docxToOdt(new Uint8Array(data));
  return bytes.buffer;
}

// ── DOCX -> HTML via mammoth ────────────────────────────

async function docxToHtml(data: ArrayBuffer): Promise<string> {
  const mammoth = await import("npm:mammoth@1.8.0");
  const result = await mammoth.convertToHtml({ buffer: new Uint8Array(data) });
  return result.value;
}

// ── ODT -> HTML via JSZip ──────────────────────────────

async function odtToHtml(data: ArrayBuffer): Promise<string> {
  const JSZip = (await import("npm:jszip@3.10.1")).default;
  const zip = await JSZip.loadAsync(data);
  const contentXml = await zip.file("content.xml")?.async("string");
  if (!contentXml) throw new Error("No se pudo leer content.xml del ODT");
  return odtContentXmlToHtml(contentXml);
}

function odtContentXmlToHtml(xml: string): string {
  const blocks: string[] = [];
  const pRegex = /<text:(p|h)[^>]*>([\s\S]*?)<\/text:(?:p|h)>/g;
  let match;
  while ((match = pRegex.exec(xml)) !== null) {
    const tag = match[1];
    const content = match[2];
    const text = odtInlineToHtml(content);
    if (!text.trim()) continue;
    if (tag === "h") blocks.push(`<h2>${text}</h2>`);
    else blocks.push(`<p>${text}</p>`);
  }
  const liRegex = /<text:list-item[^>]*>([\s\S]*?)<\/text:list-item>/g;
  let liMatch;
  const listItems: string[] = [];
  while ((liMatch = liRegex.exec(xml)) !== null) {
    const inner = liMatch[1];
    const pInLi = inner.match(/<text:p[^>]*>([\s\S]*?)<\/text:p>/);
    if (pInLi) {
      const text = odtInlineToHtml(pInLi[1]);
      if (text.trim()) listItems.push(`<li>${text}</li>`);
    }
  }
  if (listItems.length > 0) blocks.push(`<ul>${listItems.join("")}</ul>`);
  if (blocks.length === 0) {
    const allText = xml.replace(/<[^>]+>/g, "").trim();
    if (allText) blocks.push(`<p>${escapeHtml(allText)}</p>`);
  }
  return blocks.join("\n");
}

function odtInlineToHtml(content: string): string {
  let result = content;
  result = result.replace(/<text:span[^>]*text:style-name="([^"]*)"[^>]*>([\s\S]*?)<\/text:span>/g, (_, _styleName, inner) => inner.replace(/<[^>]+>/g, ""));
  result = result.replace(/<text:span[^>]*>([\s\S]*?)<\/text:span>/g, (_, inner) => inner.replace(/<[^>]+>/g, ""));
  result = result.replace(/<text:[^/][^>]*>/g, "");
  result = result.replace(/<\/text:[^>]+>/g, "");
  result = result.replace(/<[^>]+>/g, "");
  return decodeHtmlEntities(result);
}

// ── DOC/RTF -> HTML ─────────────────────────────────────

async function docToHtml(data: ArrayBuffer): Promise<string> {
  const paragraphs = parseRtfToStyledParagraphs(data);
  return styledParagraphsToHtml(paragraphs);
}

function styledParagraphsToHtml(paragraphs: StyledParagraph[]): string {
  return paragraphs.map((p) => {
    const runsHtml = p.runs.map((run) => {
      let text = escapeHtml(run.text);
      if (run.bold) text = `<strong>${text}</strong>`;
      if (run.italic) text = `<em>${text}</em>`;
      if (run.underline) text = `<u>${text}</u>`;
      if (run.strike) text = `<s>${text}</s>`;
      return text;
    }).join("");
    if (p.headingLevel) return `<h${p.headingLevel}>${runsHtml}</h${p.headingLevel}>`;
    return `<p>${runsHtml}</p>`;
  }).join("\n");
}

// ── HTML -> DOCX via docx library ───────────────────────

async function htmlToDocx(html: string): Promise<ArrayBuffer> {
  const docxLib = await import("npm:docx@9.6.1");
  const { Document, Packer, Paragraph, TextRun, HeadingLevel, ExternalHyperlink, UnderlineType } = docxLib;
  const children: any[] = [];
  const blocks = parseHtmlBlocks(html);
  for (const block of blocks) {
    if (block.type === "heading" && block.level) {
      const headingMap: Record<number, any> = {
        1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3,
        4: HeadingLevel.HEADING_4, 5: HeadingLevel.HEADING_5, 6: HeadingLevel.HEADING_6,
      };
      children.push(new Paragraph({ heading: headingMap[block.level] || HeadingLevel.HEADING_1, children: htmlInlineToDocxRuns(block.content, docxLib) }));
    } else if (block.type === "listitem") {
      children.push(new Paragraph({ children: [new TextRun({ text: "\u2022 ", bold: true }), ...htmlInlineToDocxRuns(block.content, docxLib)] }));
    } else {
      const runs = htmlInlineToDocxRuns(block.content, docxLib);
      children.push(new Paragraph({ children: runs.length > 0 ? runs : [new TextRun({ text: "" })] }));
    }
  }
  const doc = new Document({ sections: [{ children }] });
  const buffer = await Packer.toBuffer(doc);
  return buffer.buffer;
}

function htmlInlineToDocxRuns(html: string, docxLib: any): any[] {
  const { TextRun, ExternalHyperlink, UnderlineType } = docxLib;
  const runs: any[] = [];
  const tokenRegex = /<(strong|b|em|i|u|s|a|span)[^>]*>([\s\S]*?)<\/\1>|([^<]+)/gi;
  let match;
  while ((match = tokenRegex.exec(html)) !== null) {
    if (match[3] !== undefined) {
      const text = decodeHtmlEntities(match[3]);
      if (text) runs.push(new TextRun({ text }));
      continue;
    }
    const tag = match[1].toLowerCase();
    const text = decodeHtmlEntities(match[2].replace(/<[^>]+>/g, ""));
    if (!text) continue;
    if (tag === "strong" || tag === "b") runs.push(new TextRun({ text, bold: true }));
    else if (tag === "em" || tag === "i") runs.push(new TextRun({ text, italics: true }));
    else if (tag === "u") runs.push(new TextRun({ text, underline: { type: UnderlineType.SINGLE } }));
    else if (tag === "s") runs.push(new TextRun({ text, strike: true }));
    else if (tag === "a") {
      const hrefMatch = match[0].match(/href="([^"]*)"/);
      const href = hrefMatch ? hrefMatch[1] : "";
      if (href) runs.push(new ExternalHyperlink({ link: href, children: [new TextRun({ text, style: "Hyperlink" })] }));
      else runs.push(new TextRun({ text, underline: { type: UnderlineType.SINGLE } }));
    } else runs.push(new TextRun({ text }));
  }
  if (runs.length === 0) {
    const plain = decodeHtmlEntities(html.replace(/<[^>]+>/g, "").trim());
    if (plain) runs.push(new TextRun({ text: plain }));
  }
  return runs;
}

// ── HTML -> ODT via odf-kit ──────────────────────────────

async function htmlToOdt(html: string): Promise<ArrayBuffer> {
  const odfKit = await import("npm:odf-kit@0.13.1");
  const bytes = await odfKit.htmlToOdt(html);
  return bytes.buffer;
}

// ── HTML -> RTF (DOC) ───────────────────────────────────

async function htmlToRtf(html: string): Promise<ArrayBuffer> {
  const paragraphs = htmlToStyledParagraphs(html);
  return buildStyledRtf(paragraphs);
}

// ── HTML block parser ───────────────────────────────────

interface HtmlBlock { type: "heading" | "paragraph" | "listitem" | "table"; level?: number; content: string; }

function parseHtmlBlocks(html: string): HtmlBlock[] {
  const blocks: HtmlBlock[] = [];
  const headingRegex = /<h([1-6])[^>]*>([\s\S]*?)<\/h[1-6]>/gi;
  const processed = html.replace(headingRegex, (_, level, content) => { blocks.push({ type: "heading", level: parseInt(level), content: content.trim() }); return ""; });
  const liRegex = /<li[^>]*>([\s\S]*?)<\/li>/gi;
  const afterLists = processed.replace(liRegex, (_, content) => { blocks.push({ type: "listitem", content: content.trim() }); return ""; });
  const pRegex = /<p[^>]*>([\s\S]*?)<\/p>/gi;
  const afterParas = afterLists.replace(pRegex, (_, content) => { blocks.push({ type: "paragraph", content: content.trim() }); return ""; });
  const remaining = afterParas.replace(/<[^>]+>/g, "").trim();
  if (remaining) blocks.push({ type: "paragraph", content: remaining });
  if (blocks.length === 0) { const plain = html.replace(/<[^>]+>/g, "").trim(); if (plain) blocks.push({ type: "paragraph", content: plain }); }
  return blocks;
}

// ── StyledParagraph types ────────────────────────────────

interface StyledRun { text: string; bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean; fontSize?: number; }
interface StyledParagraph { runs: StyledRun[]; heading?: boolean; headingLevel?: number; }

function htmlToStyledParagraphs(html: string): StyledParagraph[] {
  const paragraphs: StyledParagraph[] = [];
  const blocks = parseHtmlBlocks(html);
  for (const block of blocks) {
    const runs = htmlInlineToStyledRuns(block.content);
    if (block.type === "heading" && block.level) {
      for (const run of runs) { run.bold = true; if (!run.fontSize) run.fontSize = Math.max(28 - block.level * 2, 12); }
      paragraphs.push({ runs, heading: true, headingLevel: block.level });
    } else paragraphs.push({ runs });
  }
  return paragraphs;
}

function htmlInlineToStyledRuns(html: string): StyledRun[] {
  const runs: StyledRun[] = [];
  const tokenRegex = /<(strong|b|em|i|u|s|a|span)[^>]*>([\s\S]*?)<\/\1>|([^<]+)/gi;
  let match;
  while ((match = tokenRegex.exec(html)) !== null) {
    if (match[3] !== undefined) { const text = decodeHtmlEntities(match[3]); if (text) runs.push({ text }); continue; }
    const tag = match[1].toLowerCase();
    const text = decodeHtmlEntities(match[2].replace(/<[^>]+>/g, ""));
    if (!text) continue;
    const run: StyledRun = { text };
    if (tag === "strong" || tag === "b") run.bold = true;
    else if (tag === "em" || tag === "i") run.italic = true;
    else if (tag === "u") run.underline = true;
    else if (tag === "s") run.strike = true;
    runs.push(run);
  }
  if (runs.length === 0) { const plain = decodeHtmlEntities(html.replace(/<[^>]+>/g, "").trim()); if (plain) runs.push({ text: plain }); }
  return runs;
}

// ── RTF parser ──────────────────────────────────────────

function parseRtfToStyledParagraphs(data: ArrayBuffer): StyledParagraph[] {
  const text = new TextDecoder("utf-8", { fatal: false }).decode(data);
  const paragraphs: StyledParagraph[] = [];
  const parts = text.split(/\\par\b/);
  for (const part of parts) {
    const runs: StyledRun[] = [];
    let bold = false, italic = false, underline = false, strike = false, fontSize: number | undefined, currentText = "";
    let i = 0;
    while (i < part.length) {
      if (part[i] === "\\") {
        const cmdMatch = part.slice(i).match(/^\\([a-zA-Z]+)(-?\d*)\s?/);
        if (cmdMatch) {
          const cmd = cmdMatch[1]; const val = cmdMatch[2] ? parseInt(cmdMatch[2]) : undefined;
          if (currentText) { runs.push({ text: currentText, bold, italic, underline, strike, fontSize }); currentText = ""; }
          if (cmd === "b" && val !== 0) bold = true; else if (cmd === "b" && val === 0) bold = false;
          else if (cmd === "i" && val !== 0) italic = true; else if (cmd === "i" && val === 0) italic = false;
          else if (cmd === "ul" && val !== 0) underline = true; else if (cmd === "ul" && val === 0) underline = false;
          else if (cmd === "ulnone") underline = false;
          else if (cmd === "strike" && val !== 0) strike = true; else if (cmd === "strike" && val === 0) strike = false;
          else if (cmd === "fs" && val) fontSize = val / 2;
          else if (cmd === "plain") { bold = false; italic = false; underline = false; strike = false; fontSize = undefined; }
          i += cmdMatch[0].length; continue;
        }
        if (part[i + 1] === "\\") { currentText += "\\"; i += 2; continue; }
        if (part[i + 1] === "{") { currentText += "{"; i += 2; continue; }
        if (part[i + 1] === "}") { currentText += "}"; i += 2; continue; }
        const hexMatch = part.slice(i).match(/^\\'([0-9a-fA-F]{2})/);
        if (hexMatch) { currentText += String.fromCharCode(parseInt(hexMatch[1], 16)); i += hexMatch[0].length; continue; }
        i++; continue;
      }
      if (part[i] === "{" || part[i] === "}") { i++; continue; }
      currentText += part[i]; i++;
    }
    if (currentText.trim()) runs.push({ text: currentText.trim(), bold, italic, underline, strike, fontSize });
    if (runs.length > 0) paragraphs.push({ runs });
  }
  return paragraphs;
}

// ── RTF builder ──────────────────────────────────────────

function buildStyledRtf(paragraphs: StyledParagraph[]): ArrayBuffer {
  const fontTable = `{\\fonttbl{\\f0 Calibri;}{\\f1 Arial;}{\\f2 Times New Roman;}}`;
  const colorTable = `{\\colortbl;\\red0\\green0\\blue0;}`;
  const rtfParas = paragraphs.map((p) => {
    const rtfRuns = p.runs.map((run) => {
      const prefixes: string[] = [];
      if (run.bold) prefixes.push("\\b"); if (run.italic) prefixes.push("\\i");
      if (run.underline) prefixes.push("\\ul"); if (run.strike) prefixes.push("\\strike");
      if (run.fontSize) prefixes.push(`\\fs${Math.round(run.fontSize * 2)}`);
      const text = run.text.replace(/\\/g, "\\\\").replace(/\{/g, "\\{").replace(/\}/g, "\\}");
      const resets: string[] = [];
      if (run.bold) resets.push("\\b0"); if (run.italic) resets.push("\\i0");
      if (run.underline) resets.push("\\ul0"); if (run.strike) resets.push("\\strike0");
      return `{${prefixes.join("")} ${text}${resets.join("")}}`;
    }).join("");
    return `${rtfRuns}\\par`;
  }).join("\n");
  const rtf = `{\\rtf1\\ansi\\deff0\n${fontTable}\n${colorTable}\n\\f0\\fs22\n${rtfParas}\n}`;
  return new TextEncoder().encode(rtf).buffer;
}

// ════════════════════════════════════════════════════════
// SPREADSHEET CONVERSIONS
// ════════════════════════════════════════════════════════

async function convertSpreadsheet(data: ArrayBuffer, sourceExt: string, targetExt: string): Promise<ArrayBuffer> {
  const hasCloudConvert = !!Deno.env.get("CLOUDCONVERT_API_KEY");
  if (hasCloudConvert) {
    try {
      return await cloudConvert(data, sourceExt, targetExt, `spreadsheet${sourceExt}`);
    } catch (ccErr) {
      console.error("CloudConvert failed for spreadsheet, falling back:", ccErr);
    }
  }

  // Local fallback
  if (sourceExt === ".xlsx" && targetExt === ".ods") {
    return await xlsxToOdsViaOdfKit(data);
  }
  const XLSX = await import("npm:xlsx@0.18.5");
  const workbook = XLSX.read(new Uint8Array(data), { type: "array", cellStyles: true, cellDates: true, cellNF: true, sheetStubs: true });
  if (targetExt === ".csv") {
    const sheets: string[] = [];
    for (const name of workbook.SheetNames) sheets.push(`--- ${name} ---\n${XLSX.utils.sheet_to_csv(workbook.Sheets[name])}`);
    return new TextEncoder().encode(sheets.join("\n\n")).buffer;
  }
  const bookTypeMap: Record<string, string> = { ".xlsx": "xlsx", ".xls": "xls", ".ods": "ods", ".csv": "csv" };
  const out = XLSX.write(workbook, { bookType: (bookTypeMap[targetExt] ?? "xlsx") as any, type: "array", cellStyles: true, cellDates: true });
  return (out as Uint8Array).buffer;
}

async function xlsxToOdsViaOdfKit(data: ArrayBuffer): Promise<ArrayBuffer> {
  const odfKit = await import("npm:odf-kit@0.13.1");
  const result = await odfKit.xlsxToOds(new Uint8Array(data));
  return result.buffer;
}

// ════════════════════════════════════════════════════════
// IMAGE CONVERSIONS
// ════════════════════════════════════════════════════════

let _resvgInitialized = false;

async function convertImage(data: ArrayBuffer, sourceExt: string, targetExt: string): Promise<ArrayBuffer> {
  if (sourceExt === targetExt) return data;
  if (sourceExt === ".svg" && targetExt === ".svg") return data;
  if (sourceExt === ".svg") return await svgToRaster(data, targetExt);
  if (targetExt === ".svg") return rasterToSvg(data, sourceExt);

  const Jimp = (await import("npm:jimp@0.22.12")).default;
  const image = await Jimp.read(data);
  const outMime = targetExt === ".png" ? "image/png" : "image/jpeg";
  const outBuffer = await image.getBufferAsync(outMime);
  return outBuffer.buffer;
}

function rasterToSvg(data: ArrayBuffer, sourceExt: string): ArrayBuffer {
  const mimeIn = sourceExt === ".png" ? "image/png" : "image/jpeg";
  const b64 = btoa(String.fromCharCode(...new Uint8Array(data)));
  const svg = `<?xml version="1.0" encoding="UTF-8"?>\n<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">\n  <image width="100%" height="100%" xlink:href="data:${mimeIn};base64,${b64}"/>\n</svg>`;
  return new TextEncoder().encode(svg).buffer;
}

async function svgToRaster(data: ArrayBuffer, targetExt: string): Promise<ArrayBuffer> {
  if (!_resvgInitialized) {
    const { initWasm } = await import("npm:@resvg/resvg-wasm@2.6.2");
    const wasmResponse = await fetch("https://cdn.jsdelivr.net/npm/@resvg/resvg-wasm@2.6.2/index_bg.wasm");
    await initWasm(wasmResponse);
    _resvgInitialized = true;
  }
  const { Resvg } = await import("npm:@resvg/resvg-wasm@2.6.2");
  const svgText = new TextDecoder().decode(data);
  const resvg = new Resvg(svgText, { fitTo: { mode: "width", value: 800 } });
  const pngData = resvg.render();
  const pngBuffer = pngData.asPng();
  if (targetExt === ".png") return pngBuffer.buffer;
  const Jimp = (await import("npm:jimp@0.22.12")).default;
  const image = await Jimp.read(pngBuffer.buffer);
  const outBuffer = await image.getBufferAsync("image/jpeg");
  return outBuffer.buffer;
}

// ════════════════════════════════════════════════════════
// LOCAL PRESENTATION CONVERSIONS (fallback)
// ════════════════════════════════════════════════════════

async function convertPresentationLocal(data: ArrayBuffer, sourceExt: string, targetExt: string): Promise<ArrayBuffer> {
  if (sourceExt === ".pptx" && targetExt === ".odp") return await pptxToOdp(data);
  if (sourceExt === ".odp" && targetExt === ".pptx") return await odpToPptx(data);
  return data;
}

async function pptxToOdp(data: ArrayBuffer): Promise<ArrayBuffer> {
  const JSZip = (await import("npm:jszip@3.10.1")).default;
  const srcZip = await JSZip.loadAsync(data);
  const slideFiles: string[] = [];
  srcZip.forEach((path) => { if (/^ppt\/slides\/slide\d+\.xml$/.test(path)) slideFiles.push(path); });
  slideFiles.sort((a, b) => parseInt(a.match(/slide(\d+)/)?.[1] || "0") - parseInt(b.match(/slide(\d+)/)?.[1] || "0"));
  let slideWidth = 9144000, slideHeight = 6858000;
  const presXml = await srcZip.file("ppt/presentation.xml")?.async("string");
  if (presXml) { const m = presXml.match(/<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"/); if (m) { slideWidth = parseInt(m[1]); slideHeight = parseInt(m[2]); } }
  const widthCm = (slideWidth / 360000).toFixed(2);
  const heightCm = (slideHeight / 360000).toFixed(2);
  const autoStyles: string[] = [];
  let styleIdx = 1;
  const odpPages: string[] = [];
  for (let si = 0; si < slideFiles.length; si++) {
    const slideXml = await srcZip.file(slideFiles[si])?.async("string");
    if (!slideXml) continue;
    const shapes = extractPptxShapes(slideXml);
    const frameElements: string[] = [];
    for (const shape of shapes) {
      const dpStyleName = `dp${styleIdx++}`;
      const xCm = (shape.x / 360000).toFixed(2); const yCm = (shape.y / 360000).toFixed(2);
      const wCm = (shape.w / 360000).toFixed(2); const hCm = (shape.h / 360000).toFixed(2);
      if (shape.texts.length > 0) {
        const textParagraphs = shape.texts.map((t) => {
          const textStyleProps: string[] = [];
          if (t.bold) textStyleProps.push("fo:font-weight=\"bold\"");
          if (t.italic) textStyleProps.push("fo:font-style=\"italic\"");
          if (t.fontSize) textStyleProps.push(`fo:font-size="${t.fontSize}pt"`);
          if (textStyleProps.length > 0) {
            const textStyleName = `T${styleIdx++}`;
            autoStyles.push(`<style:style style:name="${textStyleName}" style:family="text"><style:text-properties ${textStyleProps.join(" ")}/></style:style>`);
            return `<text:p><text:span text:style-name="${textStyleName}">${escapeXml(t.text)}</text:span></text:p>`;
          }
          return `<text:p>${escapeXml(t.text)}</text:p>`;
        }).join("\n");
        frameElements.push(`<draw:frame draw:style-name="${dpStyleName}" svg:x="${xCm}cm" svg:y="${yCm}cm" svg:width="${wCm}cm" svg:height="${hCm}cm"><draw:text-box>${textParagraphs}</draw:text-box></draw:frame>`);
      }
      autoStyles.push(`<style:style style:name="${dpStyleName}" style:family="graphic"><style:graphic-properties draw:fill="none"/></style:style>`);
    }
    const pageStyleName = `page${si + 1}`;
    autoStyles.push(`<style:style style:name="${pageStyleName}" style:family="drawing-page"><style:drawing-page-properties/></style:style>`);
    odpPages.push(`<draw:page draw:name="slide${si + 1}" draw:id="page${si + 1}" draw:style-name="${pageStyleName}">${frameElements.join("")}</draw:page>`);
  }
  const zip = new JSZip();
  zip.file("mimetype", "application/vnd.oasis.opendocument.presentation", { compression: "STORE" });
  zip.file("content.xml", `<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" office:version="1.2"><office:automatic-styles><style:page-layout style:name="PL1"><style:page-layout-properties fo:page-width="${widthCm}cm" fo:page-height="${heightCm}cm"/></style:page-layout>${autoStyles.join("\n    ")}</office:automatic-styles><office:body><office:presentation>${odpPages.join("\n")}</office:presentation></office:body></office:document-content>`);
  zip.file("META-INF/manifest.xml", buildOdtManifest("application/vnd.oasis.opendocument.presentation", "content.xml"));
  zip.file("styles.xml", `<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" office:version="1.2"><office:automatic-styles><style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="${widthCm}cm" fo:page-height="${heightCm}cm" style:print-orientation="landscape"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="PM1"><presentation:notes style:page-layout-name="PM1"/></style:master-page></office:master-styles></office:document-styles>`);
  return ((await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" })) as Uint8Array).buffer;
}

interface PptxShape { x: number; y: number; w: number; h: number; texts: { text: string; bold?: boolean; italic?: boolean; fontSize?: number }[]; }

function extractPptxShapes(xml: string): PptxShape[] {
  const shapes: PptxShape[] = [];
  const spRegex = /<p:sp[\s>][\s\S]*?<\/p:sp>/g;
  let spMatch;
  while ((spMatch = spRegex.exec(xml)) !== null) {
    const spXml = spMatch[0];
    let x = 457200, y = 274638, w = 8229600, h = 457200;
    const offMatch = spXml.match(/<a:off x="(\d+)" y="(\d+)"/);
    const extMatch = spXml.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
    if (offMatch) { x = parseInt(offMatch[1]); y = parseInt(offMatch[2]); }
    if (extMatch) { w = parseInt(extMatch[1]); h = parseInt(extMatch[2]); }
    const texts: PptxShape["texts"] = [];
    const rRegex = /<a:r[\s>][\s\S]*?<\/a:r>/g;
    let rMatch;
    while ((rMatch = rRegex.exec(spXml)) !== null) {
      const rXml = rMatch[0];
      const tMatch = rXml.match(/<a:t[^>]*>([\s\S]*?)<\/a:t>/);
      if (!tMatch) continue;
      const text = unescapeXml(tMatch[1]);
      if (!text.trim()) continue;
      const bold = /<a:rPr[^>]*b="1"/.test(rXml) || /<a:rPr[^>]*b="true"/.test(rXml);
      const italic = /<a:rPr[^>]*i="1"/.test(rXml) || /<a:rPr[^>]*i="true"/.test(rXml);
      let fontSize: number | undefined;
      const szMatch = rXml.match(/<a:rPr[^>]*sz="(\d+)"/);
      if (szMatch) fontSize = parseInt(szMatch[1]) / 100;
      texts.push({ text: text.trim(), bold, italic, fontSize });
    }
    if (texts.length > 0) shapes.push({ x, y, w, h, texts });
  }
  return shapes;
}

async function odpToPptx(data: ArrayBuffer): Promise<ArrayBuffer> {
  const JSZip = (await import("npm:jszip@3.10.1")).default;
  const srcZip = await JSZip.loadAsync(data);
  const contentXml = await srcZip.file("content.xml")?.async("string");
  if (!contentXml) throw new Error("No se pudo leer content.xml del ODP");
  let slideWidthEmu = 9144000, slideHeightEmu = 6858000;
  const wMatch = contentXml.match(/fo:page-width="([\d.]+)cm"/);
  const hMatch = contentXml.match(/fo:page-height="([\d.]+)cm"/);
  if (wMatch) slideWidthEmu = Math.round(parseFloat(wMatch[1]) * 360000);
  if (hMatch) slideHeightEmu = Math.round(parseFloat(hMatch[1]) * 360000);
  const pages = extractOdpPagesStructured(contentXml);
  const zip = new JSZip();
  const slideRels: string[] = [], slideContentTypes: string[] = [], slideEntries: string[] = [];
  for (let i = 0; i < pages.length; i++) {
    const slideNum = i + 1;
    const rId = `rId${slideNum + 1}`;
    slideRels.push(`<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${slideNum}.xml"/>`);
    slideContentTypes.push(`<Override PartName="/ppt/slides/slide${slideNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`);
    slideEntries.push(`<p:sldId id="${256 + i}" r:id="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`);
    const shapeElements = pages[i].shapes.map((shape, j) => {
      const xEmu = Math.round(shape.x * 360000); const yEmu = Math.round(shape.y * 360000);
      const wEmu = Math.round(shape.w * 360000); const hEmu = Math.round(shape.h * 360000);
      const textRuns = shape.texts.map((t) => {
        const attrs: string[] = ['lang="es"', 'dirty="0"'];
        if (t.bold) attrs.push('b="1"'); if (t.italic) attrs.push('i="1"');
        if (t.fontSize) attrs.push(`sz="${Math.round(t.fontSize * 100)}"`);
        return `<a:r><a:rPr ${attrs.join(" ")}/><a:t>${escapeXml(t.text)}</a:t></a:r>`;
      }).join("");
      return `<p:sp><p:nvSpPr><p:cNvPr id="${j + 2}" name="Text ${j + 1}"/><p:cNvSpPr txBox="0"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:off x="${xEmu}" y="${yEmu}"/><a:ext cx="${wEmu}" cy="${hEmu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle/><a:p>${textRuns}</a:p></p:txBody></p:sp>`;
    }).join("");
    zip.file(`ppt/slides/slide${slideNum}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${shapeElements}</p:spTree></p:cSld></p:sld>`);
    zip.file(`ppt/slides/_rels/slide${slideNum}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);
  }
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${slideContentTypes.join("")}</Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file("ppt/presentation.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" saveSubsetFonts="1"><p:sldIdLst>${slideEntries.join("")}</p:sldIdLst><p:sldSz cx="${slideWidthEmu}" cy="${slideHeightEmu}"/></p:presentation>`);
  zip.file("ppt/_rels/presentation.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${slideRels.join("")}</Relationships>`);
  return ((await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" })) as Uint8Array).buffer;
}

interface OdpShape { x: number; y: number; w: number; h: number; texts: { text: string; bold?: boolean; italic?: boolean; fontSize?: number }[]; }
interface OdpPage { shapes: OdpShape[]; }

function extractOdpPagesStructured(xml: string): OdpPage[] {
  const pages: OdpPage[] = [];
  const pageRegex = /<draw:page[^>]*>([\s\S]*?)<\/draw:page>/g;
  let pMatch;
  while ((pMatch = pageRegex.exec(xml)) !== null) {
    const shapes: OdpShape[] = [];
    const frameRegex = /<draw:frame[^>]*>([\s\S]*?)<\/draw:frame>/g;
    let fMatch;
    while ((fMatch = frameRegex.exec(pMatch[1])) !== null) {
      const frameXml = fMatch[0]; const frameContent = fMatch[1];
      let x = 0, y = 0, w = 10, h = 2;
      const xm = frameXml.match(/svg:x="([\d.]+)cm"/); const ym = frameXml.match(/svg:y="([\d.]+)cm"/);
      const wm = frameXml.match(/svg:width="([\d.]+)cm"/); const hm = frameXml.match(/svg:height="([\d.]+)cm"/);
      if (xm) x = parseFloat(xm[1]); if (ym) y = parseFloat(ym[1]);
      if (wm) w = parseFloat(wm[1]); if (hm) h = parseFloat(hm[1]);
      const texts: OdpShape["texts"] = [];
      const pRegex = /<text:p[^>]*>([\s\S]*?)<\/text:p>/g;
      let tMatch;
      while ((tMatch = pRegex.exec(frameContent)) !== null) {
        const pContent = tMatch[1];
        const spanRegex = /<text:span[^>]*>([\s\S]*?)<\/text:span>/g;
        let sMatch; let hasSpans = false;
        while ((sMatch = spanRegex.exec(pContent)) !== null) {
          hasSpans = true;
          const spanText = sMatch[1].replace(/<[^>]+>/g, "").trim();
          if (!spanText) continue;
          const styleNameMatch = sMatch[0].match(/text:style-name="([^"]*)"/);
          let bold = false, italic = false, fontSize: number | undefined;
          if (styleNameMatch) {
            const styleRegex = new RegExp(`<style:style[^>]*style:name="${escapeRegex(styleNameMatch[1])}"[^>]*>[\\s\\S]*?<style:text-properties([^/]*)/>`, "i");
            const styleMatch = xml.match(styleRegex);
            if (styleMatch) {
              bold = /fo:font-weight="bold"/i.test(styleMatch[1]);
              italic = /fo:font-style="italic"/i.test(styleMatch[1]);
              const fsMatch = styleMatch[1].match(/fo:font-size="([\d.]+)pt"/);
              if (fsMatch) fontSize = parseFloat(fsMatch[1]);
            }
          }
          texts.push({ text: spanText, bold, italic, fontSize });
        }
        if (!hasSpans) { const plainText = pContent.replace(/<[^>]+>/g, "").trim(); if (plainText) texts.push({ text: plainText }); }
      }
      if (texts.length > 0) shapes.push({ x, y, w, h, texts });
    }
    pages.push({ shapes });
  }
  return pages;
}

// ════════════════════════════════════════════════════════
// SHARED HELPERS
// ════════════════════════════════════════════════════════

function buildOdtManifest(mediaType: string, extraFile = "content.xml"): string {
  return `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2"><manifest:file-entry manifest:media-type="${mediaType}" manifest:full-path="/"/><manifest:file-entry manifest:media-type="text/xml" manifest:full-path="${extraFile}"/></manifest:manifest>`;
}

function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

function unescapeXml(s: string): string {
  return s.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'");
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function decodeHtmlEntities(s: string): string {
  return s.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&apos;/g, "'").replace(/&nbsp;/g, " ");
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
