const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const zlib = require("zlib");

const XLSX_URL =
  process.env.XLSX_URL ||
  "https://kdrive.infomaniak.com/2/app/579686/share/b02b8fd8-5408-418e-bf6e-e82c7917b648/files/39/download";

const APP_PASSWORD = process.env.APP_PASSWORD || "dentaire2026";

const NEED_HEADER = /besoin/i;
const CHECK_FONT_COLOR = false;

const norm = (s) =>
  String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();

const THEME = [
  "FFFFFF",
  "000000",
  "E7E6E6",
  "44546A",
  "4472C4",
  "ED7D31",
  "A5A5A5",
  "FFC000",
  "5B9BD5",
  "70AD47"
];

const INDEXED = {
  2: "FF0000",
  3: "00FF00",
  5: "FFFF00",
  6: "FFFF00",
  10: "FF0000",
  11: "00FF00",
  13: "FFFF00",
  14: "FF00FF",
  16: "800000",
  17: "008000",
  19: "808000",
  26: "FFFF99",
  27: "FFFFCC",
  29: "FF8080",
  34: "FFFF99",
  35: "CCFFCC",
  36: "FFFF99",
  42: "00CCFF",
  43: "FFFF99",
  45: "FF99CC",
  46: "FF9900",
  47: "FFCC99",
  50: "339966",
  51: "99CC00",
  52: "FF9900",
  53: "993300",
  57: "003300",
  58: "003300"
};

function hexToRgb(hex) {
  const h = hex.slice(-6);
  return [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16));
}

function applyTint(rgb, tint) {
  if (!tint) return rgb;
  return rgb.map((c) => Math.round(tint < 0 ? c * (1 + tint) : c + (255 - c) * tint));
}

function colorToRgb(color) {
  if (!color) return null;
  if (color.argb) return hexToRgb(color.argb);
  if (color.rgb) return hexToRgb(color.rgb);
  if (color.theme !== undefined && THEME[color.theme]) {
    return applyTint(hexToRgb(THEME[color.theme]), Number(color.tint || 0));
  }
  if (color.indexed !== undefined && INDEXED[color.indexed]) {
    return hexToRgb(INDEXED[color.indexed]);
  }
  return null;
}

function isGreenOrangeOrRed([r, g, b]) {
  r /= 255;
  g /= 255;
  b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  const d = max - min;
  if (d === 0) return false;
  const s = d / (1 - Math.abs(2 * l - 1));
  if (s < 0.18 || l < 0.10 || l > 0.96) return false;
  let h;
  if (max === r) h = ((g - b) / d) % 6;
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  h = (h * 60 + 360) % 360;
  return h <= 175 || h >= 335;
}

function fillIsGreenOrangeOrRed(f) {
  if (!f || f.type !== "pattern" || !f.pattern || f.pattern === "none" || f.pattern === "gray125") {
    return false;
  }
  const fg = f.fgColor;
  if (fg && fg.theme === 7 && fg.tint && fg.tint > 0.7) {
    return false;
  }
  const rgb = colorToRgb(fg) || colorToRgb(f.bgColor);
  return Boolean(rgb && isGreenOrangeOrRed(rgb));
}

function cellIsGreenOrangeOrRed(cell) {
  if (!cell) return false;
  if (fillIsGreenOrangeOrRed(cell.fill)) return true;
  if (CHECK_FONT_COLOR && cell.font && cell.font.color) {
    const rgb = colorToRgb(cell.font.color);
    if (rgb && isGreenOrangeOrRed(rgb)) return true;
  }
  return false;
}

function extractZipSheetXmls(buf) {
  const sheets = [];
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) return sheets;
  const count = buf.readUInt16LE(eocd + 10);
  let pos = buf.readUInt32LE(eocd + 16);
  for (let i = 0; i < count && pos < buf.length; i++) {
    if (buf.readUInt32LE(pos) !== 0x02014b50) break;
    const method = buf.readUInt16LE(pos + 10);
    const compSize = buf.readUInt32LE(pos + 20);
    const nameLen = buf.readUInt16LE(pos + 28);
    const extraLen = buf.readUInt16LE(pos + 30);
    const commentLen = buf.readUInt16LE(pos + 32);
    const localOffset = buf.readUInt32LE(pos + 42);
    const name = buf.toString("utf8", pos + 46, pos + 46 + nameLen);
    pos += 46 + nameLen + extraLen + commentLen;
    if (/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) {
      const lNameLen = buf.readUInt16LE(localOffset + 26);
      const lExtraLen = buf.readUInt16LE(localOffset + 28);
      const start = localOffset + 30 + lNameLen + lExtraLen;
      const raw = buf.subarray(start, start + compSize);
      const xml = (method === 0 ? raw : zlib.inflateRawSync(raw)).toString("utf8");
      sheets.push({ name, xml });
    }
  }
  return sheets;
}

function parseColorAttrs(tag) {
  const obj = {};
  const re = /(\w+)="([^"]*)"/g;
  let m;
  while ((m = re.exec(tag))) {
    const k = m[1];
    const v = m[2];
    if (k === "theme" || k === "indexed") obj[k] = Number(v);
    else if (k === "tint") obj[k] = Number(v);
    else obj[k] = v;
  }
  return obj;
}

function getConditionalColoredRows(xml, ws) {
  const badRows = new Set();
  const cfRe = /<(?:\w+:)?conditionalFormatting\b[^>]*>([\s\S]*?)<\/(?:\w+:)?conditionalFormatting>/g;
  let cfMatch;
  while ((cfMatch = cfRe.exec(xml))) {
    const block = cfMatch[1];
    const sqMatch = block.match(/<(?:\w+:)?sqref>([^<]+)<\/(?:\w+:)?sqref>/);
    if (!sqMatch) continue;
    const sqref = sqMatch[1].trim();
    const ruleRe = /<(?:\w+:)?cfRule\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?cfRule>/g;
    let rMatch;
    while ((rMatch = ruleRe.exec(block))) {
      const ruleBody = rMatch[2];
      const fgMatch = ruleBody.match(/<(?:\w+:)?fgColor\b[^>]*>/);
      if (!fgMatch) continue;
      const colorObj = parseColorAttrs(fgMatch[0]);
      if (colorObj.theme === 7 && colorObj.tint && colorObj.tint > 0.7) continue;
      const rgb = colorToRgb(colorObj);
      if (!rgb || !isGreenOrangeOrRed(rgb)) continue;

      const fMatch = ruleBody.match(/<(?:\w+:)?f>([\s\S]*?)<\/(?:\w+:)?f>/);
      if (!fMatch) continue;
      const formula = fMatch[1].replace(/&quot;/g, '"');
      const sMatch = formula.match(/SEARCH\("([^"]+)",\s*([A-Z]+\d+)\)/i);
      if (!sMatch) continue;
      const needle = sMatch[1].toLowerCase();
      const refCell = sMatch[2];
      const refVal = txt(ws.getCell(refCell).value).toLowerCase();
      if (!refVal.includes(needle)) continue;

      for (const part of sqref.split(/\s+/)) {
        const m = part.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
        if (!m) continue;
        const r1 = Number(m[2]);
        const r2 = m[4] ? Number(m[4]) : r1;
        for (let r = r1; r <= r2; r++) {
          if (r <= ws.rowCount) {
            const cellVal = txt(ws.getCell(`${m[1]}${r}`).value).toLowerCase();
            if (cellVal.includes(needle)) {
              badRows.add(r);
            }
          }
        }
      }
    }
  }
  return badRows;
}

const pad = (n) => String(n).padStart(2, "0");

function removeDashes(str) {
  return String(str || "")
    .split(/(\S+)/)
    .map((token) => (token.includes("@") ? token : token.replace(/[\u002D\u2010-\u2015\u2212]+/g, " ")))
    .join("")
    .replace(/\s{2,}/g, " ")
    .trim();
}

function txt(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) return `${pad(v.getUTCDate())}/${pad(v.getUTCMonth() + 1)}/${v.getUTCFullYear()}`;
  if (typeof v === "object") {
    if (v.richText) return v.richText.map((t) => t.text).join("").trim();
    if ("result" in v) return txt(v.result);
    if (v.text !== undefined) return txt(v.text);
    if (v.hyperlink) return v.hyperlink;
    return "";
  }
  return String(v).trim();
}

function kindOf(label) {
  const n = norm(label);
  if (NEED_HEADER.test(n)) return "need";
  if (/naissance|ne\(e\) le|\bdob\b/.test(n)) return "birth";
  if (/\btel|phone|mobile|portable|gsm|mail|courriel|contact|whatsapp/.test(n)) return "contact";
  return "other";
}

async function loadExcelBuffer() {
  try {
    const r = await fetch(XLSX_URL, { redirect: "follow" });
    if (!r.ok) throw new Error(`Téléchargement impossible (HTTP ${r.status})`);
    const buf = Buffer.from(await r.arrayBuffer());
    if (buf.slice(0, 2).toString() !== "PK") {
      throw new Error("Le lien ne renvoie pas un fichier .xlsx valide");
    }
    return { buf, modified: r.headers.get("last-modified") || null };
  } catch (err) {
    const localPath = path.resolve(__dirname, "../../File_Appel_Patients.xlsx");
    if (fs.existsSync(localPath)) {
      const buf = fs.readFileSync(localPath);
      const stat = fs.statSync(localPath);
      return { buf, modified: stat.mtime.toISOString() };
    }
    throw err;
  }
}

module.exports = async (req, res) => {
  res.setHeader("Cache-Control", "no-store");

  const providedPassword =
    req.headers["x-app-password"] ||
    (req.query && req.query.pwd) ||
    "";

  if (providedPassword !== APP_PASSWORD) {
    return res.status(401).json({
      error: "Mot de passe invalide",
      authRequired: true,
    });
  }

  try {
    const { buf, modified } = await loadExcelBuffer();

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const xmlSheets = extractZipSheetXmls(buf);

    let headers = null;
    const rows = [];
    let excluded = 0;
    let sheetIdx = 0;

    wb.eachSheet((ws) => {
      const sheetXmlObj = xmlSheets[sheetIdx++];
      let headerRowNum = 0;
      for (let i = 1; i <= Math.min(15, ws.rowCount) && !headerRowNum; i++) {
        ws.getRow(i).eachCell((c) => {
          if (!headerRowNum && NEED_HEADER.test(norm(txt(c.value)))) headerRowNum = i;
        });
      }
      if (!headerRowNum) return;

      const hr = ws.getRow(headerRowNum);
      const cols = [];
      hr.eachCell((c, n) => {
        const rawLabel = txt(c.value);
        const label = removeDashes(rawLabel);
        if (label) cols.push({ n, label, rawLabel });
      });
      if (!headers) headers = cols.map((c) => ({ label: c.label, kind: kindOf(c.rawLabel) }));

      const maxCol = Math.max(...cols.map((x) => x.n), ws.columnCount || 0);
      const cfBadRows = sheetXmlObj ? getConditionalColoredRows(sheetXmlObj.xml, ws) : new Set();

      for (let i = headerRowNum + 1; i <= ws.rowCount; i++) {
        const row = ws.getRow(i);
        const values = cols.map((c) => removeDashes(txt(row.getCell(c.n).value)));
        if (!values.some(Boolean)) continue;

        let bad = cfBadRows.has(i) || fillIsGreenOrangeOrRed(row.fill);
        const lastCol = Math.max(maxCol, row.cellCount || 0);
        for (let c = 1; !bad && c <= lastCol; c++) {
          if (cellIsGreenOrangeOrRed(row.getCell(c))) {
            bad = true;
            break;
          }
        }
        if (bad) {
          excluded++;
          continue;
        }
        rows.push({ line: i, cells: values });
      }
    });

    if (!headers) throw new Error("Colonne Besoin exprimé introuvable dans le fichier");

    res.status(200).json({
      updated: new Date().toISOString(),
      modified,
      headers,
      rows,
      excluded,
    });
  } catch (e) {
    res.status(500).json({ error: removeDashes(e.message) });
  }
};
