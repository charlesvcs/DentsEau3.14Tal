const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

const XLSX_URL =
  process.env.XLSX_URL ||
  "https://kdrive.infomaniak.com/2/app/579686/share/b02b8fd8-5408-418e-bf6e-e82c7917b648/files/39/download";

const APP_PASSWORD = process.env.APP_PASSWORD || "dentaire2026";

const NEED_HEADER = /besoin/i;
const CHECK_FONT_COLOR = false;

const norm = (s) =>
  String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();

const THEME = ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"];
const INDEXED = { 2: "FF0000", 3: "00FF00", 10: "FF0000", 11: "00FF00", 17: "008000", 51: "99CC00", 42: "00CCFF", 53: "993300", 57: "003300", 58: "003300" };

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
  if (color.theme !== undefined && THEME[color.theme]) return applyTint(hexToRgb(THEME[color.theme]), color.tint);
  if (color.indexed !== undefined && INDEXED[color.indexed]) return hexToRgb(INDEXED[color.indexed]);
  return null;
}

function isGreenOrRed([r, g, b]) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  const l = (max + min) / 2;
  const d = max - min;
  if (d === 0) return false;
  const s = d / (1 - Math.abs(2 * l - 1));
  if (s < 0.2 || l < 0.12 || l > 0.97) return false;
  let h;
  if (max === r) h = ((g - b) / d) % 6;
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  h = (h * 60 + 360) % 360;
  return h <= 20 || h >= 340 || (h >= 70 && h <= 170);
}

function cellIsGreenOrRed(cell) {
  const f = cell.fill;
  if (f && f.type === "pattern" && f.pattern && f.pattern !== "none") {
    const rgb = colorToRgb(f.fgColor);
    if (rgb && isGreenOrRed(rgb)) return true;
  }
  if (CHECK_FONT_COLOR && cell.font && cell.font.color) {
    const rgb = colorToRgb(cell.font.color);
    if (rgb && isGreenOrRed(rgb)) return true;
  }
  return false;
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

    let headers = null;
    const rows = [];
    let excluded = 0;

    wb.eachSheet((ws) => {
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

      const maxCol = Math.max(...cols.map((x) => x.n));

      for (let i = headerRowNum + 1; i <= ws.rowCount; i++) {
        const row = ws.getRow(i);
        const values = cols.map((c) => removeDashes(txt(row.getCell(c.n).value)));
        if (!values.some(Boolean)) continue;

        let bad = false;
        for (let c = 1; c <= maxCol; c++) {
          if (cellIsGreenOrRed(row.getCell(c))) {
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
