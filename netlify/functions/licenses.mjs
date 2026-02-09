// netlify/functions/licenses.mjs
import * as XLSX from "xlsx";
import JSZip from "jszip";

const BOX_SHARED_URL =
  process.env.BOX_SHARED_URL ||
  "https://app.box.com/s/07rbc57mlzd1az6y7sbx6blgmclhsl70";
const BOX_DIRECT_FILE_URL = process.env.BOX_DIRECT_FILE_URL || "";

// --- Box helpers ---
function toStaticZipUrl(sharedUrl) {
  try {
    const u = new URL(sharedUrl);
    const parts = u.pathname.split("/").filter(Boolean);
    const idx = parts.indexOf("s");
    if (u.hostname === "app.box.com" && idx !== -1 && parts[idx + 1]) {
      const id = parts[idx + 1];
      return `https://app.box.com/shared/static/${id}.zip`;
    }
  } catch {}
  return null;
}
function toDownloadUrl(u) {
  try {
    const url = new URL(u);
    if (url.hostname === "app.box.com") url.searchParams.set("download", "1");
    return url.toString();
  } catch {
    return u;
  }
}

// --- cache ---
let cached = { at: 0, rows: null, idx: null };
const TTL_MS = 10 * 60 * 1000;

// --- column mapping ---
const FIXED = {
  A: 0,
  B: 1,
  C: 2,
  D: 3,
  E: 4,
  F: 5,
  G: 6,
  H: 7,
  I: 8,
  J: 9,
  K: 10,
  L: 11,
  M: 12,
  N: 13,
  O: 14,
  P: 15,
  Q: 16,
  R: 17,
};
const norm = (s) => String(s || "").trim().toLowerCase();
const looksLikeHeader = (row) =>
  /license|licen|category|holder|city|reissu|original|date/.test(
    row.map(norm).join(" ")
  );

const isExactOriginal = (h) => /^\s*original\s+license\s+date\s*$/i.test(h);
const isExactReissue = (h) => /^\s*next\s+reissuance\s*$/i.test(h);

function detectHeaderIdx(row) {
  const headers = row.map(norm);
  const idx = { ...FIXED };

  // Exact matches first
  let q = -1,
    r = -1,
    c = -1;
  for (let i = 0; i < headers.length; i++) {
    if (q === -1 && isExactOriginal(headers[i])) q = i;
    if (r === -1 && isExactReissue(headers[i])) r = i;
    if (c === -1 && /licen.*(#|number|no)/i.test(headers[i])) c = i;
  }
  if (q !== -1) idx.Q = q;
  if (r !== -1) idx.R = r;
  if (c !== -1) idx.C = c;

  // Flexible fallbacks
  if (idx.Q === FIXED.Q && !isExactOriginal(headers[FIXED.Q] || "")) {
    for (let i = 0; i < headers.length; i++) {
      if (
        /(date\s*of\s*original\s*licen|original\s*licen.*date|date.*original)/i.test(
          headers[i]
        )
      ) {
        idx.Q = i;
        break;
      }
    }
  }
  if (idx.R === FIXED.R && !isExactReissue(headers[FIXED.R] || "")) {
    for (let i = 0; i < headers.length; i++) {
      if (
        /(next\s*reissu|re-?\s*issuance|reissuance\s*date|reissue)/i.test(
          headers[i]
        )
      ) {
        idx.R = i;
        break;
      }
    }
  }
  return idx;
}

// --- sheet helpers ---
// IMPORTANT: raw:false so we get the formatted text (cell.w) from Excel
function arrFromSheet(ws) {
  return XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: "",
    raw: false, // <— gives us the human-formatted version from Excel
  });
}

function toRecords(rows, idx) {
  const safe = (i, r) => (i >= 0 && i < r.length ? r[i] ?? "" : "");
  return rows
    .map((r) => {
      const rec = {
        license_number: String(safe(idx.C, r)).trim(),
        license_type: String(safe(idx.B, r)).trim(),
        category: String(safe(idx.A, r)).trim(),
        holder: String(safe(idx.D, r)).trim(),
        dba: String(safe(idx.E, r)).trim(),
        city: String(safe(idx.K, r)).trim(),

        // ⭐️ SHOW EXACTLY WHAT THE SHEET SHOWS:
        original_date: String(safe(idx.Q, r)).trim(),
        next_reissue: String(safe(idx.R, r)).trim(),

        // details
        qualified_rep: String(safe(idx.G, r)).trim(),
        mn_manager: String(safe(idx.H, r)).trim(),
        mn_phone: String(safe(idx.N, r)).trim(),
        corp_phone: String(safe(idx.O, r)).trim(),
        email: String(safe(idx.P, r)).trim(),
      };

      const addr1 = String(safe(idx.I, r)).trim();
      const addr2 = String(safe(idx.J, r)).trim();
      const mn_city = String(safe(idx.K, r)).trim();
      const mn_state = String(safe(idx.L, r)).trim();
      const mn_zip = String(safe(idx.M, r)).trim();
      rec.mn_address = [
        addr1,
        addr2,
        [mn_city, mn_state].filter(Boolean).join(", "),
        mn_zip,
      ]
        .filter(Boolean)
        .join(" • ");

      return rec;
    })
    .filter((rec) => rec.license_number || rec.holder || rec.dba);
}

function chooseSheetArray(wb) {
  let best = wb.SheetNames[0],
    bestLen = -1;
  for (const n of wb.SheetNames) {
    const ws = wb.Sheets[n];
    const arr = arrFromSheet(ws);
    if (arr.length > bestLen) {
      bestLen = arr.length;
      best = n;
    }
  }
  return best;
}

async function parseWorkbookFromArrayBuffer(buf) {
  // ⭐️ read with formatting preserved
  const wb = XLSX.read(new Uint8Array(buf), {
    type: "array",
    cellDates: false,
    cellNF: true,
    cellText: true,
  });
  const ws = wb.Sheets[chooseSheetArray(wb)];
  let rows = arrFromSheet(ws);
  let idx = FIXED;
  if (rows.length && looksLikeHeader(rows[0])) {
    idx = detectHeaderIdx(rows[0]);
    rows = rows.slice(1);
  }
  return { rows: toRecords(rows, idx), idx };
}

async function parseCsvText(csvText) {
  const wb = XLSX.read(csvText, { type: "string" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  let rows = arrFromSheet(ws);
  let idx = FIXED;
  if (rows.length && looksLikeHeader(rows[0])) {
    idx = detectHeaderIdx(rows[0]);
    rows = rows.slice(1);
  }
  return { rows: toRecords(rows, idx), idx };
}

// --- fetch & parse ---
async function fetchBuffer(u) {
  const res = await fetch(u, {
    redirect: "follow",
    headers: {
      Accept: "*/*",
      "User-Agent": "Mozilla/5.0 (NetlifyFunction)",
    },
  });
  const ct = (res.headers.get("content-type") || "").toLowerCase();
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(
      `Failed to fetch: HTTP ${res.status}. content-type=${ct}. ${body.slice(
        0,
        200
      )}`
    );
  }
  const buf = await res.arrayBuffer();
  return { ct, buf };
}

async function fetchAndParse() {
  const now = Date.now();
  if (cached.rows && now - cached.at < TTL_MS) return cached;

  const attempts = [];
  if (BOX_DIRECT_FILE_URL) attempts.push(BOX_DIRECT_FILE_URL);
  attempts.push(toDownloadUrl(BOX_SHARED_URL));
  const staticZip = toStaticZipUrl(BOX_SHARED_URL);
  if (staticZip) attempts.push(staticZip);

  let lastErr;
  for (const url of attempts.filter(Boolean)) {
    try {
      const { ct, buf } = await fetchBuffer(url);
      let out;
      if (ct.includes("application/zip") || ct.includes("zip")) {
        const zip = await JSZip.loadAsync(buf);
        const candidates = Object.keys(zip.files).filter((n) =>
          /\.(xlsx|csv)$/i.test(n)
        );
        if (!candidates.length)
          throw new Error("Downloaded ZIP contains no .xlsx or .csv files.");
        const preferred = candidates
          .sort((a, b) => {
            const ax = a.toLowerCase().endsWith(".xlsx") ? 0 : 1;
            const bx = b.toLowerCase().endsWith(".xlsx") ? 0 : 1;
            return ax - bx || b.length - a.length;
          })[0];
        const entry = zip.files[preferred];
        if (preferred.toLowerCase().endsWith(".csv")) {
          const text = await entry.async("string");
          out = await parseCsvText(text);
        } else {
          const ab = await entry.async("arraybuffer");
          out = await parseWorkbookFromArrayBuffer(ab);
        }
      } else if (ct.includes("text/csv") || ct.includes("csv")) {
        const text = new TextDecoder("utf-8").decode(buf);
        out = await parseCsvText(text);
      } else if (
        ct.includes(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ) ||
        ct.includes("application/vnd.ms-excel")
      ) {
        out = await parseWorkbookFromArrayBuffer(buf);
      } else {
        const sample = new TextDecoder("utf-8").decode(buf.slice(0, 2048));
        if (sample.toLowerCase().includes("<html"))
          throw new Error("Received HTML instead of file/zip.");
        out = await parseWorkbookFromArrayBuffer(buf);
      }
      cached = { at: now, rows: out.rows, idx: out.idx };
      return cached;
    } catch (e) {
      lastErr = e;
    }
  }
  throw new Error(
    `Unable to download/parse Box resource. Last error: ${
      lastErr?.message || lastErr
    }`
  );
}

// --- filter & sort (License # asc, then Holder) ---
function filterAndSort(rows, { q = "" }) {
  const needle = q.trim().toLowerCase();
  let out = rows;
  if (needle) {
    out = out.filter((r) => {
      const hay = [
        r.license_number,
        r.license_type,
        r.category,
        r.holder,
        r.dba,
        r.city,
      ]
        .join(" ")
        .toLowerCase();
      return hay.includes(needle);
    });
  }
  const toNum = (s) => {
    const n = parseInt(String(s || "").replace(/[^0-9]/g, ""), 10);
    return isNaN(n) ? Number.MAX_SAFE_INTEGER : n;
  };
  out.sort((a, b) => {
    const na = toNum(a.license_number),
      nb = toNum(b.license_number);
    if (na !== nb) return na - nb;
    return (a.holder || "").localeCompare(b.holder || "");
  });
  return out;
}

// --- handler ---
export async function handler(event) {
  try {
    const headers = {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Headers": "Content-Type",
      "Content-Type": "application/json; charset=utf-8",
    };
    if (event.httpMethod === "OPTIONS")
      return { statusCode: 200, headers, body: "" };

    const url = new URL(event.rawUrl);
    const q = url.searchParams.get("q") || "";
    const page = Math.max(
      1,
      parseInt(url.searchParams.get("page") || "1", 10)
    );
    const pageSize = Math.min(
      100,
      Math.max(10, parseInt(url.searchParams.get("pageSize") || "250", 10))
    );

    const { rows: all } = await fetchAndParse();
    const filtered = filterAndSort(all, { q });

    const total = filtered.length;
    const start = (page - 1) * pageSize;
    const slice = filtered.slice(start, start + pageSize);

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({ ok: true, total, page, pageSize, results: slice }),
    };
  } catch (err) {
    return {
      statusCode: 500,
      headers: { "Content-Type": "application/json; charset=utf-8" },
      body: JSON.stringify({ ok: false, error: err.message || String(err) }),
    };
  }
}
