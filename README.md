/* xlsx-lite.js — zero-dependency .xlsx reader for the Clinic Streak dashboard.
 *
 * Reads a Metabase export entirely in the browser with NO external library:
 *   1. Parse the ZIP central directory (an .xlsx is a ZIP of XML parts).
 *   2. Inflate the parts we need with the built-in DecompressionStream.
 *   3. Pull cells from the `Query result` worksheet + the shared-strings table.
 *   4. Map columns by header name and emit the same record shape as data.json.
 *
 * Works from a plain file:// double-click — no server, no CDN, no Python.
 * Requires a browser with DecompressionStream + async/await (all current
 * Chrome/Edge/Firefox/Safari). Exposes window.XlsxLite.read(arrayBuffer, name).
 */
(function (global) {
  "use strict";

  var WEEK_BASE_SERIAL = 45894;            // Excel serial == 2025-08-25
  var NEEDED = ["order_date", "drug_name", "clinic_name", "sales_rep", "subtotal", "status"];

  // ---- ZIP parsing -------------------------------------------------------
  // Minimal reader: find End-Of-Central-Directory, walk central directory,
  // then read each local file header to locate the compressed bytes.

  function u16(dv, o) { return dv.getUint16(o, true); }
  function u32(dv, o) { return dv.getUint32(o, true); }

  function findEOCD(dv) {
    // EOCD signature 0x06054b50, scan backwards (max comment 65535).
    var len = dv.byteLength;
    var min = Math.max(0, len - 65557);
    for (var i = len - 22; i >= min; i--) {
      if (u32(dv, i) === 0x06054b50) return i;
    }
    throw new Error("Not a valid .xlsx (no ZIP end record found).");
  }

  function listEntries(buf) {
    var dv = new DataView(buf);
    var eocd = findEOCD(dv);
    var count = u16(dv, eocd + 10);
    var cdOffset = u32(dv, eocd + 16);
    var entries = {};
    var p = cdOffset;
    for (var n = 0; n < count; n++) {
      if (u32(dv, p) !== 0x02014b50) break; // central dir header sig
      var method = u16(dv, p + 10);
      var compSize = u32(dv, p + 20);
      var nameLen = u16(dv, p + 28);
      var extraLen = u16(dv, p + 30);
      var commentLen = u16(dv, p + 32);
      var localHdr = u32(dv, p + 42);
      var name = utf8(new Uint8Array(buf, p + 46, nameLen));
      entries[name] = { method: method, compSize: compSize, localHdr: localHdr };
      p += 46 + nameLen + extraLen + commentLen;
    }
    return { buf: buf, entries: entries };
  }

  function rawBytes(zip, name) {
    var e = zip.entries[name];
    if (!e) return null;
    var dv = new DataView(zip.buf);
    // Local file header: name + extra lengths live at +26/+28.
    var lh = e.localHdr;
    if (u32(dv, lh) !== 0x04034b50) throw new Error("Bad local header for " + name);
    var nameLen = u16(dv, lh + 26);
    var extraLen = u16(dv, lh + 28);
    var start = lh + 30 + nameLen + extraLen;
    return { method: e.method, bytes: new Uint8Array(zip.buf, start, e.compSize) };
  }

  function utf8(bytes) {
    return new TextDecoder("utf-8").decode(bytes);
  }

  async function inflate(part) {
    if (part.method === 0) return part.bytes; // stored, no compression
    if (part.method !== 8) throw new Error("Unsupported ZIP method " + part.method);
    if (typeof DecompressionStream === "undefined") {
      throw new Error("Your browser lacks DecompressionStream; please use a current Chrome, Edge, Firefox, or Safari.");
    }
    var ds = new DecompressionStream("deflate-raw");
    var stream = new Response(part.bytes).body.pipeThrough(ds);
    var ab = await new Response(stream).arrayBuffer();
    return new Uint8Array(ab);
  }

  async function readText(zip, name) {
    var part = rawBytes(zip, name);
    if (!part) return null;
    return utf8(await inflate(part));
  }

  // ---- XML helpers -------------------------------------------------------

  function decodeEntities(s) {
    if (s.indexOf("&") === -1) return s;
    return s.replace(/&lt;/g, "<").replace(/&gt;/g, ">")
            .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
            .replace(/&#(\d+);/g, function (_, d) { return String.fromCharCode(+d); })
            .replace(/&#x([0-9a-fA-F]+);/g, function (_, h) { return String.fromCharCode(parseInt(h, 16)); })
            .replace(/&amp;/g, "&");
  }

  // Shared strings: each <si> may hold <t>..</t> or multiple <r><t>..</t></r>.
  function parseSharedStrings(xml) {
    var out = [];
    if (!xml) return out;
    var siRe = /<si>([\s\S]*?)<\/si>/g, m;
    var tRe = /<t[^>]*>([\s\S]*?)<\/t>/g;
    while ((m = siRe.exec(xml))) {
      var inner = m[1], t, parts = "";
      while ((t = tRe.exec(inner))) parts += t[1];
      tRe.lastIndex = 0;
      out.push(decodeEntities(parts));
    }
    return out;
  }

  function colLetters(ref) {
    // "AB12" -> "AB"
    var i = 0;
    while (i < ref.length && ref.charCodeAt(i) >= 65) i++;
    return ref.slice(0, i);
  }
  function colToIndex(letters) {
    var n = 0;
    for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n - 1;
  }

  // Resolve a worksheet cell's value to string|number given its type + shared strings.
  function cellValue(cellXml, type, shared) {
    var vm = /<v>([\s\S]*?)<\/v>/.exec(cellXml);
    if (!vm) {
      var inl = /<t[^>]*>([\s\S]*?)<\/t>/.exec(cellXml); // inline string
      return inl ? decodeEntities(inl[1]) : null;
    }
    var raw = vm[1];
    if (type === "s") return shared[+raw];
    if (type === "str") return decodeEntities(raw);
    if (type === "b") return raw === "1";
    return parseFloat(raw); // number (incl. date serials), default
  }

  // ---- main --------------------------------------------------------------

  function pickWorksheet(zip) {
    // Map sheet name -> r:id (workbook.xml), r:id -> target (rels).
    var wbName = "xl/workbook.xml";
    // we already have these inflated by caller; handled in read()
    return null;
  }

  async function read(arrayBuffer, sourceName) {
    var zip = listEntries(arrayBuffer);

    var wbXml = await readText(zip, "xl/workbook.xml");
    var relsXml = await readText(zip, "xl/_rels/workbook.xml.rels");
    if (!wbXml || !relsXml) throw new Error("Missing workbook parts in .xlsx.");

    // sheet name -> rId
    var sheets = {};
    var sRe = /<sheet\b[^>]*\/?>/g, sm;
    while ((sm = sRe.exec(wbXml))) {
      var nm = /name="([^"]+)"/.exec(sm[0]);
      var rid = /r:id="([^"]+)"/.exec(sm[0]);
      if (nm && rid) sheets[decodeEntities(nm[1])] = rid[1];
    }
    // rId -> target path
    var rels = {};
    var rRe = /<Relationship\b[^>]*\/?>/g, rm;
    while ((rm = rRe.exec(relsXml))) {
      var id = /Id="([^"]+)"/.exec(rm[0]);
      var tg = /Target="([^"]+)"/.exec(rm[0]);
      if (id && tg) rels[id[1]] = tg[1];
    }

    var sheetName = sheets["Query result"] ? "Query result"
      : Object.keys(sheets)[Object.keys(sheets).length - 1];
    if (!sheetName) throw new Error("No worksheets found.");
    var target = rels[sheets[sheetName]];
    if (!target) throw new Error("Could not resolve sheet '" + sheetName + "'.");
    var sheetPath = "xl/" + target.replace(/^\//, "").replace(/^xl\//, "");

    var shared = parseSharedStrings(await readText(zip, "xl/sharedStrings.xml"));
    var sheetXml = await readText(zip, sheetPath);
    if (!sheetXml) throw new Error("Could not read worksheet data.");

    // Walk rows.
    var rowRe = /<row\b[^>]*>([\s\S]*?)<\/row>/g, rowM;
    var cellRe = /<c\b([^>]*)(?:\/>|>([\s\S]*?)<\/c>)/g;

    var colIndex = null; // header-name -> 0-based column index
    var records = [], minW = null, maxW = null;
    var firstRow = true;

    while ((rowM = rowRe.exec(sheetXml))) {
      var rowInner = rowM[1];
      var cells = {}; // colIdx -> value
      var cm;
      cellRe.lastIndex = 0;
      while ((cm = cellRe.exec(rowInner))) {
        var attrs = cm[1], body = cm[2] || "";
        var ref = /r="([A-Z]+\d+)"/.exec(attrs);
        if (!ref) continue;
        var ci = colToIndex(colLetters(ref[1]));
        var tt = /t="([^"]+)"/.exec(attrs);
        cells[ci] = cellValue(body, tt ? tt[1] : null, shared);
      }

      if (firstRow) {
        firstRow = false;
        colIndex = {};
        Object.keys(cells).forEach(function (ci) {
          var name = cells[ci];
          if (typeof name === "string") colIndex[name.trim()] = +ci;
        });
        var missing = NEEDED.filter(function (n) {
          return n !== "status" && !(n in colIndex);
        });
        if (missing.length) {
          throw new Error("Sheet '" + sheetName + "' is missing columns: " +
            missing.join(", ") + ". Make sure you exported the 'Query result' tab.");
        }
        continue;
      }

      var serial = cells[colIndex.order_date];
      if (typeof serial !== "number" || !isFinite(serial)) continue;
      var w = Math.floor((serial - WEEK_BASE_SERIAL) / 7);

      var sub = cells[colIndex.subtotal];
      sub = (typeof sub === "number" && isFinite(sub)) ? sub :
            (sub == null || isNaN(parseFloat(sub)) ? 0 : parseFloat(sub));

      var statusCol = ("status" in colIndex) ? cells[colIndex.status] : "";

      records.push({
        rep: cells[colIndex.sales_rep] || "(no rep)",
        clinic: cells[colIndex.clinic_name] || "(no clinic)",
        drug: cells[colIndex.drug_name] || "(no drug)",
        week: w,
        subtotal: sub,
        status: statusCol || "",
        date: serialToISO(serial),
      });
      if (minW === null || w < minW) minW = w;
      if (maxW === null || w > maxW) maxW = w;
    }

    if (!records.length) throw new Error("No data rows with a valid order_date were found.");

    return {
      generated: new Date().toISOString(),
      source: sourceName || sheetName,
      sheet: sheetName,
      week_base_serial: WEEK_BASE_SERIAL,
      min_week: minW,
      max_week: maxW,
      records: records,
    };
  }

  function serialToISO(serial) {
    var ms = Date.UTC(1899, 11, 30) + Math.round(serial * 86400) * 1000;
    return new Date(ms).toISOString().slice(0, 10);
  }

  global.XlsxLite = { read: read };
})(typeof window !== "undefined" ? window : globalThis);
