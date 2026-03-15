// ============================================================
//  KEYWORD ALERTS — Google Sheets News Tracker
//  Paste into Tools > Apps Script, Save, reload your Sheet.
//  Dropdowns for Filter Mode and When are applied automatically
//  every time the sheet is opened — no setup step needed.
// ============================================================

const KEYWORDS_SHEET      = "Keywords";
const RESULTS_SHEET       = "News Results";
const MAX_RESULTS         = 8;
const DEFAULT_FILTER_MODE = "EXACT";
const DEFAULT_WHEN        = "TODAY";

// ── ON OPEN — runs automatically every time sheet is opened ──
// Builds the menu AND applies dropdowns to Keywords sheet
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("🔔 Alerts")
      .addItem("▶ Refresh All Keywords",     "refreshAll")
      .addItem("▶ Refresh Selected Keyword", "refreshSelected")
      .addSeparator()
      .addItem("⏱ Auto-Refresh Every Hour",  "setHourlyTrigger")
      .addItem("⏱ Auto-Refresh Every Day",   "setDailyTrigger")
      .addItem("✖ Stop Auto-Refresh",        "clearTriggers")
      .addSeparator()
      .addItem("ℹ Filter Mode Guide",        "showFilterGuide")
      .addToUi();
  } catch (e) {
    Logger.log("onOpen: no UI context — " + e.message);
  }

  // Always ensure Keywords sheet exists with correct headers + dropdowns
  ensureKeywordsSheet_();
}

// ── ENSURE KEYWORDS SHEET ─────────────────────────────────────
// Creates the sheet if missing, writes headers, and applies
// dropdowns for Filter Mode (col B) and When (col C).
// Safe to call repeatedly — never overwrites existing keyword data.
function ensureKeywordsSheet_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Create sheet if it doesn't exist
    let kwSheet = ss.getSheetByName(KEYWORDS_SHEET);
    if (!kwSheet) {
      kwSheet = ss.insertSheet(KEYWORDS_SHEET);
    }

    // ── Write headers only if row 1 col A is empty
    const headerCheck = kwSheet.getRange("A1").getValue().toString().trim();
    if (headerCheck === "" || headerCheck.toLowerCase() === "keyword") {
      kwSheet.getRange("A1:D1")
        .setValues([["Keyword", "Filter Mode", "When", "Region"]])
        .setFontWeight("bold")
        .setBackground("#4A86E8")
        .setFontColor("#ffffff")
        .setHorizontalAlignment("center");

      // Column widths
      kwSheet.setColumnWidth(1, 240);
      kwSheet.setColumnWidth(2, 130);
      kwSheet.setColumnWidth(3, 130);
      kwSheet.setColumnWidth(4, 80);

      // Column background colours
      kwSheet.getRange("A2:A100").setBackground("#EAF1FB");
      kwSheet.getRange("B2:B100").setBackground("#FFF9E6");
      kwSheet.getRange("C2:C100").setBackground("#E8F5E9");
      kwSheet.getRange("D2:D100").setBackground("#F3E5F5");

      // Add sample rows only if column A is completely empty
      const existing = kwSheet.getRange("A2:A10").getValues().flat()
        .filter(function(v) { return v.toString().trim() !== ""; });
      if (existing.length === 0) {
        const samples = [
          ["Tamil Nadu news",  "EXACT",   "TODAY",   "IN"],
          ["AI regulation",    "PHRASE",  "ANYTIME", "US"],
          ["Bitcoin price",    "ALL",     "TODAY",   "IN"],
          ["climate tech",     "ANY",     "ANYTIME", ""],
        ];
        kwSheet.getRange(2, 1, samples.length, 4).setValues(samples);
      }
    }

    // ── Always re-apply dropdowns (survives accidental deletion)

    // Filter Mode dropdown — column B
    const modeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["EXACT", "PHRASE", "ALL", "ANY"], true)
      .setAllowInvalid(false)
      .setHelpText("EXACT = keyword must appear exactly\nPHRASE = whole-word match\nALL = all words must appear\nANY = at least one word")
      .build();
    kwSheet.getRange("B2:B100").setDataValidation(modeRule);

    // When dropdown — column C
    const whenRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["TODAY", "ANYTIME"], true)
      .setAllowInvalid(false)
      .setHelpText("TODAY = today's news only\nANYTIME = all results regardless of date")
      .build();
    kwSheet.getRange("C2:C100").setDataValidation(whenRule);

    // Freeze header row
    kwSheet.setFrozenRows(1);

  } catch (e) {
    Logger.log("ensureKeywordsSheet_: " + e.message);
  }
}

// ── REFRESH ALL ───────────────────────────────────────────────
function refreshAll() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const kwSheet = ss.getSheetByName(KEYWORDS_SHEET);
  if (!kwSheet) {
    showAlert_("Keywords sheet not found. Please reload the spreadsheet.");
    return;
  }

  const kwData = getKeywordData_(kwSheet);
  if (!kwData.length) {
    showAlert_("No keywords found. Add them in column A of the Keywords sheet.");
    return;
  }

  const resSheet = getOrCreateResultsSheet_(ss);
  const allRows  = [];

  kwData.forEach(function(item) {
    const articles = fetchNews_(item.keyword, item.mode, item.when, item.region);
    if (articles.length) {
      articles.forEach(function(a, i) {
        allRows.push([item.keyword, i + 1, a.title, a.source, a.date, a.url, item.mode, item.when]);
      });
    } else {
      const msg = item.when === "TODAY"
        ? "No matching results found for today"
        : "No matching results found";
      allRows.push([item.keyword, "", msg, "", "", "", item.mode, item.when]);
    }
  });

  writeResults_(resSheet, allRows);
  ss.setActiveSheet(resSheet);
}

// ── REFRESH SELECTED ──────────────────────────────────────────
function refreshSelected() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const active = ss.getActiveSheet();
  let kw = "", mode = DEFAULT_FILTER_MODE, when = DEFAULT_WHEN, region = "IN";

  if (active.getName() === KEYWORDS_SHEET) {
    const row = active.getActiveCell().getRow();
    if (row >= 2) {
      kw     = (active.getRange(row, 1).getValue() || "").toString().trim();
      mode   = norm_(active.getRange(row, 2).getValue(), DEFAULT_FILTER_MODE);
      when   = norm_(active.getRange(row, 3).getValue(), DEFAULT_WHEN);
      region = norm_(active.getRange(row, 4).getValue(), "IN");
    }
  } else if (active.getName() === RESULTS_SHEET) {
    const row = active.getActiveCell().getRow();
    if (row > 2) {
      kw   = (active.getRange(row, 1).getValue() || "").toString().trim();
      mode = norm_(active.getRange(row, 7).getValue(), DEFAULT_FILTER_MODE);
      when = norm_(active.getRange(row, 8).getValue(), DEFAULT_WHEN);
    }
  }

  if (!kw) {
    try {
      const response = SpreadsheetApp.getUi().inputBox("Enter a keyword to refresh:");
      if (response === SpreadsheetApp.getUi().Button.CANCEL) return;
      kw = response.toString().trim();
    } catch (e) { return; }
  }
  if (!kw) return;

  const resSheet = getOrCreateResultsSheet_(ss);
  const existing = getAllRows_(resSheet).filter(function(r) { return r[0] !== kw; });
  const articles = fetchNews_(kw, mode, when, region);
  const newRows  = articles.length
    ? articles.map(function(a, i) { return [kw, i + 1, a.title, a.source, a.date, a.url, mode, when]; })
    : [[kw, "", "No matching results found", "", "", "", mode, when]];

  writeResults_(resSheet, existing.concat(newRows));
  ss.setActiveSheet(resSheet);
}

// ── TRIGGERS ──────────────────────────────────────────────────
function setHourlyTrigger() {
  clearTriggers_();
  ScriptApp.newTrigger("refreshAll").timeBased().everyHours(1).create();
  showAlert_("✅ Auto-refresh set: every 1 hour.");
}

function setDailyTrigger() {
  clearTriggers_();
  ScriptApp.newTrigger("refreshAll").timeBased().everyDays(1).atHour(8).create();
  showAlert_("✅ Auto-refresh set: every day at 8 AM.");
}

function clearTriggers() {
  clearTriggers_();
  showAlert_("✅ Auto-refresh stopped.");
}

function clearTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
}

// ── FILTER MODE GUIDE ─────────────────────────────────────────
function showFilterGuide() {
  showAlert_(
    "🔍 Filter Mode Guide\n\n" +
    "EXACT\n" +
    "  Title must contain your keyword exactly as typed.\n" +
    "  'Tamil news' → matches 'Latest Tamil news today'\n" +
    "  but NOT 'Tamil Nadu news update'\n\n" +
    "PHRASE\n" +
    "  Whole-word boundary match.\n" +
    "  'nova' won't match 'innovation'\n" +
    "  Best for short or brand-style keywords.\n\n" +
    "ALL\n" +
    "  All words must appear in the title (any order).\n" +
    "  'Tamil floods' matches 'Floods reported in Tamil Nadu'\n\n" +
    "ANY\n" +
    "  At least one word must appear.\n" +
    "  Broadest — most results, more noise.\n\n" +
    "📅 When Filter\n\n" +
    "  TODAY   → Only today's articles\n" +
    "  ANYTIME → All articles regardless of date\n\n" +
    "💡 Tip: Use EXACT + TODAY for tight precise tracking.\n" +
    "   Use ANY + ANYTIME for broad research coverage."
  );
}

// ── CORE: FETCH + FILTER ──────────────────────────────────────
function fetchNews_(keyword, mode, when, region) {
  try {
    mode   = (mode   || DEFAULT_FILTER_MODE).toString().toUpperCase();
    when   = (when   || DEFAULT_WHEN).toString().toUpperCase();
    region = (region || "IN").toString().toUpperCase();

    const localeMap = {
      "IN": { hl: "en-IN", gl: "IN", lang: "en" },
      "US": { hl: "en-US", gl: "US", lang: "en" },
      "GB": { hl: "en-GB", gl: "GB", lang: "en" },
      "AU": { hl: "en-AU", gl: "AU", lang: "en" },
      "SG": { hl: "en-SG", gl: "SG", lang: "en" },
      "CA": { hl: "en-CA", gl: "CA", lang: "en" },
    };
    const locale = localeMap[region] || localeMap["IN"];
    const ceid   = locale.gl + ":" + locale.lang;

    const q   = encodeURIComponent('"' + keyword + '"');
    const url = "https://news.google.com/rss/search?q=" + q
              + "&hl=" + locale.hl
              + "&gl=" + locale.gl
              + "&ceid=" + ceid;

    const xml   = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText();
    const doc   = XmlService.parse(xml);
    const items = doc.getRootElement().getChild("channel").getChildren("item");

    const tz    = Session.getScriptTimeZone();
    const today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

    const results = [];

    for (let i = 0; i < items.length; i++) {
      if (results.length >= MAX_RESULTS) break;

      const item    = items[i];
      const pubDate = item.getChildText("pubDate") || "";

      // Filter 1: Date — TODAY skips anything not from today
      if (when === "TODAY" && pubDate) {
        const articleDate = Utilities.formatDate(new Date(pubDate), tz, "yyyy-MM-dd");
        if (articleDate !== today) continue;
      }

      const rawTitle = (item.getChildText("title") || "").replace(/ - [^-]+$/, "").trim();
      const link     = (item.getChildText("link")  || "").trim();
      const source   = item.getChild("source")
                         ? item.getChild("source").getText()
                         : extractDomain_(link);
      const date     = pubDate
                         ? Utilities.formatDate(new Date(pubDate), tz, "dd MMM yyyy HH:mm")
                         : "";

      // Filter 2: Title must match based on selected mode
      if (!titleMatches_(rawTitle, keyword, mode)) continue;

      results.push({ title: rawTitle, url: link, source: source, date: date });
    }

    return results;

  } catch (e) {
    Logger.log("fetchNews_ error [" + keyword + "]: " + e.message);
    return [];
  }
}

// ── TITLE MATCHING ENGINE ─────────────────────────────────────
function titleMatches_(title, keyword, mode) {
  const t  = title.toLowerCase().trim();
  const kw = keyword.toLowerCase().trim();

  switch (mode) {
    case "EXACT":
      return t.indexOf(kw) !== -1;

    case "PHRASE":
      const escaped = kw.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const pattern = new RegExp("(^|[^a-z0-9])" + escaped + "($|[^a-z0-9])");
      return pattern.test(t);

    case "ALL":
      const allWords = kw.split(/\s+/).filter(function(w) { return w.length > 1; });
      return allWords.every(function(w) { return t.indexOf(w) !== -1; });

    case "ANY":
      const anyWords = kw.split(/\s+/).filter(function(w) { return w.length > 1; });
      return anyWords.some(function(w) { return t.indexOf(w) !== -1; });

    default:
      return t.indexOf(kw) !== -1;
  }
}

// ── SHEET HELPERS ─────────────────────────────────────────────
function getKeywordData_(sheet) {
  const vals = sheet.getRange("A2:D100").getValues();
  return vals
    .filter(function(r) { return r[0].toString().trim().length > 0; })
    .map(function(r) {
      return {
        keyword: r[0].toString().trim(),
        mode:    norm_(r[1], DEFAULT_FILTER_MODE),
        when:    norm_(r[2], DEFAULT_WHEN),
        region:  norm_(r[3], "IN")
      };
    });
}

function norm_(val, fallback) {
  const s = (val || "").toString().trim().toUpperCase();
  return s.length > 0 ? s : fallback;
}

function getOrCreateResultsSheet_(ss) {
  let s = ss.getSheetByName(RESULTS_SHEET);
  if (!s) s = ss.insertSheet(RESULTS_SHEET);
  return s;
}

function styleResultsHeader_(sheet) {
  sheet.clearContents();
  sheet.clearFormats();
  const headers = ["Keyword", "#", "Title", "Source", "Published", "URL", "Filter Mode", "When"];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#4A86E8")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  [200, 40, 600, 150, 150, 320, 110, 90].forEach(function(w, i) {
    sheet.setColumnWidth(i + 1, w);
  });
  sheet.setFrozenRows(1);
  sheet.getRange("I1")
    .setValue("Last Refreshed")
    .setFontWeight("bold")
    .setBackground("#4A86E8")
    .setFontColor("#ffffff");
  sheet.setColumnWidth(9, 170);
}

function writeResults_(sheet, rows) {
  styleResultsHeader_(sheet);
  if (!rows.length) return;

  rows.sort(function(a, b) {
    return a[0].localeCompare(b[0]) || (a[1] || 0) - (b[1] || 0);
  });

  sheet.getRange(3, 1, rows.length, 8).setValues(rows);

  // Alternating colours per keyword group
  let lastKw = null, toggle = false;
  rows.forEach(function(r, i) {
    if (r[0] !== lastKw) { toggle = !toggle; lastKw = r[0]; }
    sheet.getRange(i + 3, 1, 1, 8).setBackground(toggle ? "#EAF1FB" : "#FFFFFF");
  });

  // Clickable URLs
  rows.forEach(function(r, i) {
    const url = r[5];
    if (url && url.toString().indexOf("http") === 0) {
      sheet.getRange(i + 3, 6)
        .setFormula('=HYPERLINK("' + url.toString().replace(/"/g, '""') + '","Open Article")')
        .setFontColor("#1155CC");
    }
  });

  // Bold first row of each keyword
  rows.forEach(function(r, i) {
    if (r[1] === 1 || r[1] === "") sheet.getRange(i + 3, 1).setFontWeight("bold");
  });

  // Colour-code Filter Mode (col G)
  const modeColors = { "EXACT": "#D0F0C0", "PHRASE": "#E3F2FD", "ALL": "#FFF3CD", "ANY": "#FCE4EC" };
  rows.forEach(function(r, i) {
    if (modeColors[r[6]]) {
      sheet.getRange(i + 3, 7).setBackground(modeColors[r[6]]).setHorizontalAlignment("center");
    }
  });

  // Colour-code When (col H)
  const whenColors = { "TODAY": "#E8F5E9", "ANYTIME": "#FFF3E0" };
  rows.forEach(function(r, i) {
    if (whenColors[r[7]]) {
      sheet.getRange(i + 3, 8).setBackground(whenColors[r[7]]).setHorizontalAlignment("center");
    }
  });

  // Timestamp
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy HH:mm:ss");
  sheet.getRange("I1").setValue(now).setFontColor("#ffffff").setBackground("#4A86E8");

  sheet.getRange(3, 1, rows.length, 8).setVerticalAlignment("middle").setWrap(false);
  sheet.setRowHeightsForced(3, rows.length, 24);
}

function getAllRows_(sheet) {
  const last = sheet.getLastRow();
  if (last < 3) return [];
  return sheet.getRange(3, 1, last - 2, 8).getValues().filter(function(r) { return r[0]; });
}

function showAlert_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

function extractDomain_(url) {
  try { return new URL(url).hostname.replace("www.", ""); } catch (e) { return ""; }
}
