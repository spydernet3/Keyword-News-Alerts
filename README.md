# 🔔 Keyword Alerts — Google Sheets + Apps Script

> Track any keyword across Google News directly inside Google Sheets.  
> No API key. No sign-in. Completely free.

---

## ✨ Features

- Fetches today's news for each keyword from Google News RSS
- Writes results into a formatted **News Results** sheet
- **4 filter modes** per keyword — EXACT, PHRASE, ALL, ANY
- **When filter** — TODAY (today's articles only) or ANYTIME (all dates)
- Per-keyword region control (IN, US, GB, AU, etc.)
- Auto-refresh via hourly or daily time-based triggers
- Dropdowns auto-applied on sheet open — no manual setup needed
- Colour-coded results sheet with clickable article links

---

## 📁 File

```
KeywordAlerts.gs
```

Paste this single file into Google Apps Script. That's all you need.

---

## 🚀 Setup

**Step 1 — Open Apps Script**

In your Google Sheet click:
```
Extensions → Apps Script
```
Delete any existing code, paste the full contents of `KeywordAlerts.gs`, and press `Ctrl+S`.

**Step 2 — Reload your Sheet**

Close the Apps Script tab, go back to your Sheet and reload the page.  
A **🔔 Alerts** menu will appear in the top menu bar.

**Step 3 — Authorize (first time only)**

Click **🔔 Alerts → Refresh All Keywords**.  
Google will show a permissions screen:
```
Advanced → Go to project (unsafe) → Allow
```
> This only appears once. It is safe — you are the developer running your own script.

**Step 4 — Add your keywords**

A **Keywords** sheet is created automatically every time you open the spreadsheet.  
Fill in your keywords row by row:

| Column | Name | Options | Description |
|---|---|---|---|
| A | Keyword | Any text | The topic or phrase to track |
| B | Filter Mode | `EXACT` / `PHRASE` / `ALL` / `ANY` | How strictly to match the keyword in article titles |
| C | When | `TODAY` / `ANYTIME` | Today's articles only, or all dates |
| D | Region | `IN` `US` `GB` `AU` `SG` `CA` | Country code for localised results |

> **Columns B and C have dropdown selectors** — click any cell and pick from the list.  
> Dropdowns are re-applied automatically every time the sheet is opened.

---

## 🔍 Filter Modes

| Mode | Behaviour | Example |
|---|---|---|
| `EXACT` | Title must contain keyword exactly as typed | `AI regulation` matches "AI regulation bill passed" but NOT "AI regulations" |
| `PHRASE` | Whole-word boundary match — prevents partial hits | `nova` will NOT match "innovation" |
| `ALL` | All words must appear in title (any order) | `Tamil floods` matches "Floods hit Tamil Nadu today" |
| `ANY` | At least one word must appear in title | Broadest — most results, most noise |

> 💡 Start with `EXACT` for precise tracking. Switch to `ALL` if you get too few results.

---

## 📅 When Filter

| Value | Behaviour |
|---|---|
| `TODAY` | Only articles published today — like Google Alerts "As it happens" |
| `ANYTIME` | All articles regardless of publication date |

---

## 📋 Menu Options

| Menu Item | What it does |
|---|---|
| ▶ Refresh All Keywords | Fetches news for all keywords and writes to the News Results sheet |
| ▶ Refresh Selected Keyword | Click a keyword row first, then run this to refresh only that keyword |
| ⏱ Auto-Refresh Every Hour | Sets a background trigger to run Refresh All every 60 minutes |
| ⏱ Auto-Refresh Every Day | Sets a daily trigger at 8 AM in your script timezone |
| ✖ Stop Auto-Refresh | Removes all auto-refresh triggers |
| ℹ Filter Mode Guide | Shows a popup explanation of all filter modes |

---

## 📊 Results Sheet — Columns

| Column | Content |
|---|---|
| A — Keyword | The keyword this article belongs to |
| B — # | Article number within the keyword group |
| C — Title | Article headline (source suffix stripped) |
| D — Source | Publisher name |
| E — Published | Publication date and time |
| F — URL | Clickable "Open Article" hyperlink |
| G — Filter Mode | Colour-coded badge — EXACT (green) / PHRASE (blue) / ALL (amber) / ANY (pink) |
| H — When | Colour-coded badge — TODAY (green) / ANYTIME (amber) |
| I — Last Refreshed | Timestamp of the most recent refresh |

---

## 🌍 Supported Region Codes

| Code | Region |
|---|---|
| `IN` | India (default) |
| `US` | United States |
| `GB` | United Kingdom |
| `AU` | Australia |
| `SG` | Singapore |
| `CA` | Canada |

---

## ⚠️ Common Issues

| Issue | Fix |
|---|---|
| `Cannot call getUi()` error | Normal — happens when a trigger runs with no browser open. Handled silently. No action needed. |
| "Google hasn't verified this app" | Click **Advanced → Go to project (unsafe) → Allow**. Safe because you are the developer. |
| No results found for today | Switch `When` to `ANYTIME` or `Filter Mode` to `ANY` to broaden the search. |
| Old articles appearing | Make sure `When` column is set to `TODAY` for that keyword row. |
| Wrong / unrelated articles | Switch `Filter Mode` from `ANY` to `EXACT` or `PHRASE` for tighter matching. |

---

## 🔧 How Filtering Works

Two layers of filtering run before any article appears in results:

**Layer 1 — RSS query level**  
The keyword is wrapped in quotes before sending to Google News (e.g. `"AI regulation"`), so Google itself returns only closely matched results.

**Layer 2 — Title check**  
Each returned article title is checked against your keyword using the selected Filter Mode. Articles that don't pass the check are silently discarded.

---

## 📌 Notes

- No API key required — uses Google News RSS via Apps Script's built-in `UrlFetchApp`
- No third-party service — everything runs inside Google's infrastructure
- Free — runs within Google Apps Script's free quota limits
- Data source — Google News RSS, the same feed that powers Google Alerts

---

## 📄 License

MIT — free to use, modify, and distribute.
