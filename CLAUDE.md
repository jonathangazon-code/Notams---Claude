# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A Windows Forms application (C# / .NET 4.0) used by ASL Aviation Group flight dispatchers to download, filter, classify, and report on ICAO NOTAMs. Built and maintained in **SharpDevelop portable** — Visual Studio is not available.

## Build & Run

Open `ICAO-CSV.sln` in SharpDevelop. Build with **F8**, run with **F5**.

There are no automated tests, no lint tools, and no CLI build commands — SharpDevelop is the only build environment.

**Runtime requirements:**
- Windows only (WinForms + OleDb)
- `wkhtmltopdf.exe` must be placed next to the compiled `.exe` for PDF export to work
- Two Access databases must be present at runtime:
  - `ICAO_storedNotams.mdb` — NOTAMs, filter results, impact classification
  - `OCC.mdb` — station mapping (ICAO ↔ IATA, operator types, RWY info)

## Architecture

### Partial Class Split

`MainForm` is split across multiple partial class files — all compile into one class:

| File | Responsibility |
|------|---------------|
| `MainForm.cs` | Constructor, `StartApp`/`EndApp` (DB file sync to V: drive), button event handlers (one-liners) |
| `MainForm.Designer.cs` | Auto-generated WinForms layout — edit via SharpDevelop designer |
| `MainForm.Stations.cs` | `_stationsCache` (static Dictionary), `LoadStationsCache()`, `IsOpsType()`, `GetIATA()` |
| `MainForm.DateUtils.cs` | `MonthAbbrev()`, `dateTransformation()` |
| `MainForm.NotamData.cs` | `GetCSV()`, `GetXML()`, `Split()`, `NewNotams()`, `DelOld()`, `deleteWithdrawnedNotams()` |
| `MainForm.NotamFilter.cs` | `Filter_Notams()`, `ICAO_Notams()`, NOTAM classification logic, keyword highlighting, runway-diagram (VML) builder, UI helpers |
| `MainForm.Reports.cs` | `Report()` — HTML NOTAM report with impact rows per operator |
| `MainForm.AipSup.cs` | `Sup_Report()` — AIP SUP HTML report with interactive HTML checkboxes |
| `MainForm.AirportList.cs` | `Airport_List()`, add/edit/delete airport station entries |
| `MainForm.Rwys.cs` | `tab_RWYs()`, `Update_RWYs()` — runway info per station |
| `MainForm.Export.cs` | `ExportToPdf()` — calls `wkhtmltopdf.exe` to export HTML reports |

Any new partial file must be registered in `ICAO-CSV.csproj` with `<DependentUpon>MainForm.cs</DependentUpon>`.

### Database Access Rules

- Provider: `Microsoft.JET.OLEDB.4.0` (32-bit, x86 target)
- OleDb uses **positional `?` placeholders** — never named `@param`
- Always `Parameters.AddWithValue("?", value)` for every `?` in the query, in order
- Use `ExecuteNonQuery()` for INSERT/UPDATE/DELETE; `ExecuteReader()` for SELECT
- Open connection, execute, close immediately — no connection pooling pattern

### Key Data Flow

1. **GetCSV** → downloads ICAO CSV → saved as `ICAO-CSV.csv`
2. **Split** → parses CSV, inserts rows into `storedNotams_table`
3. **NewNotams** / **DelOld** → syncs `storedNotams_table` → `filteredNotams_table`
4. **Filter_Notams** → dispatcher classifies each NOTAM (Keep/Ignore, Impact type, Remark)
5. **Report** / **Sup_Report** → generates HTML → displayed in `WebBrowser` control → Print or PDF export

### Filter New Notams tab (`Filter_Notams`)

- Runs **automatically at startup** (called from the `MainForm` constructor) — there is no "Analyze" button. It loads the first airport that still has unchecked NOTAMs (`Checked='N'`).
- Layout per airport:
  - **`Web_FilterHeader`** (a `WebBrowser`) renders the dark "airport card": ICAO, IATA, `N new` badge, RWY blocks, and a **VML runway diagram** (oriented by QFU = designator × 10°, length proportional to distance, parallel runways offset perpendicular).
  - **Left column** — kept NOTAMs, each in a `Panel` (`Tag="dispose"`) holding a colored impact strip + a borderless `RichTextBox`.
  - **Right column** — new/unchecked NOTAMs, each NOTAM body rendered in a **mini `WebBrowser`** (`MakeNotamWebBrowser`), plus impact checkboxes/Keep button (unchanged WinForms controls).
  - A dark **status bar** + styled `Btn_submitNotams` ("SUBMIT ▶") sit at the bottom of the right column.
- When **no unchecked NOTAMs remain**, `Filter_Notams` early-returns after rendering an "ALL NOTAMS CHECKED" card in `Web_FilterHeader` and hiding the Submit button.

#### Auto-classification engine (Filter tab only)

- `SuggestImpacts(text, runways, keptUpper)` returns an `ImpactSuggestion` (flags A/C/N/D/F + independent `Sup`) from regex on the NOTAM text plus airport context (`ParseRunways` of the OCC RWY field + the upper-cased texts of all Kept NOTAMs). The contextual rules (APT CLSD if all RWYs closed, CAT I if no other CAT 2/3, No ILS if no other ILS) are **best-effort regex** and meant to be tuned against real NOTAM samples.
- `SuggestedSingleCode` collapses the suggestion to one Impact code by severity (A>N>C>D>F) because `Impact` stores a single code; **SUP is stored independently** in the dedicated `Sup` column (`Yes`/empty) added by `EnsureSchema()` (idempotent `ALTER TABLE`, run from the constructor).
- **Checkboxes are visual-only until SUBMIT.** `AddFilterCheckboxes` pre-checks the suggested box in an "AUTO" colour (yellow for impact, green for SUP) when nothing is stored yet. The impact group behaves as radio buttons (`FilterImpactToggled`); SUP is independent. Selecting an impact pre-fills the remark with the NOTAM's first line if empty. Nothing is written to the DB on click — `Btn_submitNotamsClick` reads the final state of `_pendImpactChks` / `_pendSupChk` / `_pendRemark` and persists `Impact`, `Sup`, `Remark` before marking `Checked='Y'`.
- The **ICAO lookup tab (tabPage2)** keeps the original immediate-write behaviour via `AddImpactCheckboxes` / `Impact_Notam` — the suggestion engine is not applied there.
- **Keyword highlighting**: `_notamKeywords` are bolded/reddened. `HighlightKeywords(rtb, startChar)` (RichTextBox) skips the first two lines (NOTAM ref + dates); `HighlightKeywordsHtml(text)` does the HTML equivalent for the mini WebBrowsers.

### WebBrowser control runs in IE7 mode

The embedded `WebBrowser` renders in legacy IE quirks mode: **inline SVG is not supported** — use **VML** instead (`xmlns:v="urn:schemas-microsoft-com:vml"` + `v\:*{behavior:url(#default#VML)}`). Prefer `<table>`/`float` over modern fl/grid layout; `inline-block` is unreliable.

### JavaScript ↔ C# Bridge (AIP SUP tab)

`Web_Sup_report` uses `ObjectForScripting = new AviobookScriptBridge(this)` so HTML checkboxes can call `window.external.UpdateAviobook(id)` directly into C# without redrawing the UI. The JS function `updateCheckbox(id, isYes)` updates only the affected checkbox in place.

### C# Version Constraint

Target is **.NET 4.0 / C# 4** (SharpDevelop portable). These are **not available**:
- `?.` null-conditional operator (use `x != null && x.Something`)
- `$""` string interpolation (use `string.Format` or concatenation)
- `nameof()`, `async/await`, LINQ on `ControlCollection` without `using System.Linq`

### Dynamic UI Controls

Controls created at runtime must have `Tag = "dispose"`. `ClearTaggedControls(panel)` removes and disposes them before redrawing. It handles `Label`, `TextBox`, `RichTextBox`, `CheckBox`, `Button`, `Panel`, and `WebBrowser` — **any new dynamic control type must be added to its `OfType<T>()` sweep**. This pattern is used in all tabs that rebuild their content (Filter, ICAO Notams, Airport List, RWYs).
