# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Dispatch Watch** — a Windows Forms application (C# / .NET 4.0) used by ASL Aviation Group flight dispatchers to download, filter, classify, and report on ICAO NOTAMs. Built and maintained in **SharpDevelop portable** — Visual Studio is not available.

## Repo layout

The solution file sits at the repo root; the project and all sources live in a `Dispatch Watch/` subfolder:

```
Dispatch Watch.sln            ← repo root (open this)
Dispatch Watch/               ← project subfolder
   Dispatch Watch.csproj
   MainForm*.cs, Program.cs, icon.ico, …
CLAUDE.md
```

The compiled output is `Dispatch Watch.exe` (`<AssemblyName>`); the window title and `<ApplicationIcon>` (`icon.ico`, a crossing-runways mark) are also "Dispatch Watch".

## Build & Run

Open `Dispatch Watch.sln` in SharpDevelop. Build with **F8**, run with **F5**.

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
| `MainForm.Keywords.cs` | `EnsureKeywordsTable()`, `LoadKeywords()`, Keywords tab (add/remove highlight keywords) |

Any new partial file must be registered in `Dispatch Watch.csproj` with `<DependentUpon>MainForm.cs</DependentUpon>`.

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

- Runs **automatically once the form is shown** (`this.Shown += … Filter_Notams()`), **not** from the constructor — the `Shown` event fires after the window is maximized so the layout reads the real `tabPage1.ClientSize.Width` (the constructor size is the smaller design-time width). There is no "Analyze" button. It loads the first airport that still has unchecked NOTAMs (`Checked='N'`).
- The form is **maximized, `AutoScroll=false`**; `tabControl1` is `Dock=Fill`. Only `tabPage1` scrolls (vertical). Card/control-box/status-bar widths are computed from `tabPage1.ClientSize.Width - scrollbar` so nothing triggers a horizontal scrollbar.
- Layout per airport:
  - **`Web_FilterHeader`** (a `WebBrowser`) renders the dark "airport card": ICAO, IATA, `N new` badge, RWY blocks, and a **VML runway diagram** (oriented by QFU = designator × 10°, length proportional to distance, parallel runways offset perpendicular).
  - **Left column** — kept NOTAMs, each in a `Panel` (`Tag="dispose"`) holding a colored impact strip + a borderless `RichTextBox`.
  - **Right column** — each new/unchecked NOTAM is a **card**: a white `Panel` (`Tag="dispose"`) with a left strip in the impact colour, sent to the back *after* its content is added (z-order: `SendToBack()` only moves behind controls already present). On top sit the key/date labels, the mini `WebBrowser` body (`MakeNotamWebBrowser`), the flat Keep/Ignore button, and a light **control box** holding the impact chips + remark textboxes.
  - A dark **status bar** + styled `Btn_submitNotams` ("SUBMIT ▶") sit at the bottom of the right column, right-aligned to the card width.
- When **no unchecked NOTAMs remain**, `Filter_Notams` early-returns after rendering an "ALL NOTAMS CHECKED" card in `Web_FilterHeader` and hiding the Submit button.

#### Auto-classification engine (Filter tab only)

- `SuggestImpacts(text, runways, keptUpper)` returns an `ImpactSuggestion` (flags A/C/N/D/F + independent `Sup`) from regex on the NOTAM text plus airport context (`ParseRunways` of the OCC RWY field + the upper-cased texts of all Kept NOTAMs). The contextual rules (APT CLSD if all RWYs closed, CAT I if no other CAT 2/3, No ILS if no other ILS) are **best-effort regex** and meant to be tuned against real NOTAM samples.
- `SuggestedSingleCode` collapses the suggestion to one Impact code by severity (A>N>C>D>F) because `Impact` stores a single code. **SUP is fully independent** — stored in the dedicated `Sup` column (`Yes`/empty), with its reference (e.g. `SUP 056/2026`, via `ExtractSupRef`) in a separate `SupRef` column. Both columns are added by `EnsureSchema()` (idempotent `ALTER TABLE`).
- **Auto-Keep**: a pre-pass promotes not-yet-kept NOTAMs to `Status='K'` when `SuggestImpacts` finds any impact or SUP, so the chips render. `Ignore_Notam` adds the ID to the session `_autoKeepSkip` set so an ignored NOTAM is not auto-re-kept.
- **Impact checkboxes are "chips"** (`CreateChip`): a `Panel` with a left accent strip + borderless `CheckBox` (the checkbox's `Tag` holds its accent panel). `StyleChk(cb, code)` tints the chip in the impact colour when checked (`ImpactColor` → `Tint`), grey accent when not. Columns/widths are laid out strictly inside the passed control-box geometry (`ctrlLeft`/`ctrlW`).
- **Two independent remark textboxes** per NOTAM (`LayoutRemarkBoxes`): the impact remark and the SUP-reference box. Only the box(es) whose checkbox is selected are shown; one box → full width, two → 2⁄3 impact + 1⁄3 SUP. The impact remark default (`NotamRemarkDefault`) is the NOTAM's first line, or the validity period `start - end` when that line is just `No`.
- **Visual-only until SUBMIT.** Nothing is written to the DB on click — `Btn_submitNotamsClick` reads the final state of `_pendImpactChks` / `_pendSupChk` / `_pendRemark` / `_pendSupRemark` and persists `Impact`, `Sup`, `Remark`, `SupRef` before marking `Checked='Y'`.
- The **ICAO lookup tab (tabPage2)** keeps the original immediate-write behaviour via `AddImpactCheckboxes` / `Impact_Notam` — the suggestion engine, chips, and auto-keep are not applied there.
- **Keyword highlighting**: `_notamKeywords` are bolded/reddened (whole-word match only). `HighlightKeywords(rtb, startChar)` (RichTextBox) skips the first two lines (NOTAM ref + dates); `HighlightKeywordsHtml(text)` does the HTML equivalent for the mini WebBrowsers. The list is loaded from the `Keywords` table (`ICAO_storedNotams.mdb`, seeded with `_defaultKeywords` by `EnsureKeywordsTable()`) and editable via the **Keywords tab** — `_notamKeywords` is a runtime list, not a hardcoded constant.

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
