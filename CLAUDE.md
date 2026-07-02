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
| `MainForm.NotamData.cs` | `GetCSV()`, `GetXML()`, `Split()`, `NewNotams()`, `DelOld()`, `deleteWithdrawnedNotams()`, the DB Update `BackgroundWorker` pipeline + progress dialog |
| `MainForm.NotamFilter.cs` | `Filter_Notams()`, `ICAO_Notams()`, NOTAM classification logic, keyword highlighting, runway-diagram (VML) builder, UI helpers |
| `MainForm.Reports.cs` | `Report()` — HTML NOTAM report with impact rows per operator |
| `MainForm.AipSup.cs` | `Sup_Report()` — AIP SUP HTML report (lists `Sup='Yes'` NOTAMs, shows `SupRef`), interactive HTML checkboxes |
| `MainForm.AirportList.cs` | `Airport_List()` — `DataGridView`-based airport CRUD (ICAO/IATA/name/LH/FedEx/Charters, search, inline add/edit/delete), airport-name CSV index, `ClearTaggedControls()` (shared helper, called from RWYs/NotamFilter too) |
| `MainForm.Rwys.cs` | structured `Runways` table (OCC.mdb), CSV auto-import, legacy-memo migration, `tab_RWYs()` DataGridView editor |
| `MainForm.Export.cs` | `ExportToPdf()` — calls `wkhtmltopdf.exe` to export HTML reports |
| `MainForm.Keywords.cs` | `EnsureKeywordsTable()`, `LoadKeywords()`, Keywords tab (add/remove highlight keywords) |
| `MainForm.Email.cs` | `EnsureEmailTable()`, recipients list, `Btn_sendReportsClick` (emails today's two PDFs via late-bound Outlook COM) |

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

### DB Update pipeline (`Btn_updateDBClick` / `Btn_dbUpdateQuick` on the Filter tab)

- `RunDbUpdatePipeline(onCompleted)` runs `GetXML → deleteWithdrawnedNotams → NewNotams → DelOld` on a **`BackgroundWorker`** (not the UI thread), reporting weighted progress (download jumps to ~2%, then 30/45/90% boundaries for the three DB-diff phases) into a single dark, non-modal progress dialog (`ShowDbProgressForm`/`UpdateDbProgress`/`CloseDbProgressForm`) — replacing the old sequence of blocking popups + a final `MessageBox`. `Btn_dbUpdateQuick` (top bar) and `Btn_updateDB` (DB Update tab) both call `RunDbUpdatePipeline`; the quick button's `onCompleted` refreshes the current Filter/Station view afterward.
- `NewNotams`/`deleteWithdrawnedNotams`/`DelOld` each take an optional `Action<int,string> onProgress` and open **one** OleDb connection for their whole body (previously `NewNotams` opened 2 + 2×N connections — one full-table scan with columns resolved by `GetOrdinal` name lookup, not hard-coded index, replaced the old per-key `SELECT * WHERE key=?` round trip). Diff/dedup logic uses an in-memory `HashSet<string>` of keys instead of one `SELECT COUNT(*)` query per row.
- Since these methods now run off the UI thread, writes to the debug textbox go through `SetLog()` (marshals via `Invoke` if needed) instead of `RchTxtCSV.Text = …` directly.
- `RefreshLastDbUpdateLabel` (top bar `Lbl_lastDbUpdate`) shows when the pipeline last **completed**, read from `last_db_update.txt` next to the exe (written by `SaveLastDbUpdateTimestamp`) — deliberately not the `.mdb` file's own last-write time, which also changes on V: drive sync (`StartApp`/`EndApp`) and would look "fresh" even if DB Update was never run.

### NOTAM Filter tab — two modes on `tabPage1`

The single **NOTAM Filter** tab (the old separate "Stations" tab was removed) has a top bar with two entry points, both rendering onto `tabPage1`:
- **`Filter New NOTAMS` button** (`Btn_filterNew` → `Filter_Notams`) — the auto triage below (airport-by-airport while unchecked NOTAMs remain; pending model + SUBMIT).
- **ICAO box + `OK`** (`TxtBox_ICAO` / `Btn_ICAO` → `ICAO_Notams`) — manual station view: Kept NOTAMs read-only on the left, **all** of the station's NOTAMs (kept + ignored + new) as interactive cards on the right with a Keep/Ignore button + immediate-write impact/SUP chips (`AddStationChips`). SUBMIT is hidden in this mode.

`_stationMode` tracks which view is active; `RefreshCurrentView()` re-renders the right one after Keep/Ignore/impact edits (so the left Kept list updates live). Both views share `BuildAirportCardHtml` and `RenderKeptCard`. The right column starts at `Top=48` to clear the top bar.

### Filter auto-triage (`Filter_Notams`)

- Runs **automatically once the form is shown** (`this.Shown += … Filter_Notams()`), **not** from the constructor — the `Shown` event fires after the window is maximized so the layout reads the real `tabPage1.ClientSize.Width` (the constructor size is the smaller design-time width). There is no "Analyze" button. It loads the first airport that still has unchecked NOTAMs (`Checked='N'`).
- The form is **maximized, `AutoScroll=false`**; `tabControl1` is `Dock=Fill`. Only `tabPage1` scrolls (vertical). Card/control-box/status-bar widths are computed from `tabPage1.ClientSize.Width - scrollbar` so nothing triggers a horizontal scrollbar.
- Layout per airport:
  - **`Web_FilterHeader`** (a `WebBrowser`) renders the dark "airport card": ICAO, IATA, `N new` badge, RWY blocks, and a **VML runway diagram** (oriented by QFU = designator × 10°, length proportional to distance, parallel runways offset perpendicular).
  - **Left column** — kept NOTAMs, each in a `Panel` (`Tag="dispose"`) holding a colored impact strip + a borderless `RichTextBox`.
  - **Right column** — each new/unchecked NOTAM is a **card**: a white `Panel` (`Tag="dispose"`) with a left strip in the impact colour, sent to the back *after* its content is added (z-order: `SendToBack()` only moves behind controls already present). On top sit the key/date labels, the mini `WebBrowser` body (`MakeNotamWebBrowser`), the flat Keep/Ignore button, and a light **control box** holding the impact chips + remark textboxes.
  - A dark **status bar** + styled `Btn_submitNotams` ("SUBMIT ▶") sit at the bottom of the right column, right-aligned to the card width.
- When **no unchecked NOTAMs remain**, `Filter_Notams` early-returns after rendering an "ALL NOTAMS CHECKED" card in `Web_FilterHeader` and hiding the Submit button.

#### Auto-classification engine (Filter tab only)

- `SuggestImpacts(text, runways, keptUpper, extraIlsDownRwys)` returns an `ImpactSuggestion` (flags A/C/N/D/F + independent `Sup`, plus two auto-Keep-only signals `RwyClosure`/`IlsOutage` — see below). The rules were **derived and validated against the dispatcher's historical classifications** in `ICAO_storedNotams.mdb` (recall on the sample: D 7/7, N 2/2, F 3/3, A 22/30, C 3/6), then iteratively refined against real false positives/negatives found in production use. `keptUpper` is retained in the signature but no longer used (rules are per-NOTAM).
  - **APT CLSD** fires on `AD CLSD`/`AERODROME CLOSED`, or when the NOTAM closes *every* runway of the airport (`ClosedThresholds` vs `ParseRunways` of the OCC RWY field) — a single runway closure with a parallel remaining is intentionally **not** flagged. `ClosedThresholds` requires the runway mention to carry an explicit `RWY`/`RUNWAY` keyword **and** `CLSD`/`CLOSED` within a ±20-character context window of that specific mention (`ContextWindow`) — a bare digit pair anywhere in the text, or a runway mentioned only as a location reference for an unrelated taxiway closure (e.g. "TWY R1 BTN RWY16R/34L AND TWY G"), is **not** treated as a closure. "AVBL ONLY FOR TAXI(ING)"/"NOT AVBL FOR TKOF/LDG" phrasing is treated as an equivalent closure even without the literal word CLSD.
  - **No ILS** fires on the shared `IlsOutagePattern` constant (`ILS … U/S|UNSERVICEABLE|NOT AVBL|NOT AVAILABLE|NOT USABLE|ON TEST|ON (FLT|FLIGHT) CALIBRATION|DO NOT USE`), excluding: the `EXC LVP` case (that's CAT I, not No ILS); a DME-only outage (`DME … ASSOCIATED WITH ILS … U/S` — a DME is a supporting component, losing it doesn't take the ILS itself out of service); and — like APT CLSD — suppressed if **another runway at the airport still has a working ILS** (`RwyInfo.CatMax >= 1`) not affected by this NOTAM. `extraIlsDownRwys` (built once per airport in `Filter_Notams` as `ilsDownUnion`, via `ExtractIlsDownRwys` — which also uses `IlsOutagePattern` — over every active NOTAM's text) folds in ILS outages reported by *other simultaneous* NOTAMs at the same airport, so e.g. two NOTAMs each reporting one end of the same runway's ILS "ON TEST" are recognised together as a full outage rather than each individually being suppressed by the other's (soon-to-be-down) runway.
  - **Not ALTN**: `\bPPR\b` excludes two false-positive shapes seen on NAVAID/TWY NOTAMs — a directly-following phone number ("PPR 617-561-1919", a ground-ops contact) or a short duration ("AVBL PPR 10MIN", a prior-notice window for degraded equipment) — neither is a real diversion/alternate restriction. An explicit **"EXCEPT ALTN"/"EXC ALTN"** carve-out (e.g. "RWY … NOT AVBL FOR LANDING, EXCEPT ALTN") suppresses the whole flag — the airport is still usable as a diversion despite the restriction.
  - **SUP** requires an actual `SUP nnn/yyyy` reference (`Regex.IsMatch`, same shape as `ExtractSupRef`) rather than a bare `Contains("SUP")` — the naive substring check was a false positive on any word containing "SUP" (e.g. "SECONDARY POWER **SUP**PLY").
- `SuggestedSingleCode` collapses the suggestion to one Impact code by severity (A>N>C>D>F) because `Impact` stores a single code; it ignores `RwyClosure`/`IlsOutage` (those never select a chip by themselves). **SUP is fully independent** — stored in the dedicated `Sup` column (`Yes`/empty), with its reference (e.g. `SUP 056/2026`, via `ExtractSupRef`) in a separate `SupRef` column. `Sup`/`SupRef`/**`AutoKept`** columns are added by `EnsureSchema()` (idempotent `ALTER TABLE`), which **also migrates the legacy SUP model** (`Impact='AS'` with the reference in `Remark`) to the new one (`Sup='Yes'`, ref in `SupRef`, `Impact`/`Remark` cleared). No `Impact='AS'` rows should remain after startup; `"AS"` survives only as the SUP chip's colour code.
- **Auto-Keep / Auto-Un-Keep** (`Filter_Notams` only): a pre-pass promotes not-yet-kept NOTAMs to `Status='K'` (and sets `AutoKept='Y'`) when `SuggestImpacts` finds any impact, SUP, `RwyClosure`, or `IlsOutage` — the latter two exist purely so a runway closure or ILS outage always surfaces for review even when no selectable impact chip applies (e.g. one runway closed/ILS down but an alternate remains) — so the chips (or, for these two, nothing) render without the dispatcher having to click Keep by hand. The same pass **demotes** (`Status=''`, `AutoKept=''`) a NOTAM it previously auto-Kept if a classification-rule change means it no longer triggers any signal (e.g. an older, looser rule wrongly auto-Kept a plain TWY/DVOR/DME NOTAM). **`AutoKept` is the persistent (survives app restarts) marker distinguishing an engine promotion from a dispatcher's manual Keep** — every manual-Keep path (`Keep_Notam`, `StationImpactSet`, `StationSupSet`, `Ignore_Notam`) explicitly clears it, so a NOTAM the dispatcher actually touched is *never* silently reverted by this pass, even before an impact has been picked for it. `Ignore_Notam` also adds the ID to the session-only `_autoKeepSkip` set so an ignored NOTAM is not immediately auto-re-kept on the next render.
- **Impact checkboxes are "chips"** (`CreateChip`): a `Panel` with a left accent strip + borderless `CheckBox` (the checkbox's `Tag` holds its accent panel). `StyleChk(cb, code)` tints the chip in the impact colour when checked (`ImpactColor` → `Tint`), grey accent when not. Columns/widths are laid out strictly inside the passed control-box geometry (`ctrlLeft`/`ctrlW`).
- **Two independent remark textboxes** per NOTAM (`LayoutRemarkBoxes`): the impact remark and the SUP-reference box. Only the box(es) whose checkbox is selected are shown; one box → full width, two → 2⁄3 impact + 1⁄3 SUP. The impact remark default (`NotamRemarkDefault`) is the NOTAM's first line, or the validity period `start - end` when that line is just `No`.
- **Visual-only until SUBMIT.** Nothing is written to the DB on click — `Btn_submitNotamsClick` reads the final state of `_pendImpactChks` / `_pendSupChk` / `_pendRemark` / `_pendSupRemark` and persists `Impact`, `Sup`, `Remark`, `SupRef` before marking `Checked='Y'`.
- The **station view (`ICAO_Notams`)** is **fully auto-save, no SUBMIT/SAVE button**: impact/SUP chips write immediately (`AddStationChips` → `StationImpactSet` / `StationSupSet`, both set `Status='K'` when assigned; SUP also writes `SupRef`), and the remark / SUP-ref textfields save on `Leave` (`StationSaveRemark`). Changing an impact resets `Remark` so the new impact shows its own default. Impact chips are shown **only for Kept NOTAMs** (`if (kept)`); ignored/new ones show just a Keep button. `impactOn` excludes `"AS"`. (`tabPage2`/`Web_ICAONotams`/`ChckBox_SeeIgnored` still exist in the Designer but are orphaned — the tab is no longer added to `tabControl1`.)
- **Top-bar layout** (`tabPage1`, all at y≈8): `Btn_filterNew` ("Filter New NOTAMS", green, x=7) → `Btn_dbUpdateQuick` ("DB Update", black, x=205, a quick-access duplicate of the DB Update tab's button) → `Lbl_lastDbUpdate` (gray "MAJ: …" timestamp, glued to `Btn_dbUpdateQuick.Right`) → `Lbl_ICAO` + `TxtBox_ICAO` + `Btn_ICAO` ("OK", blue-grey, x≈505+, left-aligned with the right NOTAM column, `Btn_ICAO` glued to `TxtBox_ICAO.Right`). The airport card (`Web_FilterHeader`) sits **below** the bar at `(7,44)`; the right column starts at `Top=48`.
- **`AlignTopBar()`** repositions all of the above with raw pixel values (matching the card column's hard-coded `Left=505`) at the very start of *every* `Filter_Notams`/`ICAO_Notams` render. This is necessary because these are Designer-time controls subject to WinForms DPI auto-scaling, while the dynamically-built NOTAM cards are placed with raw pixel math unaffected by auto-scale — without re-asserting their position every render, the two families drift apart on non-design-time DPI.
- **Scroll-position reset must happen *before* anything is (re)positioned**, as the very first statement of `Filter_Notams`/`ICAO_Notams` (`tabPage1.AutoScrollPosition = new Point(0, 0);`), not after. Positioning/adding controls to an `AutoScroll` panel while it's still scrolled bakes the current scroll offset into their effective location; resetting scroll afterward doesn't retroactively fix that. `TxtBox_ICAO` also responds to **Enter** (`TxtBox_ICAO_KeyDown` → `ICAO_Notams()`), and `Btn_filterNewClick` clears `TxtBox_ICAO.Text` when leaving the manual station view.
- **Keyword highlighting**: `_notamKeywords` are bolded/reddened (whole-word match only). `HighlightKeywords(rtb, startChar)` (RichTextBox) skips the first two lines (NOTAM ref + dates); `HighlightKeywordsHtml(text)` does the HTML equivalent for the mini WebBrowsers. The list is loaded from the `Keywords` table (`ICAO_storedNotams.mdb`, seeded with `_defaultKeywords` by `EnsureKeywordsTable()`) and editable via the **Keywords tab** — `_notamKeywords` is a runtime list, not a hardcoded constant.

### Runways data model (RWYs tab)

Per-runway data lives in a structured **`Runways`** table in `OCC.mdb` (`ICAO, QFU, Cat, DistM, Hdg, ThrLat, ThrLon, Ord`, created by `EnsureRunwaysTable()`). `Cat` (max approach type, e.g. `CAT 3`) is **manual** — the open-data CSVs do not include it; everything else is auto-filled from **`runways.csv`** (OurAirports data, copied next to the exe, scanned at runtime by ICAO). Adding an airport in APT List auto-imports its runways (`ImportRunwaysFromCsv` + `RegenerateRwyMemo`); existing airports migrate their legacy memo on first view (`MigrateLegacyMemo`, preserving the manual `Cat`, then `EnrichFromCsv`). The RWYs tab is a `DataGridView` editor (combo to pick the ICAO + Save / Re-import / Add RWY). **The legacy `Stations_ICAO_IATA.RWYs` memo is regenerated from the table** (`RegenerateRwyMemo`, lines `QFU: Cat DistMm`) so `ParseRunways` (ILS-downgrade context) keeps reading the memo unchanged.

**Airport diagram (`LoadRwyGeo` + `BuildRwySvgGeo`, `MainForm.NotamFilter.cs`)**: uses real threshold coordinates (equirectangular projection of `ThrLat`/`ThrLon`, true crossing angles and parallel spacing) when *at least one* threshold of the runway has coordinates (`HasGeo`/`HasCoords` — no longer all-or-nothing: a runway with a missing/zero threshold is skipped individually, projected from its other end's heading+length, rather than forcing the whole airport back to the schematic fallback `BuildRwySvg`, QFU×10 + length). `LoadRwyGeo` caches per-ICAO results in `_geoCache`, but **only once the result is non-empty** — an empty lookup (e.g. made before the CSV index below has finished loading) is left uncached so the next render retries instead of getting stuck on a stale empty result forever.

**CSV geo index (`MainForm.Rwys.cs`)**: `runways.csv` (~48k lines) is parsed once into an in-memory `_csvGeo` dictionary, loaded on a **background thread** (`PreloadCsvGeoAsync`, kicked off from the `MainForm` constructor) so the UI never blocks on it — `CsvGeoFor`/`EnsureCsvGeoLoaded` is thread-safe (lock + atomic dictionary swap) and returns empty (schematic fallback) if the render happens before the background load finishes. `ImportRunwaysFromCsv`/`MigrateLegacyMemo` (airport add, re-import, legacy migration) call `EnsureCsvGeoLoaded()` directly since those are occasional/one-off and can afford to block briefly.

### APT List tab (`DataGridView`, `MainForm.AirportList.cs`)

Rebuilt around a `DataGridView` (`_aptDgv`, same pattern as the RWYs tab's `_rwyDgv`) instead of hand-positioned dynamic controls:
- Columns: **ICAO** (bold), **IATA**, **Airport Name**, then checkbox columns **LH**/**FedEx**/**Charters** colour-coded to match the NOTAM/AIP SUP reports (RoyalBlue / purple `#663399` / SeaGreen — via `AptDgvCellPainting`, a `CellPainting` handler that custom-draws the checkbox glyph itself in the category colour while keeping the **cell background white**, not the reverse), and a **Del** button column.
- Checkbox columns commit **on click** (`CellContentClick` forces `_aptDgv.EndEdit()`) rather than waiting for focus to leave the cell, so toggling reads as instant — same feel as the AIP SUP report's checkboxes.
- **No Add/Edit form** — `AllowUserToAddRows` gives a perpetual blank row at the bottom; typing an ICAO into it triggers an INSERT (via `SaveAptRow`, fired from `CellValueChanged`) using `SELECT @@IDENTITY` to capture the new row's ID into `DataGridViewRow.Tag`, which every later edit/delete keys off. `_aptSuppressSave` guards programmatic `Rows.Add`/`Value` writes (e.g. during `LoadAptGrid` or auto-filling the name) from re-entering `SaveAptRow` through the same `CellValueChanged` event.
- **Airport name auto-fill**: `airports.csv` (OurAirports data, same file already shipped for runway geo) is parsed once into an ICAO → name index (`EnsureAirportNamesLoaded`, background-thread preload via `PreloadAirportNamesAsync`, quote-aware `CsvSplit` reused from `MainForm.Rwys.cs` since names can contain commas) — no network lookup. Existing rows with a blank `Name` are backfilled opportunistically every time the tab is loaded (`LoadAptGrid`); a brand-new row resolves its name synchronously (blocking `EnsureAirportNamesLoaded()` is fine here — an occasional one-off action, same precedent as the runway CSV import). The `Name` column on `Stations_ICAO_IATA` is added by `EnsureAirportNameColumn()` (idempotent `ALTER TABLE`, called from the constructor before `LoadStationsCache()`). `_stationsCache` (`MainForm.Stations.cs`) carries the name too (`GetOrdinal("Name")`, robust to column order) via the new `GetAirportName()` lookup, used by `BuildAirportCardHtml` to show the airport's full name as a line under "IATA: …" in the Filter tab's airport card.
- Search box (`_aptSearch`) filters ICAO/IATA live via `DataGridViewRow.Visible` (`FilterAptGrid`), client-side, no query re-run.
- `AptListTabEnter` (`APT_List.Enter`) reloads the grid every time the tab becomes active — not just once at startup — so names picked up by the background CSV index after the first visit (loading ~80k lines takes a few seconds) actually appear without the dispatcher having to edit a row to trigger a refresh.

### NOTAM Report / AIP SUP Report tabs — no button, auto-load on tab select

Both `Report()` (`MainForm.Reports.cs`) and `Sup_Report()` (`MainForm.AipSup.cs`) run automatically via `TabPage.Enter` (`TabPage4Enter` / `AipSupReportTabEnter`) — the reliable WinForms idiom for "this tab became active" (more so than `TabControl.SelectedIndexChanged`, which doesn't fire consistently in this app for tab-header clicks) — and re-run whenever the dispatcher changes the time-window radio (`RadBtn_reportWindowCheckedChanged` / `RadBtn_supReportWindowCheckedChanged`, wired to all four radios' `CheckedChanged`). There is no "Report !" button on either tab anymore.

The AIP SUP report's Avio checkboxes are **CSS-styled `<div>`s**, not native `<input type="checkbox">` — native checkboxes render as tiny, blurry glyphs through `wkhtmltopdf` (a different rendering engine than the app's `WebBrowser` preview); the div-based version (coloured square + `&#10003;` checkmark, toggled by `updateCheckbox()`/`window.external.UpdateAviobook`) renders identically in both. `ExportToPdf` (`MainForm.Export.cs`) uses `--dpi 96` (not a much higher value) — PDF text is vector-drawn regardless of DPI, so a very high `--dpi` with no matching `--zoom` just shrinks all CSS-px-sized content (text, checkboxes) tiny and blurry on the page.

### Window resize doesn't auto-relayout the Filter/Station cards

The NOTAM Filter/Station right-column cards are absolute-positioned (not docked/anchored) and compute their width from `tabPage1.ClientSize.Width` **once, at render time** — widening the window after that leaves them stuck at the old width. `MainForm.cs`'s constructor wires `this.Resize` to a debounced (`_resizeDebounce`, 200ms `Timer`) call to `RefreshCurrentView()`, so a drag-resize re-renders once after the size settles rather than on every intermediate frame (which would also re-query the DB repeatedly).

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

Controls created at runtime must have `Tag = "dispose"`. `ClearTaggedControls(panel)` (defined once, in `MainForm.AirportList.cs`; called from RWYs and NotamFilter too) removes and disposes them before redrawing. It handles `Label`, `TextBox`, `RichTextBox`, `CheckBox`, `Button`, `Panel`, `WebBrowser`, `DataGridView`, and `ComboBox` — **any new dynamic control type must be added to its `OfType<T>()` sweep**. This pattern is used in all tabs that rebuild their content (Filter, ICAO Notams, Airport List, RWYs).
