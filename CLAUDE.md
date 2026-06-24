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
| `MainForm.NotamFilter.cs` | `Filter_Notams()`, `ICAO_Notams()`, all NOTAM classification logic, UI helpers |
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

### JavaScript ↔ C# Bridge (AIP SUP tab)

`Web_Sup_report` uses `ObjectForScripting = new AviobookScriptBridge(this)` so HTML checkboxes can call `window.external.UpdateAviobook(id)` directly into C# without redrawing the UI. The JS function `updateCheckbox(id, isYes)` updates only the affected checkbox in place.

### C# Version Constraint

Target is **.NET 4.0 / C# 4** (SharpDevelop portable). These are **not available**:
- `?.` null-conditional operator (use `x != null && x.Something`)
- `$""` string interpolation (use `string.Format` or concatenation)
- `nameof()`, `async/await`, LINQ on `ControlCollection` without `using System.Linq`

### Dynamic UI Controls

Controls created at runtime must have `Tag = "dispose"`. `ClearTaggedControls(panel)` removes and disposes them before redrawing. This pattern is used in all tabs that rebuild their content (Filter, ICAO Notams, Airport List, RWYs).
