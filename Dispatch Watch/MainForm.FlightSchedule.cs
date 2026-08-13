using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		private DataGridView _fsDgv;
		private Label _fsCsvInfoLabel;
		private TextBox[] _fsFilterBoxes;
		private const int _fsDgvRowHeaderWidth = 28;
		private System.ComponentModel.BackgroundWorker _fsWorker;
		private Form _fsProgressForm;
		private Label _fsProgressStatus;
		private ProgressBar _fsProgressBar;

		private static readonly XNamespace FsNs = "http://www.fwz.aero/Schemas/StandardInterfaces";

		// Flight schedule cache in ICAO_storedNotams.mdb — one row per Fltleg_ID. STA is
		// the expensive field (one getBriefing call per flight), so it's kept here and
		// only re-fetched when a flight's STD has actually changed since last cached.
		// [Source] distinguishes a webservice-backed row ('WS') from a temporary one built
		// from the CSV fallback ('CSV', see LoadCsvSupplementalFlights) — old rows predate
		// this column and read back as "", which is treated as 'WS' everywhere below.
		public void EnsureFlightScheduleTable()
		{
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				try { new OleDbCommand("CREATE TABLE FlightSchedule ([FltlegID] LONG, [FLDt] TEXT(10), [Callsign] TEXT(20), [Reg] TEXT(10), [STD] TEXT(25), [STA] TEXT(25), [Crew] TEXT(255))", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }
				try { new OleDbCommand("ALTER TABLE FlightSchedule ADD COLUMN Source TEXT(3)", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }
				try { new OleDbCommand("ALTER TABLE FlightSchedule ADD COLUMN Origin TEXT(4)", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }
				try { new OleDbCommand("ALTER TABLE FlightSchedule ADD COLUMN Dest TEXT(4)", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }
				// Filed alternates + flight-time-in-minutes, read out of the same getBriefing
				// XML STA already comes from (FetchBriefingDetails) — only ever populated once a
				// flight plan has actually been built, same as STA, and only for Source='WS' rows
				// (MM/CSV rows have no Fltleg_ID to brief).
				try { new OleDbCommand("ALTER TABLE FlightSchedule ADD COLUMN Alt1 TEXT(4)", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }
				try { new OleDbCommand("ALTER TABLE FlightSchedule ADD COLUMN Alt2 TEXT(4)", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }
				try { new OleDbCommand("ALTER TABLE FlightSchedule ADD COLUMN FlightTimeMin LONG", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }
				try { new OleDbCommand("ALTER TABLE FlightSchedule ADD COLUMN Alt1TimeMin LONG", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }
				try { new OleDbCommand("ALTER TABLE FlightSchedule ADD COLUMN Alt2TimeMin LONG", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }
				conn.Close();
			}
			catch { }
		}

		// Where the dispatcher drops the FPM "Refuel EU" CSV export as a stand-in for
		// flights the FlightScheduleService.svc feed doesn't expose yet. A shared V:
		// location (Admin-tab editable, _flightSchedCsvPath in MainForm.Admin.cs) — not
		// Application.StartupPath-relative, since each dispatcher now runs their own local
		// install (MainForm.Deployment.cs) and a CSV dropped into a local folder would only
		// ever be visible to that one dispatcher. The exact filename doesn't matter, the most
		// recently modified *.csv in the folder is used.
		private static string FlightSchedCsvFolder { get { return _flightSchedCsvPath; } }

		public void EnsureFlightSchedFolder()
		{
			try { Directory.CreateDirectory(FlightSchedCsvFolder); } catch { }
		}

		void FlightScheduleTabEnter(object sender, EventArgs e)
		{
			if (_fsDgv == null) BuildFlightScheduleGrid();
			LoadFlightScheduleGrid();        // always show what's already stored — cheap, local-only
			TryAutoRefreshFlightSchedule();  // possibly kick a background refresh if stale (see below)
		}

		// Session-local timestamp of the last *successful* refresh completion, manual or
		// automatic alike (set in RefreshFlightSchedule()'s RunWorkerCompleted below) — the
		// single throttle every automatic trigger (opening this tab, opening the Conflict tab —
		// NOTAM triage completion used to trigger it too, removed since a refresh shouldn't fire
		// on a run where the dispatcher never actually opens either tab) checks against, so a
		// full network refresh (a getFlightList call per day 0-7, the MM backlog, the CSV pass,
		// one getBriefing call per flight) isn't re-run on every tab click when the data is only
		// seconds old.
		private static DateTime _lastFsRefreshCompletedUtc = DateTime.MinValue;

		// Silent, throttled auto-trigger — deliberately does NOT call EnsureWriterOrWarn (that
		// pops a blocking "Reader mode" MessageBox, fine for a deliberate button click but not
		// for something that fires just from opening a tab); a Reader simply gets no automatic
		// refresh, same as if nothing happened.
		private void TryAutoRefreshFlightSchedule()
		{
			if (!IsWriter) return;
			if (_fsWorker != null && _fsWorker.IsBusy) return;
			if (DateTime.UtcNow - _lastFsRefreshCompletedUtc < TimeSpan.FromMinutes(5)) return;
			RefreshFlightSchedule();
		}

		private void BuildFlightScheduleGrid()
		{
			ClearTaggedControls(tabPage_FlightSchedule);

			// Fixed-height top bar (Dock=Top) + the grid filling everything below
			// (Dock=Fill) — the grid previously sat at a hard-coded Top/Size, which meant
			// rows past the tab's visible height (and the grid's own vertical scrollbar
			// along with them) were simply clipped, with no way to reach them regardless
			// of window size. Dock=Fill makes the grid always occupy the full remaining
			// client area and scroll internally once rows overflow it.
			// Per-column filter row: a column-name label + textbox per column, positioned/
			// sized to line up with the grid's own columns below (RowHeadersWidth +
			// cumulative column widths), AND-combined across columns in FilterFsGrid. Built
			// (and added to the tab) BEFORE topBar: WinForms stacks same-Dock=Top controls
			// with the LAST one added ending up closest to the parent's top edge, so adding
			// this first is what puts it BELOW the title bar rather than above it.
			string[] colNames = { "Date", "Callsign", "Origin", "Dest", "Reg", "STD (UTC)", "STA (UTC)",
				"Flight Time", "Alt 1", "Time to Alt 1", "Alt 2", "Time to Alt 2", "Crew" };
			int[] colWidths   = { 90, 100, 70, 70, 80, 140, 140, 90, 70, 100, 70, 100, 380 };
			Panel filterRow = new Panel { Tag = "dispose", Dock = DockStyle.Top, Height = 42 };
			_fsFilterBoxes = new TextBox[colWidths.Length];
			int filterLeft = _fsDgvRowHeaderWidth;
			for (int i = 0; i < colWidths.Length; i++)
			{
				Label colLbl = new Label { Top = 2, Left = filterLeft, Width = colWidths[i] - 3, AutoSize = false,
					Font = new Font("Microsoft Sans Serif", 7.5f), ForeColor = Color.DimGray, Text = colNames[i] };
				filterRow.Controls.Add(colLbl);
				TextBox tb = new TextBox { Top = 18, Left = filterLeft, Width = colWidths[i] - 3 };
				tb.TextChanged += delegate { FilterFsGrid(); };
				filterRow.Controls.Add(tb);
				_fsFilterBoxes[i] = tb;
				filterLeft += colWidths[i];
			}
			tabPage_FlightSchedule.Controls.Add(filterRow);

			// Fixed-height top bar (Dock=Top) + the grid filling everything below
			// (Dock=Fill) — the grid previously sat at a hard-coded Top/Size, which meant
			// rows past the tab's visible height (and the grid's own vertical scrollbar
			// along with them) were simply clipped, with no way to reach them regardless
			// of window size. Dock=Fill makes the grid always occupy the full remaining
			// client area and scroll internally once rows overflow it.
			Panel topBar = new Panel { Tag = "dispose", Dock = DockStyle.Top, Height = 50 };
			Label hdr = new Label { Top = 14, Left = 20, AutoSize = true,
				Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold), Text = "Schedule next 14 days (UTC)" };
			topBar.Controls.Add(hdr);
			Button refresh = new Button { Top = 10, Left = 400, Width = 90, Height = 28, Text = "Refresh" };
			refresh.Click += delegate { RefreshFlightSchedule(); };
			topBar.Controls.Add(refresh);
			Button exportPdf = new Button { Top = 10, Left = 498, Width = 100, Height = 28, Text = "Export PDF" };
			exportPdf.Click += Btn_fsExportReportClick;
			topBar.Controls.Add(exportPdf);
			_fsCsvInfoLabel = new Label { Top = 18, Left = 610, AutoSize = true, ForeColor = Color.Gray, Text = "" };
			topBar.Controls.Add(_fsCsvInfoLabel);
			tabPage_FlightSchedule.Controls.Add(topBar);

			_fsDgv = new DataGridView { Tag = "dispose", Dock = DockStyle.Fill,
				ReadOnly = true, AllowUserToAddRows = false, RowHeadersWidth = _fsDgvRowHeaderWidth, BackgroundColor = Color.White,
				AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None };
			_fsDgv.ColumnHeadersDefaultCellStyle.Font = new Font(_fsDgv.Font, FontStyle.Bold);
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Date",     HeaderText = "Date",      Width = colWidths[0] });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Callsign", HeaderText = "Callsign",  Width = colWidths[1] });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Origin",   HeaderText = "Origin",    Width = colWidths[2] });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Dest",     HeaderText = "Dest",      Width = colWidths[3] });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Reg",      HeaderText = "Reg",       Width = colWidths[4] });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "STD",      HeaderText = "STD (UTC)", Width = colWidths[5] });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "STA",         HeaderText = "STA (UTC)",      Width = colWidths[6] });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "FlightTime",  HeaderText = "Flight Time",    Width = colWidths[7] });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Alt1",        HeaderText = "Alt 1",          Width = colWidths[8] });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Alt1Time",    HeaderText = "Time to Alt 1",  Width = colWidths[9] });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Alt2",        HeaderText = "Alt 2",          Width = colWidths[10] });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Alt2Time",    HeaderText = "Time to Alt 2",  Width = colWidths[11] });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Crew",        HeaderText = "Crew",           Width = colWidths[12] });
			// Hidden identity column (not part of colNames/colWidths — the filter row above is
			// strictly 1:1 with that array) so the "Force Conflict" button handler below knows
			// which flight it's acting on without needing a separate Row.Tag convention.
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "FltlegID", Visible = false });
			// Opens ShowForceConflictDialog to manually assign/clear a Kept NOTAM as a forced
			// conflict for this flight — for false negatives the automatic Origin/Dest/Alt1/Alt2
			// matching on the Conflict tab can't express. Per-row button text (set in
			// LoadFlightScheduleGrid) shows the assigned NOTAM's key, or "+ Force" if none.
			_fsDgv.Columns.Add(new DataGridViewButtonColumn { Name = "ForceConflict", HeaderText = "Force Conflict",
				Text = "+ Force", UseColumnTextForButtonValue = false, Width = 110 });
			_fsDgv.CellContentClick += FsDgv_CellContentClick;
			// Dock=Fill controls occupy remaining space in add order, so the grid must be
			// added after the Top-docked bars.
			tabPage_FlightSchedule.Controls.Add(_fsDgv);
		}

		// Downloads getFlightList for today..today+7, resolves STA via a cached getBriefing
		// call per flight (skipped when the flight's STD hasn't moved since the last
		// fetch), stores everything in FlightSchedule and reloads the grid sorted by STD.
		// Runs on a BackgroundWorker with a progress dialog — a 7-day window can mean a few
		// hundred individual getBriefing calls.
		public void RefreshFlightSchedule()
		{
			if (!EnsureWriterOrWarn()) return;
			if (_fsWorker != null && _fsWorker.IsBusy) return;

			EnsureFlightScheduleTable();
			EnsureArchiveConfig();
			ShowFsProgressForm();

			_fsWorker = new System.ComponentModel.BackgroundWorker { WorkerReportsProgress = true };
			_fsWorker.DoWork += delegate(object s, System.ComponentModel.DoWorkEventArgs e)
			{
				System.ComponentModel.BackgroundWorker w = (System.ComponentModel.BackgroundWorker)s;
				FetchFlightSchedule(delegate(int pct, string msg) { w.ReportProgress(pct, msg); });
			};
			_fsWorker.ProgressChanged += delegate(object s, System.ComponentModel.ProgressChangedEventArgs e)
			{
				UpdateFsProgress(e.ProgressPercentage, e.UserState as string);
			};
			_fsWorker.RunWorkerCompleted += delegate(object s, System.ComponentModel.RunWorkerCompletedEventArgs e)
			{
				if (e.Error != null)
				{
					CloseFsProgressForm();
					MessageBox.Show("Error while loading the flight schedule:\n" + e.Error.Message, "Flight Schedule Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
				LoadFlightScheduleGrid();
				// Auto-save the Flight Schedule PDF as the last step of every refresh (manual or
				// automatic alike) — folded into this same progress dialog's status line rather
				// than a second, separate confirmation popup (ExportToPdf's own success/failure
				// MessageBox is suppressed via silent=true). Wrapped so a PDF failure (missing
				// wkhtmltopdf.exe, V: unreachable) can't take down the refresh completion itself.
				try
				{
					UpdateFsProgress(100, "Saving Flight Schedule PDF...");
					ExportToPdf(BuildFlightScheduleHtml(), "Flight_Schedule", true);
				}
				catch { }
				CloseFsProgressForm();
				_lastFsRefreshCompletedUtc = DateTime.UtcNow;
				// LoadFlightScheduleGrid() above already keeps the Flight Schedule tab's own grid
				// current regardless of which tab triggered this refresh. The Conflict tab isn't
				// self-updating the same way — Build_Conflict_Report() only runs on TabPage.Enter,
				// so if the dispatcher is sitting on Conflict while a refresh it kicked off (or
				// one kicked off elsewhere, e.g. finishing NOTAM triage) completes, re-render it
				// now so the fresh flight data actually shows up without switching tabs away and back.
				if (tabControl1.SelectedTab == tabPage_Conflict) Build_Conflict_Report();
			};
			_fsWorker.RunWorkerAsync();
		}

		private static string El(XElement parent, XName name)
		{
			XElement e = parent.Element(name);
			return e != null ? e.Value : "";
		}

		// departureAerodrome/arrivalAerodrome are nested elements ({iataID, icaoID}), not
		// flat fields — this reads the IATA code out of one of those two nested elements.
		private static string NestedIata(XElement flight, XName aerodromeName)
		{
			XElement aerodrome = flight.Element(aerodromeName);
			return aerodrome != null ? El(aerodrome, FsNs + "iataID") : "";
		}

		// Some stations are the same physical airport under two different IATA codes
		// (e.g. EuroAirport Basel-Mulhouse-Freiburg: BSL on the Swiss side, MLH on the
		// French side) — different sources/messages report one or the other for what's
		// otherwise an identical flight, which otherwise reads as two distinct duplicate
		// legs. Normalized to a single canonical code wherever Origin/Dest is set, across
		// all three sources (webservice, MM, CSV).
		private static readonly Dictionary<string, string> _iataAliasMap = new Dictionary<string, string> { { "BSL", "MLH" } };
		private static string NormalizeIata(string iata)
		{
			string canon;
			string upper = (iata ?? "").Trim().ToUpper();
			return _iataAliasMap.TryGetValue(upper, out canon) ? canon : upper;
		}

		// A flight-leg's identity across all three sources (webservice/MM/CSV) — used to
		// detect a lower-priority placeholder that has since "graduated" to a
		// higher-priority source. Deliberately keyed by Callsign+Origin+Dest+day (not the
		// exact STD): the same callsign can fly two different legs the same day (e.g.
		// CDG-HEL then HEL-TLL), so Origin/Dest keeps those distinct, but a schedule
		// revision of a few minutes between two updates for the *same* leg must NOT be
		// read as a different flight — an earlier STD-based key caused exactly that: a
		// stale MM/CSV placeholder never graduated because its STD differed by ~10min from
		// the webservice's later, revised STD for the same leg, leaving both rows behind.
		private static string MergeKey(string callsign, string origin, string dest, string fldt)
		{
			return (callsign ?? "").Trim().ToUpper() + "|" + (origin ?? "").Trim().ToUpper() + "|" +
				(dest ?? "").Trim().ToUpper() + "|" + (fldt ?? "").Trim().Substring(0, Math.Min(10, (fldt ?? "").Trim().Length));
		}

		// FltlegID -> last-cached STD plus everything FetchBriefingDetails resolved for it.
		private struct CachedBriefing { public string Std; public BriefingDetails Details; }

		private void FetchFlightSchedule(Action<int, string> onProgress)
		{
			Dictionary<int, CachedBriefing> cached = new Dictionary<int, CachedBriefing>();
			OleDbConnection rconn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			rconn.Open();
			OleDbDataReader rdr = new OleDbCommand(
				"SELECT FltlegID, STD, STA, Alt1, Alt2, FlightTimeMin, Alt1TimeMin, Alt2TimeMin FROM FlightSchedule", rconn).ExecuteReader();
			while (rdr.Read())
			{
				if (rdr.IsDBNull(0)) continue;
				int id = Convert.ToInt32(rdr.GetValue(0));
				CachedBriefing cb = new CachedBriefing();
				cb.Std = rdr.IsDBNull(1) ? "" : rdr.GetString(1);
				cb.Details.Sta           = rdr.IsDBNull(2) ? "" : rdr.GetString(2);
				cb.Details.Alt1          = rdr.IsDBNull(3) ? "" : rdr.GetString(3);
				cb.Details.Alt2          = rdr.IsDBNull(4) ? "" : rdr.GetString(4);
				cb.Details.FlightTimeMin = rdr.IsDBNull(5) ? 0  : Convert.ToInt32(rdr.GetValue(5));
				cb.Details.Alt1TimeMin   = rdr.IsDBNull(6) ? 0  : Convert.ToInt32(rdr.GetValue(6));
				cb.Details.Alt2TimeMin   = rdr.IsDBNull(7) ? 0  : Convert.ToInt32(rdr.GetValue(7));
				cached[id] = cb;
			}
			rconn.Close();

			// FltlegID, FLDt, Callsign, Reg, STD, Crew, Source ("WS"/"CSV"), precomputed STA
			// (CSV rows already carry their own STA — no getBriefing call needed/possible
			// since they have no real Fltleg_ID yet).
			List<object[]> flights = new List<object[]>();
			HashSet<string> wsKeys = new HashSet<string>();   // Callsign|STD already covered by the webservice
			HashSet<int> wsIdsSeen = new HashSet<int>();      // FltlegIDs the webservice actually returned this run

			for (int dayOffset = 0; dayOffset <= 7; dayOffset++)
			{
				DateTime day = DateTime.UtcNow.Date.AddDays(dayOffset);
				string fldt = day.ToString("ddMMMyy", CultureInfo.InvariantCulture).ToUpper();
				onProgress(dayOffset * 10 / 16, "Downloading flight list for " + fldt + "...");

				string xml;
				try
				{
					using (WebClient wc = new WebClient())
						xml = wc.DownloadString(_flightScheduleBaseUrl.TrimEnd('/') + "/?METHOD=getFlightList&FLDT=" + fldt);
				}
				catch { continue; }   // a single day's failure shouldn't abort the whole refresh

				XDocument doc;
				try { doc = XDocument.Parse(xml); } catch { continue; }

				foreach (XElement flight in doc.Descendants(FsNs + "Flight"))
				{
					int fltlegId;
					if (!int.TryParse(El(flight, FsNs + "Fltleg_ID"), out fltlegId) || fltlegId == 0) continue;

					string callsign = El(flight, FsNs + "FPfx") + El(flight, FsNs + "FLNr");
					if (!CallsignAllowed(callsign)) continue;   // e.g. only TAY/FDX/DHL, set on the Admin tab

					string reg      = El(flight, FsNs + "ACREG");
					string std      = El(flight, FsNs + "STD");
					string fldtVal  = El(flight, FsNs + "FLDt");
					string origin   = NormalizeIata(NestedIata(flight, FsNs + "departureAerodrome"));
					string dest     = NormalizeIata(NestedIata(flight, FsNs + "arrivalAerodrome"));
					// getFlightList also returns placeholder/technical entries alongside the
					// real flight — same STD/ACREG as a genuine flight but no route and no
					// crew, FLNr defaulting to "3" (hence the recurring bogus "TAY3" rows).
					// A real flight always has both ends of its route; anything without one
					// isn't a flight the dispatcher needs to see.
					if (origin == "" && dest == "") continue;

					List<string> crewParts = new List<string>();
					XElement crewMembers = flight.Element(FsNs + "CrewMembers");
					if (crewMembers != null)
						foreach (XElement cm in crewMembers.Elements(FsNs + "crewMember"))
						{
							string name = El(cm, FsNs + "name");
							string function = El(cm, FsNs + "function");
							if (name != "") crewParts.Add(name + " (" + function + ")");
						}
					string crew = string.Join(" / ", crewParts.ToArray());

					flights.Add(new object[] { fltlegId, fldtVal, callsign, reg, std, crew, "WS", null, origin, dest });
					wsKeys.Add(MergeKey(callsign, origin, dest, fldtVal));
					wsIdsSeen.Add(fltlegId);
				}
			}

			onProgress(45, "Reading Movement Manager XML messages...");
			HashSet<int> mmIdsSeen = new HashSet<int>();
			HashSet<string> mmKeys;
			foreach (object[] mmFlight in LoadMmSupplementalFlights(wsKeys, mmIdsSeen, out mmKeys))
				flights.Add(mmFlight);

			onProgress(50, "Reading FlightSched CSV fallback...");
			HashSet<int> csvIdsSeen = new HashSet<int>();
			foreach (object[] csvFlight in LoadCsvSupplementalFlights(wsKeys, mmKeys, csvIdsSeen))
				flights.Add(csvFlight);

			OleDbConnection wconn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			wconn.Open();

			int total = flights.Count, done = 0;
			foreach (object[] f in flights)
			{
				int    fltlegId = (int)f[0];
				string fldtVal  = (string)f[1];
				string callsign = (string)f[2];
				string reg      = (string)f[3];
				string std      = (string)f[4];
				string crew     = (string)f[5];
				string source   = (string)f[6];
				string csvSta   = (string)f[7];
				string origin   = (string)f[8];
				string dest     = (string)f[9];

				CachedBriefing cache;
				bool exists = cached.TryGetValue(fltlegId, out cache);
				string sta;
				BriefingDetails details;
				if (source == "CSV")
				{
					sta = csvSta;   // no Fltleg_ID to brief yet — STA comes straight from the CSV
					details = new BriefingDetails { Alt1 = "", Alt2 = "" };
				}
				else if (exists && cache.Std == std && cache.Details.Sta != "")
				{
					// getBriefing legitimately 404s/errors for flights whose briefing hasn't
					// been generated yet (e.g. far-future legs) — FetchBriefingDetails already
					// swallows that, and the flight is still persisted with blanks rather than
					// being dropped from the list.
					sta = cache.Details.Sta;
					details = cache.Details;
				}
				else
				{
					details = FetchBriefingDetails(fltlegId);
					sta = details.Sta;
				}

				try
				{
				if (exists)
				{
					OleDbCommand upd = new OleDbCommand("UPDATE FlightSchedule SET FLDt=?, Callsign=?, Reg=?, STD=?, STA=?, Crew=?, Source=?, Origin=?, Dest=?, Alt1=?, Alt2=?, FlightTimeMin=?, Alt1TimeMin=?, Alt2TimeMin=? WHERE FltlegID=?", wconn);
					upd.Parameters.AddWithValue("?", fldtVal);
					upd.Parameters.AddWithValue("?", callsign);
					upd.Parameters.AddWithValue("?", reg);
					upd.Parameters.AddWithValue("?", std);
					upd.Parameters.AddWithValue("?", sta);
					upd.Parameters.AddWithValue("?", crew);
					upd.Parameters.AddWithValue("?", source);
					upd.Parameters.AddWithValue("?", origin);
					upd.Parameters.AddWithValue("?", dest);
					upd.Parameters.AddWithValue("?", details.Alt1);
					upd.Parameters.AddWithValue("?", details.Alt2);
					upd.Parameters.AddWithValue("?", details.FlightTimeMin);
					upd.Parameters.AddWithValue("?", details.Alt1TimeMin);
					upd.Parameters.AddWithValue("?", details.Alt2TimeMin);
					upd.Parameters.AddWithValue("?", fltlegId);
					upd.ExecuteNonQuery();
				}
				else
				{
					OleDbCommand ins = new OleDbCommand("INSERT INTO FlightSchedule ([FltlegID],[FLDt],[Callsign],[Reg],[STD],[STA],[Crew],[Source],[Origin],[Dest],[Alt1],[Alt2],[FlightTimeMin],[Alt1TimeMin],[Alt2TimeMin]) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", wconn);
					ins.Parameters.AddWithValue("?", fltlegId);
					ins.Parameters.AddWithValue("?", fldtVal);
					ins.Parameters.AddWithValue("?", callsign);
					ins.Parameters.AddWithValue("?", reg);
					ins.Parameters.AddWithValue("?", std);
					ins.Parameters.AddWithValue("?", sta);
					ins.Parameters.AddWithValue("?", crew);
					ins.Parameters.AddWithValue("?", source);
					ins.Parameters.AddWithValue("?", origin);
					ins.Parameters.AddWithValue("?", dest);
					ins.Parameters.AddWithValue("?", details.Alt1);
					ins.Parameters.AddWithValue("?", details.Alt2);
					ins.Parameters.AddWithValue("?", details.FlightTimeMin);
					ins.Parameters.AddWithValue("?", details.Alt1TimeMin);
					ins.Parameters.AddWithValue("?", details.Alt2TimeMin);
					ins.ExecuteNonQuery();
				}
				}
				catch { /* one bad row shouldn't drop the rest of the week */ }

				done++;
				onProgress(50 + done * 45 / Math.Max(1, total), "Loading briefing " + done + " / " + total + "...");
			}

			// Normalize any already-stored row still holding a pre-alias-map IATA code
			// (e.g. "BSL" for EuroAirport Basel-Mulhouse-Freiburg) unconditionally, every
			// refresh — NormalizeIata only ever touches freshly-fetched data, so a row this
			// run's fetch doesn't happen to re-touch (a CSV row, or an MM row whose source
			// message has scrolled out of the backlog) would otherwise keep its old value
			// forever, and the dedup pass below compares stored values as-is, so two rows
			// that "should" be the same leg never actually matched keys.
			foreach (string alias in _iataAliasMap.Keys)
			{
				string canon = _iataAliasMap[alias];
				OleDbCommand fixOrigin = new OleDbCommand("UPDATE FlightSchedule SET Origin=? WHERE Origin=?", wconn);
				fixOrigin.Parameters.AddWithValue("?", canon);
				fixOrigin.Parameters.AddWithValue("?", alias);
				fixOrigin.ExecuteNonQuery();
				OleDbCommand fixDest = new OleDbCommand("UPDATE FlightSchedule SET Dest=? WHERE Dest=?", wconn);
				fixDest.Parameters.AddWithValue("?", canon);
				fixDest.Parameters.AddWithValue("?", alias);
				fixDest.ExecuteNonQuery();
			}

			// Drop stale CSV placeholders: ones that graduated to a real webservice row
			// (merge key now covered by a "WS" row — the CSV row is superseded and would
			// otherwise sit alongside it as a duplicate) or ones simply no longer present
			// in the latest CSV file (cancelled/dropped from the export).
			List<int> toDrop = new List<int>();
			OleDbDataReader csvRdr = new OleDbCommand("SELECT FltlegID, Callsign, Origin, Dest, FLDt FROM FlightSchedule WHERE Source='CSV'", wconn).ExecuteReader();
			while (csvRdr.Read())
			{
				int id = Convert.ToInt32(csvRdr.GetValue(0));
				string csvCallsign = csvRdr.IsDBNull(1) ? "" : csvRdr.GetString(1);
				string csvOrigin   = csvRdr.IsDBNull(2) ? "" : csvRdr.GetString(2);
				string csvDest     = csvRdr.IsDBNull(3) ? "" : csvRdr.GetString(3);
				string csvFldt     = csvRdr.IsDBNull(4) ? "" : csvRdr.GetString(4);
				if (wsKeys.Contains(MergeKey(csvCallsign, csvOrigin, csvDest, csvFldt)) || !csvIdsSeen.Contains(id))
					toDrop.Add(id);
			}
			csvRdr.Close();

			// Also drop any previously-stored row (either source) whose callsign no longer
			// matches the Admin tab's filter — covers rows fetched before the filter was
			// set/changed, since the day/CSV loops above already skip disallowed callsigns
			// going forward but don't touch what's already in the table.
			OleDbDataReader allRdr = new OleDbCommand("SELECT FltlegID, Callsign FROM FlightSchedule", wconn).ExecuteReader();
			while (allRdr.Read())
			{
				int id = Convert.ToInt32(allRdr.GetValue(0));
				string callsign = allRdr.IsDBNull(1) ? "" : allRdr.GetString(1);
				if (!CallsignAllowed(callsign) && !toDrop.Contains(id)) toDrop.Add(id);
			}
			allRdr.Close();

			// Purge already-stored ghost/placeholder rows from before this guard existed —
			// getFlightList's routeless "TAY3"-style entries (see above) may already sit in
			// the table from earlier refreshes.
			OleDbDataReader routeRdr = new OleDbCommand("SELECT FltlegID, Origin, Dest FROM FlightSchedule", wconn).ExecuteReader();
			while (routeRdr.Read())
			{
				int id = Convert.ToInt32(routeRdr.GetValue(0));
				string o = routeRdr.IsDBNull(1) ? "" : routeRdr.GetString(1);
				string d = routeRdr.IsDBNull(2) ? "" : routeRdr.GetString(2);
				if (o == "" && d == "" && !toDrop.Contains(id)) toDrop.Add(id);
			}
			routeRdr.Close();

			// Purge stale webservice rows: FPM sometimes reissues a new Fltleg_ID for a
			// flight that already has a row from an earlier refresh (a crew update or a
			// schedule shift of a few minutes), leaving the old ID's row behind forever —
			// it never gets updated (different key) and nothing previously deleted it,
			// producing duplicate-looking rows for the "same" flight. Any Source='WS' row
			// (or legacy blank-Source row, treated the same) whose FltlegID the webservice
			// didn't return in *this* run's day+0..+7 window is either superseded or simply
			// no longer scheduled, so it's dropped. 'MM' rows are handled by their own pass
			// below (wsIdsSeen only ever contains real webservice IDs).
			OleDbDataReader wsRdr = new OleDbCommand("SELECT FltlegID, Source FROM FlightSchedule", wconn).ExecuteReader();
			while (wsRdr.Read())
			{
				int id = Convert.ToInt32(wsRdr.GetValue(0));
				string source = wsRdr.IsDBNull(1) ? "" : wsRdr.GetString(1);
				if (source != "CSV" && source != "MM" && !wsIdsSeen.Contains(id) && !toDrop.Contains(id)) toDrop.Add(id);
			}
			wsRdr.Close();

			// Same idea for Source='MM' rows: a merge key that graduated to the webservice
			// (now in wsKeys) drops its MM placeholder, and a FltlegID (synthetic, MmSyntheticId)
			// that this run's MM catch-up didn't (re)encounter — cancelled, or simply no
			// longer within the message backlog window — is dropped too.
			OleDbDataReader mmRdr = new OleDbCommand("SELECT FltlegID, Callsign, Origin, Dest, FLDt FROM FlightSchedule WHERE Source='MM'", wconn).ExecuteReader();
			while (mmRdr.Read())
			{
				int id = Convert.ToInt32(mmRdr.GetValue(0));
				string mmCallsign = mmRdr.IsDBNull(1) ? "" : mmRdr.GetString(1);
				string mmOrigin   = mmRdr.IsDBNull(2) ? "" : mmRdr.GetString(2);
				string mmDest     = mmRdr.IsDBNull(3) ? "" : mmRdr.GetString(3);
				string mmFldt     = mmRdr.IsDBNull(4) ? "" : mmRdr.GetString(4);
				if ((wsKeys.Contains(MergeKey(mmCallsign, mmOrigin, mmDest, mmFldt)) || !mmIdsSeen.Contains(id)) && !toDrop.Contains(id))
					toDrop.Add(id);
			}
			mmRdr.Close();

			// Purge flights whose STD has already passed — the grid should only ever show
			// upcoming (or currently in-progress) flights.
			DateTime nowUtc = DateTime.UtcNow;
			OleDbDataReader pastRdr = new OleDbCommand("SELECT FltlegID, STD FROM FlightSchedule", wconn).ExecuteReader();
			while (pastRdr.Read())
			{
				int id = Convert.ToInt32(pastRdr.GetValue(0));
				string stdRaw = pastRdr.IsDBNull(1) ? "" : pastRdr.GetString(1);
				DateTime stdDt;
				if (DateTime.TryParse(stdRaw, CultureInfo.InvariantCulture,
					DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out stdDt)
					&& stdDt < nowUtc && !toDrop.Contains(id))
					toDrop.Add(id);
			}
			pastRdr.Close();

			// General duplicate cleanup, across ALL sources: rows that share the same
			// flight-leg identity (Callsign+Origin+Dest+day) are the same physical flight
			// no matter how they ended up duplicated (e.g. stale rows left over from
			// before MergeKey was fixed to ignore exact STD, or a tactical-vs-real callsign
			// mismatch that has since been corrected). Keep the single best row per
			// identity — preferring a non-blank Crew, then Source priority WS > MM > CSV,
			// then a non-blank STA — and drop the rest.
			Dictionary<string, int[]> bestBySameLeg = new Dictionary<string, int[]>();   // key -> [bestId, bestScore]
			OleDbDataReader dedupRdr = new OleDbCommand(
				"SELECT FltlegID, Callsign, Origin, Dest, FLDt, Crew, Source, STA FROM FlightSchedule", wconn).ExecuteReader();
			while (dedupRdr.Read())
			{
				int id           = Convert.ToInt32(dedupRdr.GetValue(0));
				if (toDrop.Contains(id)) continue;   // already being dropped by an earlier pass
				string callsign  = dedupRdr.IsDBNull(1) ? "" : dedupRdr.GetString(1);
				string origin    = dedupRdr.IsDBNull(2) ? "" : dedupRdr.GetString(2);
				string dest      = dedupRdr.IsDBNull(3) ? "" : dedupRdr.GetString(3);
				string fldt      = dedupRdr.IsDBNull(4) ? "" : dedupRdr.GetString(4);
				string crew      = dedupRdr.IsDBNull(5) ? "" : dedupRdr.GetString(5);
				string source    = dedupRdr.IsDBNull(6) ? "" : dedupRdr.GetString(6);
				string sta       = dedupRdr.IsDBNull(7) ? "" : dedupRdr.GetString(7);

				string key = MergeKey(callsign, origin, dest, fldt);
				int sourceScore = source == "WS" ? 2 : (source == "MM" ? 1 : 0);
				int score = (crew != "" ? 100 : 0) + sourceScore * 10 + (sta != "" ? 1 : 0);

				int[] best;
				if (!bestBySameLeg.TryGetValue(key, out best))
				{
					bestBySameLeg[key] = new int[] { id, score };
				}
				else if (score > best[1])
				{
					toDrop.Add(best[0]);
					bestBySameLeg[key] = new int[] { id, score };
				}
				else
				{
					toDrop.Add(id);
				}
			}
			dedupRdr.Close();

			foreach (int id in toDrop)
			{
				OleDbCommand del = new OleDbCommand("DELETE FROM FlightSchedule WHERE FltlegID=?", wconn);
				del.Parameters.AddWithValue("?", id);
				del.ExecuteNonQuery();
			}

			wconn.Close();
		}

		// Parses the most recently modified *.csv in FlightSchedCsvFolder (the FPM
		// "Refuel EU" export the dispatcher drops there manually) and returns flights not
		// already covered by the webservice feed (wsKeys), as synthetic FlightSchedule
		// rows with Source="CSV". The CSV's own numeric "ID" column is negated to build a
		// stable FltlegID that can never collide with a real (always positive) Fltleg_ID,
		// so re-running with the same CSV updates the same row in place, and the row can
		// later be matched/dropped once the flight appears for real on the webservice.
		private List<object[]> LoadCsvSupplementalFlights(HashSet<string> wsKeys, HashSet<string> mmKeys, HashSet<int> csvIdsSeen)
		{
			List<object[]> result = new List<object[]>();
			string path = FindLatestFlightSchedCsv();
			if (path == null) return result;

			List<Dictionary<string, string>> rows;
			try { rows = ParseCsvFile(path); } catch { return result; }

			foreach (Dictionary<string, string> row in rows)
			{
				string id, atcRaw, flightRaw, aircraft, stdRaw, staRaw, dep, arr;
				if (!row.TryGetValue("ID", out id)) continue;
				row.TryGetValue("ATC", out atcRaw);
				row.TryGetValue("Flight", out flightRaw);
				row.TryGetValue("Aircraft", out aircraft);
				row.TryGetValue("STD", out stdRaw);
				row.TryGetValue("STA", out staRaw);
				row.TryGetValue("DEP", out dep);
				row.TryGetValue("ARR", out arr);

				int csvId;
				if (!int.TryParse((id ?? "").Trim(), out csvId) || csvId == 0) continue;

				// "ATC" is the tactical/rotation callsign (e.g. "TAY7LL") and can be shared
				// by several different legs of the same rotation — using it as-is created
				// duplicate-looking rows. The operator prefix still comes from ATC (it's
				// reliably the operating carrier's 3-letter code), but the numeric part
				// comes from "Flight" (e.g. "3V4267" -> "4267"), matching the webservice's
				// own FPfx+FLNr callsign convention (e.g. "TAY4267").
				string callsign = BuildCsvCallsign(atcRaw, flightRaw);
				if (!CallsignAllowed(callsign)) continue;   // e.g. only TAY/FDX/DHL, set on the Admin tab

				string reg = (aircraft ?? "").Replace("-", "").Trim().ToUpper();
				string origin = NormalizeIata(dep);
				string dest = NormalizeIata(arr);

				DateTime stdDt, staDt;
				string std = TryParseCsvDate(stdRaw, out stdDt) ? stdDt.ToString("yyyy-MM-ddTHH:mm:ss") + "Z" : "";
				string sta = TryParseCsvDate(staRaw, out staDt) ? staDt.ToString("yyyy-MM-ddTHH:mm:ss") + "Z" : "";
				if (callsign == "" || std == "") continue;   // not enough to place it in the grid
				if (origin == "" && dest == "") continue;   // same guard as the webservice path — no route, not a real flight

				string fldt = stdDt.ToString("yyyy-MM-dd");
				string mergeKey = MergeKey(callsign, origin, dest, fldt);
				int fltlegId = -csvId;
				csvIdsSeen.Add(fltlegId);
				if (wsKeys.Contains(mergeKey) || mmKeys.Contains(mergeKey)) continue;   // already covered by a higher-priority source

				result.Add(new object[] { fltlegId, fldt, callsign, reg, std, "", "CSV", sta, origin, dest });
			}
			return result;
		}

		// Rebuilds a webservice-style callsign (operator prefix + flight number, e.g.
		// "TAY4267") from the CSV's two separate fields: "ATC" reliably carries the 3-letter
		// operator code as its leading letters, but the digits that follow it are the
		// tactical rotation callsign, not the flight number — the flight number lives in
		// "Flight" instead (e.g. "3V4267", prefixed by a marketing/codeshare code that
		// itself starts with a digit — "3V"). The flight number is the LAST run of digits
		// in "Flight", not the first: taking the first run on "3V4267" grabs the stray "3"
		// from the codeshare prefix instead of "4267" (this was the TAY3 bug — regularly
		// inserting a bogus "TAY3" flight). Falls back to the raw, trimmed ATC value if
		// either field doesn't parse as expected.
		private static string BuildCsvCallsign(string atcRaw, string flightRaw)
		{
			string atc = (atcRaw ?? "").Trim().ToUpper();
			string prefix = Regex.Match(atc, @"^[A-Z]+").Value;
			if (prefix.Length > 3) prefix = prefix.Substring(0, 3);

			MatchCollection digitRuns = Regex.Matches((flightRaw ?? "").Trim(), @"\d+");
			string digits = digitRuns.Count > 0 ? digitRuns[digitRuns.Count - 1].Value : "";

			return (prefix != "" && digits != "") ? prefix + digits : atc;
		}

		// The CSV's Date/STD/STA columns are "dd/MM/yyyy" and "dd/MM/yyyy HH:mm" — assumed
		// UTC to match the webservice feed's convention (the export doesn't say either way;
		// worth confirming against a known flight if the merged times look off by a few hours).
		private static bool TryParseCsvDate(string raw, out DateTime result)
		{
			return DateTime.TryParseExact((raw ?? "").Trim(), "dd/MM/yyyy HH:mm",
				CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out result);
		}

		private static string FindLatestFlightSchedCsv()
		{
			try
			{
				if (!Directory.Exists(FlightSchedCsvFolder)) return null;
				string[] files = Directory.GetFiles(FlightSchedCsvFolder, "*.csv");
				if (files.Length == 0) return null;
				string latest = files[0];
				DateTime latestTime = File.GetLastWriteTimeUtc(latest);
				for (int i = 1; i < files.Length; i++)
				{
					DateTime t = File.GetLastWriteTimeUtc(files[i]);
					if (t > latestTime) { latestTime = t; latest = files[i]; }
				}
				return latest;
			}
			catch { return null; }
		}

		// Quote-aware CSV parsing (headers can contain no commas here, but values like
		// "TAY812 " and dates do), keyed by header name rather than fixed column index
		// since the FPM export's column order/count isn't a contract Dispatch Watch
		// controls. Reuses the same CsvSplit convention as MainForm.Rwys.cs/AirportList.cs.
		private static List<Dictionary<string, string>> ParseCsvFile(string path)
		{
			List<Dictionary<string, string>> rows = new List<Dictionary<string, string>>();
			using (StreamReader sr = new StreamReader(path))
			{
				string headerLine = sr.ReadLine();
				if (headerLine == null) return rows;
				string[] headers = CsvSplit(headerLine);

				string line;
				while ((line = sr.ReadLine()) != null)
				{
					if (line.Trim() == "") continue;
					string[] fields = CsvSplit(line);
					Dictionary<string, string> row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
					for (int i = 0; i < headers.Length && i < fields.Length; i++)
						row[headers[i].Trim()] = fields[i];
					rows.Add(row);
				}
			}
			return rows;
		}

		private struct BriefingDetails
		{
			public string Sta;
			public string Alt1, Alt2;
			public int FlightTimeMin, Alt1TimeMin, Alt2TimeMin;
		}

		// A getBriefing response bundles several BriefingContentProduct blocks (OFP, MET,
		// NOTAM...); the fields this reads (ScheduledTimeOfArrival, FlightTime, and the
		// alternates) all live under the OFP's ARINC-633 FlightPlan payload, which uses a
		// SEPARATE namespace (http://aeec.aviation-ia.net/633) from the outer envelope — so
		// lookups here match on local name only (.Descendants().Where(e => e.Name.LocalName
		// == ...)), never a fully-qualified XName, same idiom already proven in the sibling
		// WebService-PDF Archives project's briefing reader. Each destination alternate is one
		// <AlternateFuel> element: a direct-child <Airport airportFunction="..."> names it
		// (PrimaryArrivalAlternateAirport -> Alt1, the next distinct ArrivalAlternateAirport ->
		// Alt2 — DepartureAlternateAirport/ETOPSAdequateAirport are the take-off alternate and
		// en-route adequate airports, not relevant here), and the AlternateFuel's own
		// direct-child <Duration> is the flight time from destination to that alternate —
		// deliberately NOT the nested <FinalReserve><Duration> sibling, which is a fixed
		// reserve-fuel duration, not a flight time. A flat regex (still fine for the single,
		// non-repeated ScheduledTimeOfArrival) can't safely tell those two <Duration> elements
		// apart, hence XDocument here instead.
		private BriefingDetails FetchBriefingDetails(int fltlegId)
		{
			BriefingDetails d = new BriefingDetails();
			try
			{
				string xml;
				using (WebClient wc = new WebClient())
					xml = wc.DownloadString(_briefingBaseUrl.TrimEnd('/') + "/?METHOD=getBriefing&FLTLEG_ID=" + fltlegId);

				Match staMatch = Regex.Match(xml, "<ScheduledTimeOfArrival>([^<]*)</ScheduledTimeOfArrival>");
				d.Sta = staMatch.Success ? staMatch.Groups[1].Value : "";

				XDocument doc = XDocument.Parse(xml);

				XElement flightTimeEl = doc.Descendants().Where(e => e.Name.LocalName == "FlightTime").FirstOrDefault();
				XElement flightTimeVal = flightTimeEl != null
					? flightTimeEl.Descendants().Where(e => e.Name.LocalName == "Value").FirstOrDefault()
					: null;
				d.FlightTimeMin = flightTimeVal != null ? ParseIsoDurationMinutes(flightTimeVal.Value) : 0;

				List<string> seenAlt = new List<string>();
				foreach (XElement alt in doc.Descendants().Where(e => e.Name.LocalName == "AlternateFuel"))
				{
					XElement airport = alt.Elements().FirstOrDefault(e => e.Name.LocalName == "Airport");
					if (airport == null) continue;
					string function = (string)airport.Attribute("airportFunction") ?? "";
					if (function != "PrimaryArrivalAlternateAirport" && function != "ArrivalAlternateAirport") continue;

					XElement icaoEl = airport.Descendants().FirstOrDefault(e => e.Name.LocalName == "AirportICAOCode");
					string icao = icaoEl != null ? icaoEl.Value : "";
					if (icao == "" || seenAlt.Contains(icao)) continue;
					seenAlt.Add(icao);

					XElement durVal = alt.Elements().Where(e => e.Name.LocalName == "Duration")
						.SelectMany(e => e.Descendants().Where(v => v.Name.LocalName == "Value")).FirstOrDefault();
					int minutes = durVal != null ? ParseIsoDurationMinutes(durVal.Value) : 0;

					if (d.Alt1 == null) { d.Alt1 = icao; d.Alt1TimeMin = minutes; }
					else if (d.Alt2 == null) { d.Alt2 = icao; d.Alt2TimeMin = minutes; }
				}
				d.Alt1 = d.Alt1 ?? ""; d.Alt2 = d.Alt2 ?? "";
			}
			catch { /* briefing legitimately unavailable yet (e.g. far-future leg) — return whatever was parsed so far */ }
			return d;
		}

		// "PT4H16M00.000S" -> 256. Tolerant of a missing hours or minutes component.
		private static int ParseIsoDurationMinutes(string iso)
		{
			if (string.IsNullOrEmpty(iso)) return 0;
			Match m = Regex.Match(iso, @"PT(?:(\d+)H)?(?:(\d+)M)?");
			if (!m.Success) return 0;
			int hours = m.Groups[1].Success ? int.Parse(m.Groups[1].Value) : 0;
			int minutes = m.Groups[2].Success ? int.Parse(m.Groups[2].Value) : 0;
			return hours * 60 + minutes;
		}

		// 0 (no briefing yet / no such alternate) displays blank rather than "0:00".
		private static string FormatMinutesHhmm(int minutes)
		{
			if (minutes <= 0) return "";
			return (minutes / 60) + ":" + (minutes % 60).ToString("00");
		}

		private void LoadFlightScheduleGrid()
		{
			if (_fsDgv == null) return;
			_fsDgv.Rows.Clear();

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbDataReader reader = new OleDbCommand(
				"SELECT FltlegID, FLDt, Callsign, Reg, STD, STA, Crew, Origin, Dest, FlightTimeMin, Alt1, Alt1TimeMin, Alt2, Alt2TimeMin FROM FlightSchedule ORDER BY STD", conn).ExecuteReader();
			Dictionary<int, string> manualConflicts = LoadManualConflictsByFltlegId();
			// Every flight with an active Conflict-tab match (automatic or manually-forced) —
			// same matching logic as the Conflict tab itself (MainForm.Conflict.cs), so a
			// dismissed/false-positive match doesn't highlight a row here either.
			HashSet<int> conflictedIds = GetConflictedFltlegIds();
			while (reader.Read())
			{
				int    fltlegId = reader.IsDBNull(0) ? 0 : Convert.ToInt32(reader.GetValue(0));
				string fldt     = reader.IsDBNull(1) ? "" : reader.GetString(1);
				string callsign = reader.IsDBNull(2) ? "" : reader.GetString(2);
				string reg      = reader.IsDBNull(3) ? "" : reader.GetString(3);
				string std      = reader.IsDBNull(4) ? "" : reader.GetString(4);
				string sta      = reader.IsDBNull(5) ? "" : reader.GetString(5);
				string crew     = reader.IsDBNull(6) ? "" : reader.GetString(6);
				string origin   = reader.IsDBNull(7) ? "" : reader.GetString(7);
				string dest     = reader.IsDBNull(8) ? "" : reader.GetString(8);
				string flightTime = FormatMinutesHhmm(reader.IsDBNull(9) ? 0 : Convert.ToInt32(reader.GetValue(9)));
				string alt1      = reader.IsDBNull(10) ? "" : reader.GetString(10);
				string alt1Time  = FormatMinutesHhmm(reader.IsDBNull(11) ? 0 : Convert.ToInt32(reader.GetValue(11)));
				string alt2      = reader.IsDBNull(12) ? "" : reader.GetString(12);
				string alt2Time  = FormatMinutesHhmm(reader.IsDBNull(13) ? 0 : Convert.ToInt32(reader.GetValue(13)));

				string assignedKey;
				string forceBtnText = manualConflicts.TryGetValue(fltlegId, out assignedKey) && assignedKey != "" ? assignedKey : "+ Force";

				int rowIndex = _fsDgv.Rows.Add(fldt, callsign, origin, dest, reg, std, sta, flightTime, alt1, alt1Time, alt2, alt2Time, crew, fltlegId, forceBtnText);
				if (conflictedIds.Contains(fltlegId))
				{
					DataGridViewRow row = _fsDgv.Rows[rowIndex];
					row.DefaultCellStyle.BackColor = Color.FromArgb(255, 205, 210);
					row.DefaultCellStyle.SelectionBackColor = Color.FromArgb(239, 154, 154);
				}
			}
			conn.Close();

			// Lets the dispatcher visually confirm a fresh CSV drop into FlightSched/ was
			// actually picked up, since LoadCsvSupplementalFlights only ever reads the
			// single most-recently-modified file there silently.
			if (_fsCsvInfoLabel != null)
			{
				string csvPath = FindLatestFlightSchedCsv();
				_fsCsvInfoLabel.Text = csvPath != null
					? "Active CSV: " + Path.GetFileName(csvPath) + " (modified " + File.GetLastWriteTime(csvPath).ToString("dd/MM HH:mm") + ")"
					: "Active CSV: none";
			}

			FilterFsGrid();   // rows just got repopulated (and Visible reset) — reapply any active filters

			// Re-populating always leaves the grid scrolled to wherever it happened to be
			// before (or fully scrolled if the previous load had more rows), which can read
			// as the header row/first rows being clipped — pin the view back to the top.
			if (_fsDgv.Rows.Count > 0) _fsDgv.FirstDisplayedScrollingRowIndex = 0;
		}

		private void FsDgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
		{
			if (e.RowIndex < 0) return;
			if (_fsDgv.Columns[e.ColumnIndex].Name != "ForceConflict") return;
			int fltlegId = Convert.ToInt32(_fsDgv.Rows[e.RowIndex].Cells["FltlegID"].Value);
			ShowForceConflictDialog(fltlegId);
			LoadFlightScheduleGrid();   // refresh button text in case the assignment changed
		}

		// Modal dialog letting the dispatcher manually assign (or clear) a Kept, impact-
		// classified NOTAM as a forced conflict for one flight — for cases the automatic
		// Origin/Dest/Alt1/Alt2 matching on the Conflict tab (MainForm.Conflict.cs) can't
		// express. Same dark-theme, hand-positioned Form pattern as ShowLockPromptDialog
		// (MainForm.Deployment.cs).
		private void ShowForceConflictDialog(int fltlegId)
		{
			Dictionary<int, string> current = LoadManualConflictsByFltlegId();
			string currentKey;
			current.TryGetValue(fltlegId, out currentKey);

			using (Form dlg = new Form
			{
				Text = "Dispatch Watch", FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterScreen,
				ControlBox = false, MinimizeBox = false, MaximizeBox = false, Width = 480, Height = 420, BackColor = Color.FromArgb(38, 50, 56)
			})
			{
				Label current_lbl = new Label
				{
					Text = "Current: " + (!string.IsNullOrEmpty(currentKey) ? currentKey : "(none)"),
					ForeColor = Color.White, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold), Top = 14, Left = 20, AutoSize = true
				};
				dlg.Controls.Add(current_lbl);

				Label icaoLbl = new Label { Text = "ICAO", ForeColor = Color.FromArgb(207, 216, 220), Top = 46, Left = 20, AutoSize = true };
				dlg.Controls.Add(icaoLbl);
				TextBox icaoBox = new TextBox { Top = 64, Left = 20, Width = 100, BackColor = Color.White, ForeColor = Color.Black };
				dlg.Controls.Add(icaoBox);
				Button search = new Button { Text = "Search", Top = 62, Left = 128, Width = 80, Height = 24,
					BackColor = Color.FromArgb(69, 90, 100), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
				search.FlatAppearance.BorderColor = Color.FromArgb(96, 125, 139);
				dlg.Controls.Add(search);

				ListBox list = new ListBox { Top = 100, Left = 20, Width = 430, Height = 220, Font = new Font("Consolas", 8.5f),
					BackColor = Color.White, ForeColor = Color.Black };
				dlg.Controls.Add(list);

				search.Click += delegate
				{
					list.Items.Clear();
					foreach (string line in FindKeptNotamsForIcao(icaoBox.Text.Trim().ToUpper()))
						list.Items.Add(line);
				};

				Button assign = new Button { Text = "Assign", Top = 334, Left = 20, Width = 100, Height = 30,
					BackColor = Color.FromArgb(46, 125, 82), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
				assign.FlatAppearance.BorderColor = Color.FromArgb(96, 125, 139);
				Button clear = new Button { Text = "Clear assignment", Top = 334, Left = 130, Width = 140, Height = 30,
					BackColor = Color.FromArgb(69, 90, 100), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
				clear.FlatAppearance.BorderColor = Color.FromArgb(96, 125, 139);
				Button cancel = new Button { Text = "Cancel", Top = 334, Left = 280, Width = 100, Height = 30,
					BackColor = Color.FromArgb(69, 90, 100), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
				cancel.FlatAppearance.BorderColor = Color.FromArgb(96, 125, 139);

				assign.Click += delegate
				{
					if (list.SelectedItem == null) { MessageBox.Show("Select a NOTAM from the list first."); return; }
					string line = list.SelectedItem.ToString();
					string selectedKey = line.Split(new[] { "  " }, StringSplitOptions.None)[0].Trim();
					if (!EnsureWriterOrWarn()) return;
					AssignManualConflict(fltlegId, selectedKey);
					dlg.Close();
				};
				clear.Click += delegate
				{
					if (!EnsureWriterOrWarn()) return;
					DeleteManualConflict(fltlegId);
					dlg.Close();
				};
				cancel.Click += delegate { dlg.Close(); };

				dlg.Controls.Add(assign);
				dlg.Controls.Add(clear);
				dlg.Controls.Add(cancel);

				dlg.ShowDialog();
			}
		}

		// Kept, impact-classified NOTAMs at one station, formatted "KEY  Impact Label  snippet"
		// for ShowForceConflictDialog's picker list — ImpactLabel (MainForm.NotamFilter.cs) is
		// the same full-word label used everywhere else in the app (e.g. "No ILS" rather than the
		// raw "N" code), so the list reads consistently with the Conflict tab's own section
		// headers. "KEY" is split off by the two-space separator when the dispatcher assigns.
		private List<string> FindKeptNotamsForIcao(string icao)
		{
			List<string> result = new List<string>();
			if (icao == "") return result;
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand cmd = new OleDbCommand("SELECT [key], Impact, Remark, [all] FROM filteredNotams_table WHERE location=? AND Status='K' AND Impact<>''", conn);
			cmd.Parameters.AddWithValue("?", icao);
			OleDbDataReader reader = cmd.ExecuteReader();
			while (reader.Read())
			{
				string key    = reader.IsDBNull(0) ? "" : reader.GetString(0);
				string impact = reader.IsDBNull(1) ? "" : reader.GetString(1);
				string remark = reader.IsDBNull(2) ? "" : reader.GetString(2);
				string all    = reader.IsDBNull(3) ? "" : reader.GetString(3);
				if (key == "") continue;
				string snippet = remark != "" ? remark : all;
				snippet = snippet.Replace("\r", " ").Replace("\n", " ");
				if (snippet.Length > 70) snippet = snippet.Substring(0, 70) + "...";
				result.Add(key + "  " + ImpactLabel(impact) + "  " + snippet);
			}
			conn.Close();
			return result;
		}

		// Standalone HTML rendering of the next 7 days of FlightSchedule, for the "Export
		// PDF" button and the Send Reports email attachment — reuses the same query as
		// LoadFlightScheduleGrid, just clipped to 7 days and grouped by calendar date with a
		// colored header row per day (same visual language as MainForm.Reports.cs's Section()
		// headers), instead of the tab's flat 14-day grid.
		public string BuildFlightScheduleHtml()
		{
			DateTime nowUtc = DateTime.UtcNow;
			DateTime windowEnd = nowUtc.AddDays(7);

			// Same conflict set the live grid highlights with (MainForm.Conflict.cs) — a flight
			// with an active automatic or manually-forced match is flagged the same way here so
			// the exported PDF/email version carries the same information as the tab.
			HashSet<int> conflictedIds = GetConflictedFltlegIds();

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbDataReader reader = new OleDbCommand("SELECT FltlegID, FLDt, Callsign, Reg, STD, STA, Crew, Origin, Dest FROM FlightSchedule ORDER BY STD", conn).ExecuteReader();

			// Grouped by calendar date (UTC) in the order flights are encountered — the query
			// is already ORDER BY STD, so dates come out in ascending order naturally.
			List<string> dateOrder = new List<string>();
			Dictionary<string, StringBuilder> rowsByDate = new Dictionary<string, StringBuilder>();

			while (reader.Read())
			{
				int    fltlegId = reader.IsDBNull(0) ? 0 : Convert.ToInt32(reader.GetValue(0));
				string callsign = reader.IsDBNull(2) ? "" : reader.GetString(2);
				string reg      = reader.IsDBNull(3) ? "" : reader.GetString(3);
				string stdRaw   = reader.IsDBNull(4) ? "" : reader.GetString(4);
				string staRaw   = reader.IsDBNull(5) ? "" : reader.GetString(5);
				string crew     = reader.IsDBNull(6) ? "" : reader.GetString(6);
				string origin   = reader.IsDBNull(7) ? "" : reader.GetString(7);
				string dest     = reader.IsDBNull(8) ? "" : reader.GetString(8);

				DateTime std;
				if (!DateTime.TryParse(stdRaw, CultureInfo.InvariantCulture,
					DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out std)) continue;
				if (std < nowUtc || std > windowEnd) continue;

				DateTime sta;
				bool hasSta = DateTime.TryParse(staRaw, CultureInfo.InvariantCulture,
					DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out sta);

				string dateKey = std.ToString("yyyy-MM-dd");
				StringBuilder rows;
				if (!rowsByDate.TryGetValue(dateKey, out rows))
				{
					rows = new StringBuilder();
					rowsByDate[dateKey] = rows;
					dateOrder.Add(dateKey);
				}

				string trStyle = conflictedIds.Contains(fltlegId) ? " style=\"background:#ffcdd2\"" : "";
				rows.Append("<tr" + trStyle + ">" +
					"<td>" + callsign.Replace("&", "&amp;") + "</td>" +
					"<td>" + origin + "</td>" +
					"<td>" + dest + "</td>" +
					"<td>" + reg + "</td>" +
					"<td>" + std.ToString("HH:mm") + "Z</td>" +
					"<td>" + (hasSta ? sta.ToString("HH:mm") + "Z" : "") + "</td>" +
					"<td>" + crew.Replace("&", "&amp;") + "</td>" +
					"</tr>");
			}
			conn.Close();

			StringBuilder table = new StringBuilder();
			foreach (string dateKey in dateOrder)
			{
				DateTime d = DateTime.ParseExact(dateKey, "yyyy-MM-dd", CultureInfo.InvariantCulture);
				table.Append("<tr><th colspan=\"7\" class=\"dateHdr\">" + d.ToString("dddd dd MMMM yyyy", CultureInfo.InvariantCulture) + "</th></tr>");
				table.Append("<tr class=\"colHdr\"><th>Callsign</th><th>Origin</th><th>Dest</th><th>Reg</th><th>STD (UTC)</th><th>STA (UTC)</th><th>Crew</th></tr>");
				table.Append(rowsByDate[dateKey]);
			}

			string reportDate = DateTime.Now.ToString("ddMMMMyyyy HHmm") + "CET";
			return
				"<html><head><title>FLIGHT SCHEDULE</title><style>" +
				"body{font-family:Calibri}" +
				"table{width:700px;text-align:left;font-size:12px;border:1px solid black;border-collapse:collapse}" +
				"th,td{border:1px solid black;padding:3px 6px}" +
				".dateHdr{background:SlateGray;color:white;font-size:14px;font-weight:bold;text-align:left}" +
				".colHdr{background:LightSlateGrey;color:white;font-weight:bold}" +
				"</style></head><body>" +
				"<h1>Flight Schedule - Next 7 Days</h1><p>" + reportDate + "</p>" +
				"<table>" + table + "</table>" +
				"</body></html>";
		}

		// AND-combines every non-empty per-column filter textbox against that column's
		// cell text (substring, case-insensitive) — a row must match ALL active filters
		// to stay visible.
		private void FilterFsGrid()
		{
			if (_fsDgv == null || _fsFilterBoxes == null) return;
			foreach (DataGridViewRow row in _fsDgv.Rows)
			{
				if (row.IsNewRow) continue;
				bool visible = true;
				for (int i = 0; i < _fsFilterBoxes.Length; i++)
				{
					string filter = _fsFilterBoxes[i].Text.Trim();
					if (filter == "") continue;
					object val = row.Cells[i].Value;
					string cellText = val != null ? val.ToString() : "";
					if (cellText.IndexOf(filter, StringComparison.OrdinalIgnoreCase) < 0) { visible = false; break; }
				}
				row.Visible = visible;
			}
		}

		// ── progress dialog (dark, matches the DB Update dialog's styling) ────
		private void ShowFsProgressForm()
		{
			_fsProgressForm = new Form
			{
				StartPosition = FormStartPosition.CenterScreen,
				FormBorderStyle = FormBorderStyle.FixedDialog,
				ControlBox = false, MinimizeBox = false, MaximizeBox = false,
				TopMost = true, ShowInTaskbar = false,
				Width = 420, Height = 130,
				BackColor = Color.FromArgb(38, 50, 56),
				Text = "Dispatch Watch"
			};
			Label title = new Label { Text = "LOADING FLIGHT SCHEDULE", ForeColor = Color.White,
				Font = new Font("Segoe UI", 10, FontStyle.Bold), Dock = DockStyle.Top, Height = 32,
				TextAlign = ContentAlignment.MiddleCenter };
			_fsProgressStatus = new Label { Text = "Starting...", ForeColor = Color.FromArgb(207, 216, 220),
				Font = new Font("Segoe UI", 9), Dock = DockStyle.Top, Height = 26, TextAlign = ContentAlignment.MiddleCenter };
			_fsProgressBar = new ProgressBar { Minimum = 0, Maximum = 100, Value = 0, Style = ProgressBarStyle.Continuous, Dock = DockStyle.Bottom, Height = 20 };
			Panel pad = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(38, 50, 56) };
			_fsProgressForm.Controls.Add(pad);
			_fsProgressForm.Controls.Add(_fsProgressBar);
			_fsProgressForm.Controls.Add(_fsProgressStatus);
			_fsProgressForm.Controls.Add(title);
			_fsProgressForm.Show(this);
		}

		private void UpdateFsProgress(int percent, string status)
		{
			if (_fsProgressForm == null || _fsProgressForm.IsDisposed) return;
			_fsProgressBar.Value = Math.Max(0, Math.Min(100, percent));
			if (status != null) _fsProgressStatus.Text = status;
		}

		private void CloseFsProgressForm()
		{
			if (_fsProgressForm != null && !_fsProgressForm.IsDisposed) _fsProgressForm.Close();
			_fsProgressForm = null;
		}
	}
}
