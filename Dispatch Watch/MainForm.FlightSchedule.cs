using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		private DataGridView _fsDgv;
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
				conn.Close();
			}
			catch { }
		}

		// Where the dispatcher drops the FPM "Refuel EU" CSV export as a stand-in for
		// flights the FlightScheduleService.svc feed doesn't expose yet. Created next to
		// the exe so it's obvious where to put the file; the exact filename doesn't
		// matter, the most recently modified *.csv in the folder is used.
		private static string FlightSchedCsvFolder { get { return Path.Combine(Application.StartupPath, "FlightSched"); } }

		public void EnsureFlightSchedFolder()
		{
			try { Directory.CreateDirectory(FlightSchedCsvFolder); } catch { }
		}

		void FlightScheduleTabEnter(object sender, EventArgs e)
		{
			if (_fsDgv == null) BuildFlightScheduleGrid();
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
			Panel topBar = new Panel { Tag = "dispose", Dock = DockStyle.Top, Height = 50 };
			Label hdr = new Label { Top = 14, Left = 20, AutoSize = true,
				Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold), Text = "Flight Schedule — next 7 days (UTC)" };
			topBar.Controls.Add(hdr);
			Button refresh = new Button { Top = 10, Left = 400, Width = 90, Height = 28, Text = "Refresh" };
			refresh.Click += delegate { RefreshFlightSchedule(); };
			topBar.Controls.Add(refresh);
			tabPage_FlightSchedule.Controls.Add(topBar);

			_fsDgv = new DataGridView { Tag = "dispose", Dock = DockStyle.Fill,
				ReadOnly = true, AllowUserToAddRows = false, RowHeadersWidth = 28, BackgroundColor = Color.White,
				AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None };
			_fsDgv.ColumnHeadersDefaultCellStyle.Font = new Font(_fsDgv.Font, FontStyle.Bold);
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Date",     HeaderText = "Date",      Width = 90 });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Callsign", HeaderText = "Callsign",  Width = 100 });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Origin",   HeaderText = "Origin",    Width = 70 });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Dest",     HeaderText = "Dest",      Width = 70 });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Reg",      HeaderText = "Reg",       Width = 80 });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "STD",      HeaderText = "STD (UTC)", Width = 140 });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "STA",      HeaderText = "STA (UTC)", Width = 140 });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Crew",     HeaderText = "Crew",      Width = 380 });
			// Dock=Fill controls occupy remaining space in add order, so the grid must be
			// added after the Top-docked bar.
			tabPage_FlightSchedule.Controls.Add(_fsDgv);
		}

		// Downloads getFlightList for today..today+7, resolves STA via a cached getBriefing
		// call per flight (skipped when the flight's STD hasn't moved since the last
		// fetch), stores everything in FlightSchedule and reloads the grid sorted by STD.
		// Runs on a BackgroundWorker with a progress dialog — a 7-day window can mean a few
		// hundred individual getBriefing calls.
		public void RefreshFlightSchedule()
		{
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
				CloseFsProgressForm();
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

		private void FetchFlightSchedule(Action<int, string> onProgress)
		{
			Dictionary<int, string[]> cached = new Dictionary<int, string[]>();   // FltlegID -> [STD, STA]
			OleDbConnection rconn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			rconn.Open();
			OleDbDataReader rdr = new OleDbCommand("SELECT FltlegID, STD, STA FROM FlightSchedule", rconn).ExecuteReader();
			while (rdr.Read())
			{
				if (rdr.IsDBNull(0)) continue;
				int id = Convert.ToInt32(rdr.GetValue(0));
				string std = rdr.IsDBNull(1) ? "" : rdr.GetString(1);
				string sta = rdr.IsDBNull(2) ? "" : rdr.GetString(2);
				cached[id] = new string[] { std, sta };
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

				string[] cache;
				bool exists = cached.TryGetValue(fltlegId, out cache);
				string sta;
				if (source == "CSV")
					sta = csvSta;   // no Fltleg_ID to brief yet — STA comes straight from the CSV
				else
					// getBriefing legitimately 404s/errors for flights whose briefing hasn't
					// been generated yet (e.g. far-future legs) — FetchSta already swallows
					// that and returns "", and the flight is still persisted with a blank STA
					// rather than being dropped from the list.
					sta = (exists && cache[0] == std && cache[1] != "") ? cache[1] : FetchSta(fltlegId);

				try
				{
				if (exists)
				{
					OleDbCommand upd = new OleDbCommand("UPDATE FlightSchedule SET FLDt=?, Callsign=?, Reg=?, STD=?, STA=?, Crew=?, Source=?, Origin=?, Dest=? WHERE FltlegID=?", wconn);
					upd.Parameters.AddWithValue("?", fldtVal);
					upd.Parameters.AddWithValue("?", callsign);
					upd.Parameters.AddWithValue("?", reg);
					upd.Parameters.AddWithValue("?", std);
					upd.Parameters.AddWithValue("?", sta);
					upd.Parameters.AddWithValue("?", crew);
					upd.Parameters.AddWithValue("?", source);
					upd.Parameters.AddWithValue("?", origin);
					upd.Parameters.AddWithValue("?", dest);
					upd.Parameters.AddWithValue("?", fltlegId);
					upd.ExecuteNonQuery();
				}
				else
				{
					OleDbCommand ins = new OleDbCommand("INSERT INTO FlightSchedule ([FltlegID],[FLDt],[Callsign],[Reg],[STD],[STA],[Crew],[Source],[Origin],[Dest]) VALUES (?,?,?,?,?,?,?,?,?,?)", wconn);
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
					ins.ExecuteNonQuery();
				}
				}
				catch { /* one bad row shouldn't drop the rest of the week */ }

				done++;
				onProgress(50 + done * 45 / Math.Max(1, total), "Loading briefing " + done + " / " + total + "...");
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

		// A getBriefing response bundles several BriefingContentProduct blocks (OFP, MET,
		// NOTAM...); ScheduledTimeOfArrival appears once, deep under the OFP's
		// FlightPlanSummary. A plain regex search (same style already used for the NOTAM
		// XML in MainForm.NotamData.cs) is simpler and more robust here than walking the
		// deeply nested, namespaced OFP tree with LINQ to XML.
		private string FetchSta(int fltlegId)
		{
			try
			{
				string xml;
				using (WebClient wc = new WebClient())
					xml = wc.DownloadString(_briefingBaseUrl.TrimEnd('/') + "/?METHOD=getBriefing&FLTLEG_ID=" + fltlegId);
				Match m = Regex.Match(xml, "<ScheduledTimeOfArrival>([^<]*)</ScheduledTimeOfArrival>");
				return m.Success ? m.Groups[1].Value : "";
			}
			catch { return ""; }
		}

		private void LoadFlightScheduleGrid()
		{
			if (_fsDgv == null) return;
			_fsDgv.Rows.Clear();

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbDataReader reader = new OleDbCommand("SELECT FLDt, Callsign, Reg, STD, STA, Crew, Origin, Dest FROM FlightSchedule ORDER BY STD", conn).ExecuteReader();
			while (reader.Read())
			{
				string fldt     = reader.IsDBNull(0) ? "" : reader.GetString(0);
				string callsign = reader.IsDBNull(1) ? "" : reader.GetString(1);
				string reg      = reader.IsDBNull(2) ? "" : reader.GetString(2);
				string std      = reader.IsDBNull(3) ? "" : reader.GetString(3);
				string sta      = reader.IsDBNull(4) ? "" : reader.GetString(4);
				string crew     = reader.IsDBNull(5) ? "" : reader.GetString(5);
				string origin   = reader.IsDBNull(6) ? "" : reader.GetString(6);
				string dest     = reader.IsDBNull(7) ? "" : reader.GetString(7);
				_fsDgv.Rows.Add(fldt, callsign, origin, dest, reg, std, sta, crew);
			}
			conn.Close();
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
