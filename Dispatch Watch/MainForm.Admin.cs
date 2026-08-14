using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		// Base URLs for the Flight Schedule feature's two web services. Editable from the
		// Admin tab and persisted to ArchiveConfig.xml next to the exe, mirroring the
		// external "WebService-Archives" app's own config file/Admin tab. Scoped to Flight
		// Schedule only — the NOTAM tab's getAdHocNOTAM call in MainForm.NotamData.cs stays
		// hardcoded, to avoid any regression risk on the existing NOTAM pipeline.
		private static string _flightScheduleBaseUrl = "http://10.48.12.43:5458/FlightScheduleService.svc/web/";
		private static string _briefingBaseUrl        = "http://10.48.12.43:5455/BriefingService.svc/web/";
		// Callsign prefixes (e.g. "TAY","FDX","DHL") the Flight Schedule tab keeps — every
		// other operator's flights are filtered out of both the webservice feed and the
		// CSV fallback. Also editable here.
		private static List<string> _callsignPrefixFilters = new List<string> { "TAY", "FDX", "DHL" };
		// ±window (hours) the Conflict tab uses around a flight's STD/STA to decide
		// whether an overlapping NOTAM counts as a conflict. Editable here.
		private static int _conflictWindowHours = 1;
		// ±window (hours) the Conflict tab uses around a flight's ESTIMATED ARRIVAL AT AN
		// ALTERNATE (STD + FlightTimeMin + AltNTimeMin) to decide whether a "Not ALTN" NOTAM
		// at that alternate is a real diversion-time conflict — kept independent of
		// _conflictWindowHours (origin/destination) since it's a different kind of estimate.
		private static int _altnConflictWindowHours = 2;
		// Root folder of the Movement Manager XML "flightInfo_V2" message feed (see
		// MainForm.MmSchedule.cs) — a network share in production. Editable here in case
		// the path changes.
		private static string _mmMessagesPath = @"\\PRODWFPMFLS.aslairlines.com\CaeFpmProd\Archive\INTERFACING\SCHEDULE\IN";
		// FlightSched CSV fallback folder (MainForm.FlightSchedule.cs) — must be a shared V:
		// location, not a per-dispatcher local folder: since the multi-user deployment
		// (MainForm.Deployment.cs) gives every dispatcher their own local install, a CSV
		// dropped into an Application.StartupPath-relative folder would only ever be visible
		// to that one dispatcher's own instance.
		private static string _flightSchedCsvPath =
			@"V:\TAY Ops Control Centre\Flight Dispatch\AIP SUP -  Notams report\NOTAMS APP\FlightSched";
		// TAF Analysis tab's weather webservice base URL (getAdHocMET/METTYPE=FT) — a separate
		// service/port from the NOTAM and Flight Schedule feeds, editable here for consistency
		// with the other endpoint settings even though the TAF tab has its own UI for thresholds.
		private static string _metBaseUrl = "http://10.48.12.43:65080/efb-webservice/";
		private static bool   _archiveConfigLoaded;

		// Bumped by hand on notable releases — no build-time version stamping exists in this
		// SharpDevelop project, so this is the one source of truth for "what version is this".
		private const string AppVersion = "Dispatch Watch 1.0";
		// User guide, placed by hand on V: next to the app (VAppFolder, MainForm.Deployment.cs) —
		// not auto-copied to the local install like the app files themselves, since it's
		// reference material a dispatcher opens directly from the shared drive, not something
		// each local install needs its own copy of.
		private static string UserGuidePath { get { return Path.Combine(VAppFolder, "Dispatch Watch - User Guide.pdf"); } }

		private TextBox _adminFlightScheduleUrl;
		private TextBox _adminBriefingUrl;
		private TextBox _adminCallsignPrefixes;
		private TextBox _adminConflictWindow;
		private TextBox _adminAltnConflictWindow;
		private DataGridView _adminUsersDgv;
		private TextBox _adminMmMessagesPath;
		private TextBox _adminFlightSchedCsvPath;
		private TextBox _adminMetUrl;
		private Label   _adminStatus;

		private static string ArchiveConfigPath { get { return Path.Combine(Application.StartupPath, "ArchiveConfig.xml"); } }

		public void EnsureArchiveConfig()
		{
			if (_archiveConfigLoaded) return;
			try
			{
				if (File.Exists(ArchiveConfigPath))
				{
					XDocument doc = XDocument.Load(ArchiveConfigPath);
					XElement root = doc.Root;
					XElement fs = root.Element("FlightScheduleServiceUrl");
					XElement br = root.Element("BriefingServiceUrl");
					XElement cs = root.Element("CallsignPrefixes");
					XElement cw = root.Element("ConflictWindowHours");
					XElement acw = root.Element("AltnConflictWindowHours");
					XElement mm = root.Element("MmMessagesPath");
					XElement fc = root.Element("FlightSchedCsvPath");
					XElement met = root.Element("MetServiceUrl");
					if (fs != null && !string.IsNullOrEmpty(fs.Value)) _flightScheduleBaseUrl = fs.Value;
					if (br != null && !string.IsNullOrEmpty(br.Value)) _briefingBaseUrl = br.Value;
					if (cs != null && !string.IsNullOrEmpty(cs.Value)) _callsignPrefixFilters = ParsePrefixes(cs.Value);
					if (mm != null && !string.IsNullOrEmpty(mm.Value)) _mmMessagesPath = mm.Value;
					if (fc != null && !string.IsNullOrEmpty(fc.Value)) _flightSchedCsvPath = fc.Value;
					if (met != null && !string.IsNullOrEmpty(met.Value)) _metBaseUrl = met.Value;
					int parsedWindow;
					if (cw != null && int.TryParse(cw.Value, out parsedWindow) && parsedWindow > 0) _conflictWindowHours = parsedWindow;
					int parsedAltnWindow;
					if (acw != null && int.TryParse(acw.Value, out parsedAltnWindow) && parsedAltnWindow > 0) _altnConflictWindowHours = parsedAltnWindow;

					// TAF Analysis tab's 8 threshold values — declared in MainForm.TafAnalysis.cs,
					// loaded/saved here alongside the other ArchiveConfig-backed settings.
					XElement visCatI = root.Element("TafVisCatIm");
					XElement visAdv  = root.Element("TafVisAdvisoryM");
					XElement ceilCatI = root.Element("TafCeilCatIHundredFt");
					XElement ceilAdv  = root.Element("TafCeilAdvisoryHundredFt");
					XElement windKtCatI = root.Element("TafWindCatIKt");
					XElement windKtAdv  = root.Element("TafWindAdvisoryKt");
					XElement windMpsCatI = root.Element("TafWindCatIMps");
					XElement windMpsAdv  = root.Element("TafWindAdvisoryMps");
					int parsedTaf;
					if (visCatI != null && int.TryParse(visCatI.Value, out parsedTaf) && parsedTaf > 0) _tafVisCatIm = parsedTaf;
					if (visAdv != null && int.TryParse(visAdv.Value, out parsedTaf) && parsedTaf > 0) _tafVisAdvisoryM = parsedTaf;
					if (ceilCatI != null && int.TryParse(ceilCatI.Value, out parsedTaf) && parsedTaf > 0) _tafCeilCatIHundredFt = parsedTaf;
					if (ceilAdv != null && int.TryParse(ceilAdv.Value, out parsedTaf) && parsedTaf > 0) _tafCeilAdvisoryHundredFt = parsedTaf;
					if (windKtCatI != null && int.TryParse(windKtCatI.Value, out parsedTaf) && parsedTaf > 0) _tafWindCatIKt = parsedTaf;
					if (windKtAdv != null && int.TryParse(windKtAdv.Value, out parsedTaf) && parsedTaf > 0) _tafWindAdvisoryKt = parsedTaf;
					if (windMpsCatI != null && int.TryParse(windMpsCatI.Value, out parsedTaf) && parsedTaf > 0) _tafWindCatIMps = parsedTaf;
					if (windMpsAdv != null && int.TryParse(windMpsAdv.Value, out parsedTaf) && parsedTaf > 0) _tafWindAdvisoryMps = parsedTaf;
				}
				else
				{
					SaveArchiveConfig();
				}
			}
			catch { /* keep defaults */ }
			_archiveConfigLoaded = true;
		}

		private void SaveArchiveConfig()
		{
			try
			{
				XDocument doc = new XDocument(
					new XElement("ArchiveConfig",
						new XElement("FlightScheduleServiceUrl", _flightScheduleBaseUrl),
						new XElement("BriefingServiceUrl", _briefingBaseUrl),
						new XElement("CallsignPrefixes", string.Join(",", _callsignPrefixFilters.ToArray())),
						new XElement("ConflictWindowHours", _conflictWindowHours),
						new XElement("AltnConflictWindowHours", _altnConflictWindowHours),
						new XElement("MmMessagesPath", _mmMessagesPath),
						new XElement("FlightSchedCsvPath", _flightSchedCsvPath),
						new XElement("MetServiceUrl", _metBaseUrl),
						new XElement("TafVisCatIm", _tafVisCatIm),
						new XElement("TafVisAdvisoryM", _tafVisAdvisoryM),
						new XElement("TafCeilCatIHundredFt", _tafCeilCatIHundredFt),
						new XElement("TafCeilAdvisoryHundredFt", _tafCeilAdvisoryHundredFt),
						new XElement("TafWindCatIKt", _tafWindCatIKt),
						new XElement("TafWindAdvisoryKt", _tafWindAdvisoryKt),
						new XElement("TafWindCatIMps", _tafWindCatIMps),
						new XElement("TafWindAdvisoryMps", _tafWindAdvisoryMps)));
				doc.Save(ArchiveConfigPath);
			}
			catch { }
		}

		private static List<string> ParsePrefixes(string csv)
		{
			List<string> result = new List<string>();
			foreach (string part in csv.Split(','))
			{
				string p = part.Trim().ToUpper();
				if (p != "") result.Add(p);
			}
			return result;
		}

		// Used by MainForm.FlightSchedule.cs to keep only the operators the dispatcher
		// tracks; an empty filter list is treated as "no filter" rather than "keep nothing".
		public static bool CallsignAllowed(string callsign)
		{
			if (_callsignPrefixFilters == null || _callsignPrefixFilters.Count == 0) return true;
			string upper = (callsign ?? "").Trim().ToUpper();
			foreach (string prefix in _callsignPrefixFilters)
				if (upper.StartsWith(prefix)) return true;
			return false;
		}

		void AdminTabEnter(object sender, EventArgs e)
		{
			if (_adminFlightScheduleUrl == null) BuildAdminTab();
		}

		private void BuildAdminTab()
		{
			EnsureArchiveConfig();
			ClearTaggedControls(tabPage_Admin);

			Label versionLbl = new Label { Tag = "dispose", Top = 14, Left = 20, AutoSize = true,
				Font = new Font("Microsoft Sans Serif", 11f, FontStyle.Bold), Text = AppVersion };
			tabPage_Admin.Controls.Add(versionLbl);

			LinkLabel guideLink = new LinkLabel { Tag = "dispose", Top = 38, Left = 20, AutoSize = true, Text = "Open User Guide" };
			guideLink.LinkClicked += delegate
			{
				try
				{
					if (!File.Exists(UserGuidePath))
					{
						MessageBox.Show("User guide not found:\n" + UserGuidePath, "Open User Guide", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}
					Process.Start(UserGuidePath);
				}
				catch (Exception ex)
				{
					MessageBox.Show("Could not open the user guide:\n" + ex.Message, "Open User Guide", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			};
			tabPage_Admin.Controls.Add(guideLink);

			Label hdr = new Label { Tag = "dispose", Top = 66, Left = 20, AutoSize = true,
				Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold), Text = "Admin — Archive service endpoints" };
			tabPage_Admin.Controls.Add(hdr);

			Label lbl1 = new Label { Tag = "dispose", Top = 108, Left = 20, AutoSize = true,
				Text = "Flight Schedule Service URL (getFlightList)" };
			tabPage_Admin.Controls.Add(lbl1);
			_adminFlightScheduleUrl = new TextBox { Tag = "dispose", Top = 130, Left = 20, Width = 520, Text = _flightScheduleBaseUrl };
			tabPage_Admin.Controls.Add(_adminFlightScheduleUrl);

			Label lbl2 = new Label { Tag = "dispose", Top = 166, Left = 20, AutoSize = true,
				Text = "Briefing Service URL (getBriefing)" };
			tabPage_Admin.Controls.Add(lbl2);
			_adminBriefingUrl = new TextBox { Tag = "dispose", Top = 188, Left = 20, Width = 520, Text = _briefingBaseUrl };
			tabPage_Admin.Controls.Add(_adminBriefingUrl);

			Label lbl3 = new Label { Tag = "dispose", Top = 224, Left = 20, AutoSize = true,
				Text = "Flight Schedule — callsign prefixes to keep (comma-separated)" };
			tabPage_Admin.Controls.Add(lbl3);
			_adminCallsignPrefixes = new TextBox { Tag = "dispose", Top = 246, Left = 20, Width = 520,
				Text = string.Join(",", _callsignPrefixFilters.ToArray()) };
			tabPage_Admin.Controls.Add(_adminCallsignPrefixes);

			Label lbl4 = new Label { Tag = "dispose", Top = 282, Left = 20, AutoSize = true,
				Text = "Conflict tab — window around STD/STA (± hours)" };
			tabPage_Admin.Controls.Add(lbl4);
			_adminConflictWindow = new TextBox { Tag = "dispose", Top = 304, Left = 20, Width = 80,
				Text = _conflictWindowHours.ToString() };
			tabPage_Admin.Controls.Add(_adminConflictWindow);

			Label lbl4b = new Label { Tag = "dispose", Top = 340, Left = 20, AutoSize = true,
				Text = "Conflict tab — Not ALTN window around estimated alternate arrival (± hours)" };
			tabPage_Admin.Controls.Add(lbl4b);
			_adminAltnConflictWindow = new TextBox { Tag = "dispose", Top = 362, Left = 20, Width = 80,
				Text = _altnConflictWindowHours.ToString() };
			tabPage_Admin.Controls.Add(_adminAltnConflictWindow);

			Label lbl5 = new Label { Tag = "dispose", Top = 398, Left = 20, AutoSize = true,
				Text = "Movement Manager XML messages folder" };
			tabPage_Admin.Controls.Add(lbl5);
			_adminMmMessagesPath = new TextBox { Tag = "dispose", Top = 420, Left = 20, Width = 520, Text = _mmMessagesPath };
			tabPage_Admin.Controls.Add(_adminMmMessagesPath);

			Label lbl6 = new Label { Tag = "dispose", Top = 456, Left = 20, AutoSize = true,
				Text = "Flight Sched CSV folder (shared — every dispatcher's fallback CSV drop)" };
			tabPage_Admin.Controls.Add(lbl6);
			_adminFlightSchedCsvPath = new TextBox { Tag = "dispose", Top = 478, Left = 20, Width = 520, Text = _flightSchedCsvPath };
			tabPage_Admin.Controls.Add(_adminFlightSchedCsvPath);

			Label lbl7 = new Label { Tag = "dispose", Top = 514, Left = 20, AutoSize = true,
				Text = "Weather Service URL (getAdHocMET, TAF Analysis tab)" };
			tabPage_Admin.Controls.Add(lbl7);
			_adminMetUrl = new TextBox { Tag = "dispose", Top = 536, Left = 20, Width = 520, Text = _metBaseUrl };
			tabPage_Admin.Controls.Add(_adminMetUrl);

			Button save = new Button { Tag = "dispose", Top = 572, Left = 20, Width = 100, Height = 30, Text = "Save" };
			save.Click += delegate
			{
				if (!EnsureWriterOrWarn()) return;
				_flightScheduleBaseUrl   = _adminFlightScheduleUrl.Text.Trim();
				_briefingBaseUrl         = _adminBriefingUrl.Text.Trim();
				_callsignPrefixFilters   = ParsePrefixes(_adminCallsignPrefixes.Text);
				_mmMessagesPath          = _adminMmMessagesPath.Text.Trim();
				_flightSchedCsvPath      = _adminFlightSchedCsvPath.Text.Trim();
				_metBaseUrl              = _adminMetUrl.Text.Trim();
				int parsedWindow;
				if (int.TryParse(_adminConflictWindow.Text.Trim(), out parsedWindow) && parsedWindow > 0)
					_conflictWindowHours = parsedWindow;
				int parsedAltnWindow;
				if (int.TryParse(_adminAltnConflictWindow.Text.Trim(), out parsedAltnWindow) && parsedAltnWindow > 0)
					_altnConflictWindowHours = parsedAltnWindow;
				SaveArchiveConfig();
				EnsureFlightSchedFolder();
				LogUserAction("Save Admin settings");
				_adminStatus.Text = "Saved.";
			};
			tabPage_Admin.Controls.Add(save);

			_adminStatus = new Label { Tag = "dispose", Top = 580, Left = 130, AutoSize = true, ForeColor = Color.Gray, Text = "" };
			tabPage_Admin.Controls.Add(_adminStatus);

			Label usersHdr = new Label { Tag = "dispose", Top = 622, Left = 20, AutoSize = true,
				Font = new Font("Microsoft Sans Serif", 11f, FontStyle.Bold), Text = "Connected users" };
			tabPage_Admin.Controls.Add(usersHdr);

			_adminUsersDgv = new DataGridView
			{
				Tag = "dispose", Top = 650, Left = 20, Width = 520, Height = 220,
				AllowUserToAddRows = false, AllowUserToDeleteRows = false, ReadOnly = true,
				AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
				SelectionMode = DataGridViewSelectionMode.FullRowSelect,
				RowHeadersVisible = false
			};
			_adminUsersDgv.Columns.Add("Email", "Email");
			_adminUsersDgv.Columns.Add("LastSeen", "Last connected");
			_adminUsersDgv.Columns.Add("FirstSeen", "First connected");
			tabPage_Admin.Controls.Add(_adminUsersDgv);
			LoadUsersGrid();
		}

		// Every dispatcher who has ever launched the app, from the shared Users table
		// (Email/FirstSeen/LastSeen, MainForm.Deployment.cs — updated on every launch via
		// UpsertCurrentUser). Sorted most-recently-connected first.
		private void LoadUsersGrid()
		{
			if (_adminUsersDgv == null) return;
			_adminUsersDgv.Rows.Clear();
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				OleDbDataReader r = new OleDbCommand("SELECT Email, LastSeen, FirstSeen FROM Users", conn).ExecuteReader();
				List<string[]> rows = new List<string[]>();
				while (r.Read())
				{
					string email = r.IsDBNull(0) ? "" : r.GetString(0);
					string lastSeen = r.IsDBNull(1) ? "" : r.GetString(1);
					string firstSeen = r.IsDBNull(2) ? "" : r.GetString(2);
					rows.Add(new string[] { email, FormatUserTimestamp(lastSeen), FormatUserTimestamp(firstSeen), lastSeen });
				}
				conn.Close();
				// Most-recently-connected first — sort on the raw ISO LastSeen (4th column,
				// not shown), since it sorts correctly as plain text while the display-formatted
				// "dd/MM/yyyy HH:mm" version doesn't.
				rows.Sort(delegate(string[] a, string[] b) { return string.CompareOrdinal(b[3], a[3]); });
				foreach (string[] row in rows)
					_adminUsersDgv.Rows.Add(row[0], row[1], row[2]);
			}
			catch { /* Users table not ready yet on a brand-new DB */ }
		}

		// Users.FirstSeen/LastSeen are stored via DateTime.Now.ToString("o") (round-trip ISO,
		// local time) — reformatted here to "dd/MM/yyyy HH:mm" (date AND time, not just time) for
		// a compact, readable column. Falls back to the raw stored value if it doesn't parse.
		private static string FormatUserTimestamp(string iso)
		{
			if (string.IsNullOrEmpty(iso)) return "";
			DateTime dt;
			if (DateTime.TryParse(iso, System.Globalization.CultureInfo.InvariantCulture,
				System.Globalization.DateTimeStyles.RoundtripKind, out dt))
				return dt.ToString("dd/MM/yyyy HH:mm");
			return iso;
		}
	}
}
