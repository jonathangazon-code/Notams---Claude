using System;
using System.Collections.Generic;
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
		private static int _conflictWindowHours = 12;
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
		private static bool   _archiveConfigLoaded;

		private TextBox _adminFlightScheduleUrl;
		private TextBox _adminBriefingUrl;
		private TextBox _adminCallsignPrefixes;
		private TextBox _adminConflictWindow;
		private TextBox _adminMmMessagesPath;
		private TextBox _adminFlightSchedCsvPath;
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
					XElement mm = root.Element("MmMessagesPath");
					XElement fc = root.Element("FlightSchedCsvPath");
					if (fs != null && !string.IsNullOrEmpty(fs.Value)) _flightScheduleBaseUrl = fs.Value;
					if (br != null && !string.IsNullOrEmpty(br.Value)) _briefingBaseUrl = br.Value;
					if (cs != null && !string.IsNullOrEmpty(cs.Value)) _callsignPrefixFilters = ParsePrefixes(cs.Value);
					if (mm != null && !string.IsNullOrEmpty(mm.Value)) _mmMessagesPath = mm.Value;
					if (fc != null && !string.IsNullOrEmpty(fc.Value)) _flightSchedCsvPath = fc.Value;
					int parsedWindow;
					if (cw != null && int.TryParse(cw.Value, out parsedWindow) && parsedWindow > 0) _conflictWindowHours = parsedWindow;
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
						new XElement("MmMessagesPath", _mmMessagesPath),
						new XElement("FlightSchedCsvPath", _flightSchedCsvPath)));
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

			Label hdr = new Label { Tag = "dispose", Top = 18, Left = 20, AutoSize = true,
				Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold), Text = "Admin — Archive service endpoints" };
			tabPage_Admin.Controls.Add(hdr);

			Label lbl1 = new Label { Tag = "dispose", Top = 60, Left = 20, AutoSize = true,
				Text = "Flight Schedule Service URL (getFlightList)" };
			tabPage_Admin.Controls.Add(lbl1);
			_adminFlightScheduleUrl = new TextBox { Tag = "dispose", Top = 82, Left = 20, Width = 520, Text = _flightScheduleBaseUrl };
			tabPage_Admin.Controls.Add(_adminFlightScheduleUrl);

			Label lbl2 = new Label { Tag = "dispose", Top = 118, Left = 20, AutoSize = true,
				Text = "Briefing Service URL (getBriefing)" };
			tabPage_Admin.Controls.Add(lbl2);
			_adminBriefingUrl = new TextBox { Tag = "dispose", Top = 140, Left = 20, Width = 520, Text = _briefingBaseUrl };
			tabPage_Admin.Controls.Add(_adminBriefingUrl);

			Label lbl3 = new Label { Tag = "dispose", Top = 176, Left = 20, AutoSize = true,
				Text = "Flight Schedule — callsign prefixes to keep (comma-separated)" };
			tabPage_Admin.Controls.Add(lbl3);
			_adminCallsignPrefixes = new TextBox { Tag = "dispose", Top = 198, Left = 20, Width = 520,
				Text = string.Join(",", _callsignPrefixFilters.ToArray()) };
			tabPage_Admin.Controls.Add(_adminCallsignPrefixes);

			Label lbl4 = new Label { Tag = "dispose", Top = 234, Left = 20, AutoSize = true,
				Text = "Conflict tab — window around STD/STA (± hours)" };
			tabPage_Admin.Controls.Add(lbl4);
			_adminConflictWindow = new TextBox { Tag = "dispose", Top = 256, Left = 20, Width = 80,
				Text = _conflictWindowHours.ToString() };
			tabPage_Admin.Controls.Add(_adminConflictWindow);

			Label lbl5 = new Label { Tag = "dispose", Top = 292, Left = 20, AutoSize = true,
				Text = "Movement Manager XML messages folder" };
			tabPage_Admin.Controls.Add(lbl5);
			_adminMmMessagesPath = new TextBox { Tag = "dispose", Top = 314, Left = 20, Width = 520, Text = _mmMessagesPath };
			tabPage_Admin.Controls.Add(_adminMmMessagesPath);

			Label lbl6 = new Label { Tag = "dispose", Top = 350, Left = 20, AutoSize = true,
				Text = "Flight Sched CSV folder (shared — every dispatcher's fallback CSV drop)" };
			tabPage_Admin.Controls.Add(lbl6);
			_adminFlightSchedCsvPath = new TextBox { Tag = "dispose", Top = 372, Left = 20, Width = 520, Text = _flightSchedCsvPath };
			tabPage_Admin.Controls.Add(_adminFlightSchedCsvPath);

			Button save = new Button { Tag = "dispose", Top = 408, Left = 20, Width = 100, Height = 30, Text = "Save" };
			save.Click += delegate
			{
				if (!EnsureWriterOrWarn()) return;
				_flightScheduleBaseUrl   = _adminFlightScheduleUrl.Text.Trim();
				_briefingBaseUrl         = _adminBriefingUrl.Text.Trim();
				_callsignPrefixFilters   = ParsePrefixes(_adminCallsignPrefixes.Text);
				_mmMessagesPath          = _adminMmMessagesPath.Text.Trim();
				_flightSchedCsvPath      = _adminFlightSchedCsvPath.Text.Trim();
				int parsedWindow;
				if (int.TryParse(_adminConflictWindow.Text.Trim(), out parsedWindow) && parsedWindow > 0)
					_conflictWindowHours = parsedWindow;
				SaveArchiveConfig();
				EnsureFlightSchedFolder();
				_adminStatus.Text = "Saved.";
			};
			tabPage_Admin.Controls.Add(save);

			_adminStatus = new Label { Tag = "dispose", Top = 416, Left = 130, AutoSize = true, ForeColor = Color.Gray, Text = "" };
			tabPage_Admin.Controls.Add(_adminStatus);
		}
	}
}
