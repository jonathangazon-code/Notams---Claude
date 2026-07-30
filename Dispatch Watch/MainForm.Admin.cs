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
		private static bool   _archiveConfigLoaded;

		private TextBox _adminFlightScheduleUrl;
		private TextBox _adminBriefingUrl;
		private TextBox _adminCallsignPrefixes;
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
					if (fs != null && !string.IsNullOrEmpty(fs.Value)) _flightScheduleBaseUrl = fs.Value;
					if (br != null && !string.IsNullOrEmpty(br.Value)) _briefingBaseUrl = br.Value;
					if (cs != null && !string.IsNullOrEmpty(cs.Value)) _callsignPrefixFilters = ParsePrefixes(cs.Value);
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
						new XElement("CallsignPrefixes", string.Join(",", _callsignPrefixFilters.ToArray()))));
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

			Button save = new Button { Tag = "dispose", Top = 234, Left = 20, Width = 100, Height = 30, Text = "Save" };
			save.Click += delegate
			{
				_flightScheduleBaseUrl   = _adminFlightScheduleUrl.Text.Trim();
				_briefingBaseUrl         = _adminBriefingUrl.Text.Trim();
				_callsignPrefixFilters   = ParsePrefixes(_adminCallsignPrefixes.Text);
				SaveArchiveConfig();
				_adminStatus.Text = "Saved.";
			};
			tabPage_Admin.Controls.Add(save);

			_adminStatus = new Label { Tag = "dispose", Top = 242, Left = 130, AutoSize = true, ForeColor = Color.Gray, Text = "" };
			tabPage_Admin.Controls.Add(_adminStatus);
		}
	}
}
