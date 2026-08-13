using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ICAO_CSV
{
	// TAF Analysis tab: downloads the FT (TAF) for every Airport List station from the fleet's
	// weather webservice, applies the same threshold/color-coding logic as the old standalone
	// "TAF analysis" app (TAF analysis/Dispatch/MainForm.cs, Btn_addTAFClick), and shows a
	// single shared, dispatcher-editable report — sent by email to its own recipient list,
	// independent of the Conflict tab's EmailRecipients.
	public partial class MainForm
	{
		// Threshold defaults match the legacy standalone app's hardcoded values. Persisted via
		// ArchiveConfig.xml (MainForm.Admin.cs's EnsureArchiveConfig/SaveArchiveConfig), but
		// edited from this tab's own top bar rather than the Admin tab, since these are the
		// dispatcher's day-to-day working values, not one-off endpoint config.
		private static int _tafVisCatIm              = 550;
		private static int _tafVisAdvisoryM          = 1000;
		private static int _tafCeilCatIHundredFt     = 2;
		private static int _tafCeilAdvisoryHundredFt = 4;
		private static int _tafWindCatIKt            = 44;
		private static int _tafWindAdvisoryKt        = 34;
		private static int _tafWindCatIMps           = 22;
		private static int _tafWindAdvisoryMps       = 17;

		private Panel _tafTopBar;
		private TextBox _tafVisCatIBox, _tafVisAdvBox, _tafCeilCatIBox, _tafCeilAdvBox,
			_tafWindKtCatIBox, _tafWindKtAdvBox, _tafWindMpsCatIBox, _tafWindMpsAdvBox;
		private bool _tafBodyLoaded;

		// ── TafRecipients: independent recipient list from the Conflict tab's EmailRecipients ──

		public void EnsureTafRecipientsTable()
		{
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				try { new OleDbCommand("CREATE TABLE TafRecipients ([Email] TEXT(255))", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }
				conn.Close();
			}
			catch { /* DB not ready */ }
		}

		List<string> LoadTafRecipients()
		{
			List<string> list = new List<string>();
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				OleDbDataReader r = new OleDbCommand("SELECT Email FROM TafRecipients ORDER BY Email", conn).ExecuteReader();
				while (r.Read())
					if (!r.IsDBNull(0) && r.GetString(0).Trim() != "") list.Add(r.GetString(0).Trim());
				conn.Close();
			}
			catch { }
			return list;
		}

		// ── TafReport: single shared row holding the current (possibly dispatcher-edited) report ──

		public void EnsureTafReportTable()
		{
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				try { new OleDbCommand("CREATE TABLE TafReport ([ID] LONG, [Html] MEMO, [TimeIssued] TEXT(50))", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }
				try
				{
					object count = new OleDbCommand("SELECT COUNT(*) FROM TafReport WHERE ID=1", conn).ExecuteScalar();
					if (Convert.ToInt32(count) == 0)
						new OleDbCommand("INSERT INTO TafReport ([ID],[Html],[TimeIssued]) VALUES (1,'','')", conn).ExecuteNonQuery();
				}
				catch { }
				conn.Close();
			}
			catch { /* DB not ready */ }
		}

		string LoadTafReportHtml(out string timeIssued)
		{
			timeIssued = "";
			string html = "";
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				OleDbDataReader r = new OleDbCommand("SELECT Html,TimeIssued FROM TafReport WHERE ID=1", conn).ExecuteReader();
				if (r.Read())
				{
					if (!r.IsDBNull(0)) html = r.GetString(0);
					if (!r.IsDBNull(1)) timeIssued = r.GetString(1);
				}
				conn.Close();
			}
			catch { }
			return html;
		}

		void SaveTafReportHtml(string html, string timeIssued)
		{
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				OleDbCommand upd = new OleDbCommand("UPDATE TafReport SET Html=?, TimeIssued=? WHERE ID=1", conn);
				upd.Parameters.AddWithValue("?", html);
				upd.Parameters.AddWithValue("?", timeIssued);
				upd.ExecuteNonQuery();
				conn.Close();
			}
			catch { }
		}

		// ── Tab lifecycle ──

		void TafAnalysisTabEnter(object sender, EventArgs e)
		{
			if (_tafTopBar == null) BuildTafTopBar();
			LoadTafReportIntoBrowser();
		}

		void TafAnalysisTabLeave(object sender, EventArgs e)
		{
			SaveEditedTafReport();
		}

		private void LoadTafReportIntoBrowser()
		{
			string timeIssued;
			string html = LoadTafReportHtml(out timeIssued);
			if (html == "")
				html = "<html><body style=\"font-family:Segoe UI, Arial, sans-serif; font-size:13px\">" +
					"No TAF analysis yet — click \"TAF Analysis\" to download and analyze.</body></html>";

			_tafBodyLoaded = false;
			Web_TafAnalysis.DocumentCompleted -= Web_TafAnalysisDocumentCompleted;
			Web_TafAnalysis.DocumentCompleted += Web_TafAnalysisDocumentCompleted;
			Web_TafAnalysis.DocumentText = html;
		}

		// contentEditable can only be applied once MSHTML has actually finished loading the
		// document — setting it right after DocumentText= is a no-op. Only a Writer's browser
		// is made editable; a Reader's stays genuinely read-only rather than looking editable
		// and silently discarding edits.
		void Web_TafAnalysisDocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
		{
			if (_tafBodyLoaded) return;
			_tafBodyLoaded = true;
			try
			{
				if (IsWriter && Web_TafAnalysis.Document != null && Web_TafAnalysis.Document.Body != null)
					Web_TafAnalysis.Document.Body.SetAttribute("contentEditable", "true");
			}
			catch { }
		}

		// Persists whatever the dispatcher currently sees/edited back to the shared TafReport
		// row. Called on tab Leave and again right before Send. No-ops for a Reader (whose
		// WebBrowser was never made contentEditable) or before anything has ever loaded.
		private void SaveEditedTafReport()
		{
			if (!IsWriter) return;
			try
			{
				if (Web_TafAnalysis.Document == null || Web_TafAnalysis.Document.Body == null) return;
				string html = Web_TafAnalysis.Document.Body.InnerHtml;
				if (html == null) return;
				string timeIssued;
				LoadTafReportHtml(out timeIssued); // keep the existing "last analysis" timestamp — editing text isn't a new analysis
				SaveTafReportHtml(html, timeIssued);
			}
			catch { }
		}

		// ── Top bar: TAF Analysis / Send / Recipients / Attach Image + threshold fields ──

		private void BuildTafTopBar()
		{
			_tafTopBar = new Panel { Dock = DockStyle.Top, Height = 92 };

			Button analyzeBtn = new Button
			{
				Top = 8, Left = 14, Width = 140, Height = 30, Text = "TAF Analysis",
				Font = new Font("Microsoft Sans Serif", 10f, FontStyle.Bold),
				BackColor = Color.SeaGreen, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, UseVisualStyleBackColor = false
			};
			analyzeBtn.Click += delegate { DownloadAndAnalyzeTafs(); };
			_tafTopBar.Controls.Add(analyzeBtn);

			Button sendBtn = new Button
			{
				Top = 8, Left = 164, Width = 160, Height = 30, Text = "✉  Send TAF Report",
				BackColor = Color.FromArgb(21, 101, 192), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, UseVisualStyleBackColor = false
			};
			sendBtn.Click += Btn_sendTafReportClick;
			_tafTopBar.Controls.Add(sendBtn);

			Button recipientsBtn = new Button
			{
				Top = 8, Left = 334, Width = 170, Height = 30, Text = "Update TAF Recipients",
				BackColor = Color.FromArgb(21, 101, 192), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, UseVisualStyleBackColor = false
			};
			recipientsBtn.Click += delegate { ShowTafRecipientsDialog(); };
			_tafTopBar.Controls.Add(recipientsBtn);

			Button attachBtn = new Button
			{
				Top = 8, Left = 514, Width = 130, Height = 30, Text = "Attach Image",
				BackColor = Color.FromArgb(69, 90, 100), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, UseVisualStyleBackColor = false
			};
			attachBtn.Click += delegate { AttachImageToTafReport(); };
			_tafTopBar.Controls.Add(attachBtn);

			int labelTop = 46, boxTop = 62, fieldWidth = 55, spacing = 92, x = 14;
			Font smallFont = new Font("Microsoft Sans Serif", 7.5f);

			Label lVisCatI = new Label { Top = labelTop, Left = x, AutoSize = true, Font = smallFont, Text = "Vis CAT I (m)" };
			_tafTopBar.Controls.Add(lVisCatI);
			_tafVisCatIBox = new TextBox { Top = boxTop, Left = x, Width = fieldWidth, Text = _tafVisCatIm.ToString() };
			_tafTopBar.Controls.Add(_tafVisCatIBox);
			x += spacing;

			Label lVisAdv = new Label { Top = labelTop, Left = x, AutoSize = true, Font = smallFont, Text = "Vis Adv (m)" };
			_tafTopBar.Controls.Add(lVisAdv);
			_tafVisAdvBox = new TextBox { Top = boxTop, Left = x, Width = fieldWidth, Text = _tafVisAdvisoryM.ToString() };
			_tafTopBar.Controls.Add(_tafVisAdvBox);
			x += spacing;

			Label lCeilCatI = new Label { Top = labelTop, Left = x, AutoSize = true, Font = smallFont, Text = "Ceil CAT I (x100ft)" };
			_tafTopBar.Controls.Add(lCeilCatI);
			_tafCeilCatIBox = new TextBox { Top = boxTop, Left = x, Width = fieldWidth, Text = _tafCeilCatIHundredFt.ToString() };
			_tafTopBar.Controls.Add(_tafCeilCatIBox);
			x += spacing;

			Label lCeilAdv = new Label { Top = labelTop, Left = x, AutoSize = true, Font = smallFont, Text = "Ceil Adv (x100ft)" };
			_tafTopBar.Controls.Add(lCeilAdv);
			_tafCeilAdvBox = new TextBox { Top = boxTop, Left = x, Width = fieldWidth, Text = _tafCeilAdvisoryHundredFt.ToString() };
			_tafTopBar.Controls.Add(_tafCeilAdvBox);
			x += spacing;

			Label lWindKtCatI = new Label { Top = labelTop, Left = x, AutoSize = true, Font = smallFont, Text = "Wind CAT I (kt)" };
			_tafTopBar.Controls.Add(lWindKtCatI);
			_tafWindKtCatIBox = new TextBox { Top = boxTop, Left = x, Width = fieldWidth, Text = _tafWindCatIKt.ToString() };
			_tafTopBar.Controls.Add(_tafWindKtCatIBox);
			x += spacing;

			Label lWindKtAdv = new Label { Top = labelTop, Left = x, AutoSize = true, Font = smallFont, Text = "Wind Adv (kt)" };
			_tafTopBar.Controls.Add(lWindKtAdv);
			_tafWindKtAdvBox = new TextBox { Top = boxTop, Left = x, Width = fieldWidth, Text = _tafWindAdvisoryKt.ToString() };
			_tafTopBar.Controls.Add(_tafWindKtAdvBox);
			x += spacing;

			Label lWindMpsCatI = new Label { Top = labelTop, Left = x, AutoSize = true, Font = smallFont, Text = "Wind CAT I (mps)" };
			_tafTopBar.Controls.Add(lWindMpsCatI);
			_tafWindMpsCatIBox = new TextBox { Top = boxTop, Left = x, Width = fieldWidth, Text = _tafWindCatIMps.ToString() };
			_tafTopBar.Controls.Add(_tafWindMpsCatIBox);
			x += spacing;

			Label lWindMpsAdv = new Label { Top = labelTop, Left = x, AutoSize = true, Font = smallFont, Text = "Wind Adv (mps)" };
			_tafTopBar.Controls.Add(lWindMpsAdv);
			_tafWindMpsAdvBox = new TextBox { Top = boxTop, Left = x, Width = fieldWidth, Text = _tafWindAdvisoryMps.ToString() };
			_tafTopBar.Controls.Add(_tafWindMpsAdvBox);
			x += spacing;

			Button saveThresholds = new Button
			{
				Top = 58, Left = x, Width = 120, Height = 26, Text = "Save Thresholds",
				BackColor = Color.FromArgb(69, 90, 100), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, UseVisualStyleBackColor = false
			};
			saveThresholds.Click += delegate { SaveTafThresholds(); };
			_tafTopBar.Controls.Add(saveThresholds);

			tabPage_TafAnalysis.Controls.Add(_tafTopBar);
			Web_TafAnalysis.BringToFront();
		}

		private void SaveTafThresholds()
		{
			if (!EnsureWriterOrWarn()) return;
			int v;
			if (int.TryParse(_tafVisCatIBox.Text.Trim(), out v) && v > 0) _tafVisCatIm = v;
			if (int.TryParse(_tafVisAdvBox.Text.Trim(), out v) && v > 0) _tafVisAdvisoryM = v;
			if (int.TryParse(_tafCeilCatIBox.Text.Trim(), out v) && v > 0) _tafCeilCatIHundredFt = v;
			if (int.TryParse(_tafCeilAdvBox.Text.Trim(), out v) && v > 0) _tafCeilAdvisoryHundredFt = v;
			if (int.TryParse(_tafWindKtCatIBox.Text.Trim(), out v) && v > 0) _tafWindCatIKt = v;
			if (int.TryParse(_tafWindKtAdvBox.Text.Trim(), out v) && v > 0) _tafWindAdvisoryKt = v;
			if (int.TryParse(_tafWindMpsCatIBox.Text.Trim(), out v) && v > 0) _tafWindCatIMps = v;
			if (int.TryParse(_tafWindMpsAdvBox.Text.Trim(), out v) && v > 0) _tafWindAdvisoryMps = v;
			SaveArchiveConfig();
			MessageBox.Show("Thresholds saved.", "TAF Analysis", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		// ── Download + analyze ──

		public void DownloadAndAnalyzeTafs()
		{
			if (!EnsureWriterOrWarn()) return;
			if (_stationsCache == null) LoadStationsCache();

			List<string> icaos = new List<string>();
			foreach (var entry in _stationsCache)
			{
				string[] row = entry.Value;
				if (row[1] == "Yes" || row[2] == "Yes" || row[3] == "Yes")
					icaos.Add(entry.Key);
			}
			if (icaos.Count == 0)
			{
				MessageBox.Show("No airports flagged LH/FedEx/Charters on the Airport List.", "TAF Analysis", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			// Same ddMMMyy computation GetXML() (MainForm.NotamData.cs) already uses.
			string todayStr = DateTime.Now.ToString("yyyyMMdd");
			string stringToday = todayStr.Substring(6, 2) + MonthAbbrev(todayStr.Substring(4, 2)) + todayStr.Substring(2, 2);

			string request = string.Join("-", icaos.ToArray());
			string url = _metBaseUrl.TrimEnd('/') + "/?METHOD=getAdHocMET&METTYPE=FT&REQUEST=" + request +
				"&PERIODSTART=" + stringToday + "&PERIODEND=" + stringToday;

			string xml;
			try
			{
				using (WebClient wc = new WebClient())
					xml = wc.DownloadString(url);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Failed to download TAFs:\n" + ex.Message, "TAF Analysis", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			// Parsed by local name only (same fix FetchBriefingDetails / MainForm.FlightSchedule.cs
			// already applies to this ARINC 633 namespace family) — matching on a fully-qualified
			// XName against the wrapping http://aeec.aviation-ia.net/633 namespace silently finds
			// nothing.
			Dictionary<string, string> rawByIcao = new Dictionary<string, string>();
			try
			{
				XDocument doc = XDocument.Parse(xml);
				foreach (XElement bulletin in doc.Descendants().Where(el => el.Name.LocalName == "WeatherBulletin"))
				{
					XElement icaoEl = bulletin.Descendants().FirstOrDefault(el => el.Name.LocalName == "AirportICAOCode");
					if (icaoEl == null || string.IsNullOrEmpty(icaoEl.Value)) continue;
					string icao = icaoEl.Value.Trim();

					StringBuilder sb = new StringBuilder();
					foreach (XElement textEl in bulletin.Descendants().Where(el => el.Name.LocalName == "Text"))
						sb.Append(textEl.Value).Append(" ");
					string raw = sb.ToString().Trim().TrimEnd('=').Trim();
					if (raw == "") continue;
					rawByIcao[icao] = raw; // last bulletin for a station wins if more than one is returned
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Failed to parse TAF response:\n" + ex.Message, "TAF Analysis", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			Dictionary<string, TafFragments> fragByIcao = new Dictionary<string, TafFragments>(StringComparer.OrdinalIgnoreCase);
			foreach (string icao in icaos)
			{
				string raw;
				if (!rawByIcao.TryGetValue(icao, out raw)) continue; // no FT returned for this station
				fragByIcao[icao] = AnalyzeTaf(raw);
			}

			// Organised like the legacy standalone app's db_read(): one section per ops type
			// (Long Haul / FedEx / Charters, same title colours), each broken into the same 4
			// sub-categories (Ceiling & Visibility, Wind, Thunderstorms, Snow) — a station only
			// appears under a category if that category actually flagged something for it.
			StringBuilder report = new StringBuilder();
			report.Append("<html><body style=\"font-family:Segoe UI, Arial, sans-serif; font-size:13px\">");
			report.Append("<span style=\"font-weight:bold;color:#888\">Last analysis: " +
				DateTime.Now.ToString("dd/MM/yyyy HH:mm") + " (local)</span><hr />");
			report.Append("<table style=\"text-align:left; font-size:13px\">");
			AppendOpsSection(report, "LH", "Long Haul Ops :", "RoyalBlue", icaos, fragByIcao);
			AppendOpsSection(report, "FedEx", "FedEx Ops :", "DarkMagenta", icaos, fragByIcao);
			AppendOpsSection(report, "Charters", "Charters Ops :", "Green", icaos, fragByIcao);
			report.Append("</table></body></html>");

			string timeIssued = DateTime.Now.ToString("dd/MM/yyyy HH:mm") + " (local)";
			SaveTafReportHtml(report.ToString(), timeIssued);
			LoadTafReportIntoBrowser();
		}

		// Per-station threshold-analysis result, one HTML fragment per category — mirrors the
		// legacy app's Vis_Ceiling/Wind/TS/Snow accumulator columns (TAF_analysis table there).
		// Each fragment only ever contains the tokens that actually matched that category (same
		// as the legacy accumulators), colored red/amber/blue exactly like AnalyzeTaf's combined
		// mode did before this change — it's the same underlying per-token analysis, just routed
		// into 4 separate buckets instead of one combined colored line.
		private struct TafFragments
		{
			public string VisCeil, Wind, TS, Snow;
		}

		// One ops-type section (title + hr), each broken into the 4 sub-category blocks below.
		private static void AppendOpsSection(StringBuilder report, string opsType, string title, string color,
			List<string> icaos, Dictionary<string, TafFragments> fragByIcao)
		{
			report.Append("<tr><th colspan=\"5\"><span style=\"font-weight:bold; font-size:16px; color:" + color + ";\">" +
				title + "</span></th></tr>");
			AppendCategory(report, "Ceiling & Visibility", opsType, icaos, fragByIcao, delegate(TafFragments f) { return f.VisCeil; });
			AppendCategory(report, "Wind", opsType, icaos, fragByIcao, delegate(TafFragments f) { return f.Wind; });
			AppendCategory(report, "Thunderstorms", opsType, icaos, fragByIcao, delegate(TafFragments f) { return f.TS; });
			AppendCategory(report, "Snow", opsType, icaos, fragByIcao, delegate(TafFragments f) { return f.Snow; });
			report.Append("<tr><th colspan=\"5\"><hr /></th></tr>");
		}

		private static void AppendCategory(StringBuilder report, string categoryTitle, string opsType, List<string> icaos,
			Dictionary<string, TafFragments> fragByIcao, Func<TafFragments, string> selector)
		{
			report.Append("<tr><th colspan=\"5\"><b>" + categoryTitle + "</b></th></tr><tr><th colspan=\"5\">");
			report.Append("<span style=\"font-family:'Courier New'; font-weight:normal;\">");
			bool any = false;
			foreach (string icao in icaos)
			{
				if (IsOpsType(opsType, icao) != "Yes") continue;
				TafFragments frag;
				if (!fragByIcao.TryGetValue(icao, out frag)) continue;
				string val = selector(frag);
				if (string.IsNullOrEmpty(val)) continue;
				string iata = GetIATA(icao);
				report.Append("<b style=\"color:DarkBlue;\">" + icao + (string.IsNullOrEmpty(iata) ? "" : " - " + iata) +
					" : </b>" + val + "<br />");
				any = true;
			}
			if (!any) report.Append("Nil");
			report.Append("</span></th></tr>");
		}

		// Direct port of the legacy "TAF analysis" app's per-station algorithm
		// (TAF analysis/Dispatch/MainForm.cs, Btn_addTAFClick, lines ~180-407) — the
		// station/time splitting that code needs is unnecessary here since the webservice XML
		// already hands back one clean forecast string per station. Like the legacy app, only
		// the specific value(s) that breached a threshold are shown per category (not the whole
		// TAF clause/token) — e.g. "OVC002" or "34012KT", colored by tier — routed into the same
		// 4 category buckets (VisCeil/Wind/TS/Snow) the report groups by.
		private static TafFragments AnalyzeTaf(string raw)
		{
			// Display granularity matches the legacy app exactly: each flagged trend group is
			// shown as its WHOLE block/clause (e.g. "BECMG 1310/1312 34012KT", not just
			// "34012KT"), so the hourly period and rest of that block stay attached to whatever
			// triggered it. Blocks are appended to their category bucket in the TAF's own
			// chronological order, so a breach followed later by its recovery (BECMG/FM back
			// above threshold) both land in the same fragment, one after another — the
			// surrounding trend is visible without extra bookkeeping. Colors match the legacy
			// app exactly: red for a CAT I breach, plain/uncolored text for the advisory tier,
			// blue for a trend-group recovery — no third "amber" color exists there.
			// Trend markers: BECMG/TEMPO/PROB30/PROB40/FM groups. The pattern's capturing group
			// means Regex.Split also returns each matched marker as its own array element.
			string patternTrend = @"(BECMG [0-9]{4}|TEMPO [0-9]{4}|PROB30 [0-9]{4}|PROB40 [0-9]{4}|FM[0-9]{4})";
			string[] parts = Regex.Split(raw, patternTrend);

			// PROB30/PROB40 + TEMPO reattachment: for "PROB30 TEMPO 1315/1318 ...", the TEMPO
			// alternative matches starting at "TEMPO", leaving "PROB30 " dangling on the END of
			// the PRECEDING chunk instead of prefixed onto the delimiter — ported verbatim from
			// the legacy app's two-pass fixup (lines 194-218 there).
			for (int i = 0; i < parts.Length - 1; i++)
			{
				int len = parts[i].Length;
				string end7 = len > 7 ? parts[i].Substring(len - 7, 7) : "";
				if (end7 == "PROB30 " || end7 == "PROB40 ")
				{
					parts[i] = parts[i].Substring(0, len - 7);
					parts[i + 1] = end7 + parts[i + 1];
					continue;
				}
				string end10 = len > 10 ? parts[i].Substring(len - 10, 7) : "";
				if (end10 == "PROB30 " || end10 == "PROB40 ")
				{
					parts[i] = parts[i].Substring(0, len - 10);
					parts[i + 1] = end10 + parts[i + 1];
				}
			}

			// Break into display tokens — each trend group starts its own token unless it's a
			// bare continuation (leading "/" or "0"), matching the legacy app's <br/> heuristic.
			StringBuilder joined = new StringBuilder();
			foreach (string part in parts)
			{
				string lead = part.Length > 0 ? part.Substring(0, 1) : "-";
				if (lead == "/" || lead == "0") joined.Append(part);
				else joined.Append("<br />").Append(part);
			}
			string[] tokens = Regex.Split(joined.ToString(), "<br />");

			StringBuilder visCeil = new StringBuilder(), wind = new StringBuilder(), ts = new StringBuilder(), snow = new StringBuilder();
			bool visTrend = false, windTrend = false;
			foreach (string token in tokens)
			{
				if (token.Trim() == "") continue;
				string lead2 = token.Length >= 2 ? token.Substring(0, 2) : "";
				bool isTrendMarker = lead2 == "BE" || lead2 == "FM" || (token.Length > 0 && char.IsDigit(token[0]));

				// Vis + ceiling flags are aggregated across every match in this token, then the
				// WHOLE block is appended once (red wins over the advisory tier over a trend
				// recovery) — exactly the legacy app's ceilCatI/ceilTresh/visCatI/visTresh
				// aggregate-then-decide-once pattern (lines 300-317 there), not a per-match append.
				bool visCatI = false, visTresh = false, ceilCatI = false, ceilTresh = false;
				foreach (Match m in Regex.Matches(token, @"(?<= )\b[0-9]{4,4}\b(?= )"))
				{
					int val = int.Parse(m.Value);
					if (val < _tafVisCatIm) { visCatI = true; if (isTrendMarker) visTrend = true; }
					else if (val <= _tafVisAdvisoryM) { visTresh = true; if (isTrendMarker) visTrend = true; }
				}
				foreach (Match m in Regex.Matches(token, @"(?<=BKN|OVC|VV)[0-9]{3}"))
				{
					int val = int.Parse(m.Value);
					if (val <= _tafCeilCatIHundredFt) { ceilCatI = true; if (isTrendMarker) visTrend = true; }
					else if (val <= _tafCeilAdvisoryHundredFt) { ceilTresh = true; if (isTrendMarker) visTrend = true; }
				}
				if (ceilCatI || visCatI) AppendBlock(visCeil, token, "red");
				else if (ceilTresh || visTresh) AppendBlock(visCeil, token, null);
				else if (isTrendMarker && visTrend) { AppendBlock(visCeil, token, "blue"); visTrend = false; }

				// Wind KT/MPS append per match (as the legacy app does), since a token normally
				// carries at most one wind group.
				foreach (Match m in Regex.Matches(token, @"([a-zA-Z0-9]{5,8})KT"))
				{
					string w = m.Value;
					string spd = w.Length == 10 ? w.Substring(6, 2) : w.Substring(3, 2);
					int kt;
					if (int.TryParse(spd, out kt))
					{
						if (kt > _tafWindCatIKt) { AppendBlock(wind, token, "red"); if (isTrendMarker) windTrend = true; }
						else if (kt >= _tafWindAdvisoryKt) { AppendBlock(wind, token, null); if (isTrendMarker) windTrend = true; }
						else if (isTrendMarker && windTrend) { AppendBlock(wind, token, "blue"); windTrend = false; }
					}
				}
				foreach (Match m in Regex.Matches(token, @"([a-zA-Z0-9]{5,8})MPS"))
				{
					string w = m.Value;
					string spd = w.Length == 11 ? w.Substring(6, 2) : w.Substring(3, 2);
					int mps;
					if (int.TryParse(spd, out mps))
					{
						if (mps > _tafWindCatIMps) { AppendBlock(wind, token, "red"); if (isTrendMarker) windTrend = true; }
						else if (mps >= _tafWindAdvisoryMps) { AppendBlock(wind, token, null); if (isTrendMarker) windTrend = true; }
						else if (isTrendMarker && windTrend) { AppendBlock(wind, token, "blue"); windTrend = false; }
					}
				}
				// TS / snow / freezing precip — always flagged, no threshold tiers, plain text
				// (unconditional, exactly like the legacy app's TS+=valueBr / SN+=valueBr).
				if (Regex.IsMatch(token, "TS")) AppendBlock(ts, token, null);
				if (Regex.IsMatch(token, "SN|FZRA|FZDZ")) AppendBlock(snow, token, null);
			}

			TafFragments frag;
			frag.VisCeil = visCeil.ToString();
			frag.Wind = wind.ToString();
			frag.TS = ts.ToString();
			frag.Snow = snow.ToString();
			return frag;
		}

		// Appends the whole trend-group block/clause, colored (or plain, when color is null —
		// the advisory tier's uncolored text, matching the legacy app's plain "sb+=valueBr").
		private static void AppendBlock(StringBuilder sb, string block, string color)
		{
			if (color == null) { sb.Append(block); return; }
			sb.Append("<span style=\"color:").Append(color).Append(color == "red" ? ";font-weight:bold\">" : "\">").Append(block).Append("</span>");
		}

		// ── Attach image (inserted directly into the editable report) ──

		private void AttachImageToTafReport()
		{
			if (!EnsureWriterOrWarn()) return;
			using (OpenFileDialog dlg = new OpenFileDialog { Filter = "Images|*.png;*.jpg;*.jpeg;*.gif;*.bmp", Title = "Attach image to TAF report" })
			{
				if (dlg.ShowDialog() != DialogResult.OK) return;
				try
				{
					// Copied into %TEMP% under a unique name so a since-moved/deleted/overwritten
					// source file can't break re-rendering the saved report later.
					string ext = Path.GetExtension(dlg.FileName);
					string tempPath = Path.Combine(Path.GetTempPath(), "taf_img_" + Guid.NewGuid().ToString("N") + ext);
					File.Copy(dlg.FileName, tempPath, true);
					string src = new Uri(tempPath).AbsoluteUri;
					string imgTag = "<br /><img src=\"" + src + "\" style=\"max-width:600px\" />";

					if (Web_TafAnalysis.Document != null && Web_TafAnalysis.Document.Body != null)
					{
						Web_TafAnalysis.Document.Body.InnerHtml = Web_TafAnalysis.Document.Body.InnerHtml + imgTag;
						SaveEditedTafReport();
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show("Could not attach image:\n" + ex.Message, "Attach Image", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		// ── Recipients dialog (clone of MainForm.Conflict.cs's ShowRecipientsDialog, targeting
		// TafRecipients instead of EmailRecipients) ──

		private void ShowTafRecipientsDialog()
		{
			using (Form dlg = new Form
			{
				Text = "Dispatch Watch", FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterScreen,
				ControlBox = false, MinimizeBox = false, MaximizeBox = false, Width = 440, Height = 480, BackColor = Color.FromArgb(38, 50, 56)
			})
			{
				Label title = new Label { Text = "TAF recipients", ForeColor = Color.White,
					Font = new Font("Segoe UI", 11f, FontStyle.Bold), Top = 14, Left = 20, AutoSize = true };
				dlg.Controls.Add(title);

				ListBox list = new ListBox { Top = 48, Left = 20, Width = 380, Height = 300, BackColor = Color.White, ForeColor = Color.Black };
				foreach (string a in LoadTafRecipients()) list.Items.Add(a);
				dlg.Controls.Add(list);

				TextBox addBox = new TextBox { Top = 358, Left = 20, Width = 270, BackColor = Color.White, ForeColor = Color.Black };
				dlg.Controls.Add(addBox);

				Button add = new Button { Text = "Add", Top = 356, Left = 300, Width = 100, Height = 26,
					BackColor = Color.FromArgb(46, 125, 82), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
				add.FlatAppearance.BorderColor = Color.FromArgb(96, 125, 139);
				dlg.Controls.Add(add);

				Action doAdd = delegate
				{
					string a = addBox.Text.Trim();
					if (a == "" || !a.Contains("@")) return;
					foreach (object it in list.Items)
						if (string.Equals(it.ToString(), a, StringComparison.OrdinalIgnoreCase)) { addBox.Clear(); return; }
					if (!EnsureWriterOrWarn()) return;

					OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
					conn.Open();
					OleDbCommand ins = new OleDbCommand("INSERT INTO TafRecipients ([Email]) VALUES (?)", conn);
					ins.Parameters.AddWithValue("?", a);
					ins.ExecuteNonQuery();
					conn.Close();

					addBox.Clear();
					list.Items.Clear();
					foreach (string r in LoadTafRecipients()) list.Items.Add(r);
				};
				add.Click += delegate { doAdd(); };
				addBox.KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { doAdd(); e.SuppressKeyPress = true; } };

				Button remove = new Button { Text = "Remove selected", Top = 392, Left = 20, Width = 380, Height = 28,
					BackColor = Color.FromArgb(200, 40, 40), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
				remove.FlatAppearance.BorderColor = Color.FromArgb(96, 125, 139);
				remove.Click += delegate
				{
					if (list.SelectedItem == null) return;
					string a = list.SelectedItem.ToString();
					if (!EnsureWriterOrWarn()) return;

					OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
					conn.Open();
					OleDbCommand del = new OleDbCommand("DELETE FROM TafRecipients WHERE Email=?", conn);
					del.Parameters.AddWithValue("?", a);
					del.ExecuteNonQuery();
					conn.Close();

					list.Items.Clear();
					foreach (string r in LoadTafRecipients()) list.Items.Add(r);
				};
				dlg.Controls.Add(remove);

				Button close = new Button { Text = "Close", Top = 430, Left = 20, Width = 380, Height = 28,
					BackColor = Color.FromArgb(69, 90, 100), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
				close.FlatAppearance.BorderColor = Color.FromArgb(96, 125, 139);
				close.Click += delegate { dlg.Close(); };
				dlg.Controls.Add(close);

				dlg.ShowDialog();
			}
		}

		// ── Send (clone of MainForm.Email.cs's Btn_sendReportsClick's Outlook-COM mechanics,
		// simplified: no PDF attachments, body = the current TafReport HTML) ──

		void Btn_sendTafReportClick(object sender, EventArgs e)
		{
			if (!EnsureWriterOrWarn()) return;
			SaveEditedTafReport();

			string timeIssued;
			string bodyHtml = LoadTafReportHtml(out timeIssued);
			if (string.IsNullOrEmpty(bodyHtml) || bodyHtml.IndexOf("No TAF analysis yet", StringComparison.OrdinalIgnoreCase) >= 0)
			{
				MessageBox.Show("No TAF analysis to send yet. Click \"TAF Analysis\" first.", "Send TAF Report", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			List<string> rcp = LoadTafRecipients();
			if (rcp.Count == 0)
			{
				MessageBox.Show("No TAF recipients defined. Add at least one address.", "Send TAF Report", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			if (MessageBox.Show("Send the current TAF report to " + rcp.Count + " recipient(s)?",
				"Send TAF Report", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

			string titleDate = DateTime.Now.ToString("ddMMMyyyy", System.Globalization.CultureInfo.InvariantCulture).ToUpper();
			string subject = "TAF Analysis - " + titleDate;
			string fullBody = "<html><body style=\"font-family:Segoe UI, Arial, sans-serif; font-size:13px\">" + bodyHtml + "</body></html>";

			string step = "init";
			try
			{
				Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
				if (outlookType == null)
				{
					MessageBox.Show("Outlook is not installed or registered on this machine.", "Send TAF Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
				step = "CreateInstance";
				object outlook = Activator.CreateInstance(outlookType);
				step = "CreateItem";
				object mail = outlookType.InvokeMember("CreateItem", BindingFlags.InvokeMethod, null, outlook, new object[] { 0 }); // 0 = olMailItem
				Type mt = mail.GetType();
				step = "Subject";
				mt.InvokeMember("Subject", BindingFlags.SetProperty, null, mail, new object[] { subject });

				step = "Recipients";
				object recips = mt.InvokeMember("Recipients", BindingFlags.GetProperty, null, mail, null);
				Type rct = recips.GetType();
				foreach (string addr in rcp)
				{
					object r = rct.InvokeMember("Add", BindingFlags.InvokeMethod, null, recips, new object[] { addr });
					try { r.GetType().InvokeMember("Type", BindingFlags.SetProperty, null, r, new object[] { 1 }); } catch { } // 1 = olTo
				}
				bool allResolved = (bool)rct.InvokeMember("ResolveAll", BindingFlags.InvokeMethod, null, recips, null);
				if (!allResolved)
				{
					List<string> bad = new List<string>();
					int rcount = (int)rct.InvokeMember("Count", BindingFlags.GetProperty, null, recips, null);
					for (int i = 1; i <= rcount; i++)
					{
						object r = rct.InvokeMember("Item", BindingFlags.GetProperty, null, recips, new object[] { i });
						Type rrt = r.GetType();
						bool res = (bool)rrt.InvokeMember("Resolved", BindingFlags.GetProperty, null, r, null);
						if (!res) bad.Add((string)rrt.InvokeMember("Name", BindingFlags.GetProperty, null, r, null));
					}
					MessageBox.Show("Outlook could not recognise these address(es):\n\n" + string.Join("\n", bad.ToArray()) +
						"\n\nFix or remove them in the TAF recipients list, then try again.",
						"Send TAF Report", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				step = "Body";
				int bodyCloseIdx = fullBody.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
				string tail = ReadDefaultSignatureHtml();
				fullBody = bodyCloseIdx >= 0 ? fullBody.Insert(bodyCloseIdx, tail) : fullBody + tail;
				mt.InvokeMember("HTMLBody", BindingFlags.SetProperty, null, mail, new object[] { fullBody });

				step = "Send";
				mt.InvokeMember("Send", BindingFlags.InvokeMethod, null, mail, null);

				MessageBox.Show("TAF report sent to " + rcp.Count + " recipient(s).", "Send TAF Report", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				string msg = (ex.InnerException != null) ? ex.InnerException.Message : ex.Message;
				MessageBox.Show("Failed to send via Outlook (step: " + step + "):\n" + msg, "Send TAF Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}
