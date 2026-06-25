using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		private static readonly string[] _notamKeywords = {
			"CLSD", "U/S", "OTS", "UNSERVICEABLE", "OUT OF SERVICE",
			"ILS", "GP", "LOC", "RWY", "TWY", "APCH", "DEP",
			"FUEL", "AVBL", "NOT AVBL", "NIL", "LTD",
			"CAT I", "CAT II", "CAT III", "PERM", "H24", "DAILY"
		};

		private static Color ImpactColor(string impact)
		{
			switch (impact)
			{
				case "A":  return Color.FromArgb(200, 40,  40);
				case "R":  return Color.FromArgb(200, 90,   0);
				case "C":  return Color.FromArgb(170, 100,  0);
				case "N":  return Color.FromArgb(150, 80,   0);
				case "D":  return Color.FromArgb(100,  0, 150);
				case "F":  return Color.FromArgb(  0, 60, 180);
				case "AS": return Color.FromArgb(  0, 120, 60);
				default:   return Color.FromArgb( 80, 80,  80);
			}
		}

		private static string ImpactLabel(string impact)
		{
			switch (impact)
			{
				case "A":  return "AP CLSD";
				case "R":  return "RWY CLSD";
				case "C":  return "CAT I";
				case "N":  return "No ILS";
				case "D":  return "Not ALTN";
				case "F":  return "FUEL";
				case "AS": return "AIP SUP";
				case "M":  return "MISC";
				default:   return "";
			}
		}

		private static void AppendRtb(RichTextBox rtb, string text, Color color, bool bold, float size = 10f)
		{
			rtb.SelectionStart  = rtb.TextLength;
			rtb.SelectionLength = 0;
			rtb.SelectionColor  = color;
			rtb.SelectionFont   = new Font("Courier New", size, bold ? FontStyle.Bold : FontStyle.Regular);
			rtb.AppendText(text);
		}

		private static void HighlightKeywords(RichTextBox rtb)
		{
			Font boldFont = new Font("Courier New", 10, FontStyle.Bold);
			Color kwColor = Color.FromArgb(180, 0, 0);
			string txt = rtb.Text;
			foreach (string kw in _notamKeywords)
			{
				int idx = 0;
				while (true)
				{
					idx = txt.IndexOf(kw, idx, StringComparison.OrdinalIgnoreCase);
					if (idx < 0) break;
					rtb.SelectionStart  = idx;
					rtb.SelectionLength = kw.Length;
					rtb.SelectionFont   = boldFont;
					rtb.SelectionColor  = kwColor;
					idx += kw.Length;
				}
			}
			rtb.SelectionStart  = 0;
			rtb.SelectionLength = 0;
		}

		// Builds a VML airport diagram (WebBrowser control runs IE7 mode -> no inline SVG).
		// Each physical runway is a strip oriented by its QFU (designator x 10 deg),
		// scaled by its length, with both threshold labels.
		private static string BuildRwySvg(System.Collections.Generic.List<string> rwyClean)
		{
			int W = 130, H = 110, cx = 65, cy = 55, maxHalf = 40;

			System.Collections.Generic.List<int>    headings = new System.Collections.Generic.List<int>();
			System.Collections.Generic.List<double> lengths  = new System.Collections.Generic.List<double>();
			System.Collections.Generic.List<string> end1     = new System.Collections.Generic.List<string>();
			System.Collections.Generic.List<string> end2     = new System.Collections.Generic.List<string>();

			for (int i = 0; i < rwyClean.Count; i += 2)
			{
				string d1 = ParseDesignator(rwyClean[i]);
				string d2 = (i + 1 < rwyClean.Count) ? ParseDesignator(rwyClean[i + 1]) : "";
				int hdg = ParseHeading(d1);
				if (hdg < 0) continue;
				double len = ParseLength(rwyClean[i]);
				headings.Add(hdg);
				lengths.Add(len);
				end1.Add(d1);
				end2.Add(d2);
			}

			if (headings.Count == 0) return "";

			double maxLen = 0;
			foreach (double l in lengths) if (l > maxLen) maxLen = l;
			if (maxLen <= 0) maxLen = 1;

			System.Text.StringBuilder shapes = new System.Text.StringBuilder();
			System.Text.StringBuilder labels = new System.Text.StringBuilder();

			for (int i = 0; i < headings.Count; i++)
			{
				double half = maxHalf * (lengths[i] > 0 ? lengths[i] / maxLen : 1.0);
				if (half < 12) half = 12;
				double rad = headings[i] * Math.PI / 180.0;
				double dx = Math.Sin(rad) * half;
				double dy = -Math.Cos(rad) * half;

				double x1 = cx - dx, y1 = cy - dy;
				double x2 = cx + dx, y2 = cy + dy;

				shapes.Append("<v:line style=\"position:absolute\" from=\"" + F(x1) + "," + F(y1) +
					"\" to=\"" + F(x2) + "," + F(y2) + "\" strokecolor=\"#607d8b\" strokeweight=\"7px\">" +
					"<v:stroke endcap=\"round\"/></v:line>");
				shapes.Append("<v:line style=\"position:absolute\" from=\"" + F(x1) + "," + F(y1) +
					"\" to=\"" + F(x2) + "," + F(y2) + "\" strokecolor=\"#cfd8dc\" strokeweight=\"1px\">" +
					"<v:stroke dashstyle=\"dash\"/></v:line>");
				labels.Append(RwyLabel(end1[i], x1, y1, cx, cy));
				if (end2[i] != "") labels.Append(RwyLabel(end2[i], x2, y2, cx, cy));
			}

			return "<div style=\"position:relative;width:" + W + "px;height:" + H + "px\">" +
				shapes.ToString() + labels.ToString() + "</div>";
		}

		private static string RwyLabel(string text, double x, double y, double cx, double cy)
		{
			double ox = (x - cx) * 0.22, oy = (y - cy) * 0.22;
			double lx = x + ox - 9, ly = y + oy - 7;
			return "<div style=\"position:absolute;left:" + F(lx) + "px;top:" + F(ly) +
				"px;width:18px;text-align:center;font-size:9px;color:#b0bec5;font-family:monospace\">" + text + "</div>";
		}

		private static string F(double v) { return v.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture); }

		private static string ParseDesignator(string line)
		{
			int colon = line.IndexOf(':');
			string d = colon > 0 ? line.Substring(0, colon) : line;
			return d.Trim();
		}

		private static int ParseHeading(string designator)
		{
			string digits = "";
			foreach (char c in designator)
			{
				if (c >= '0' && c <= '9') digits += c;
				else break;
			}
			if (digits.Length == 0) return -1;
			int n;
			if (!Int32.TryParse(digits, out n)) return -1;
			return (n * 10) % 360;
		}

		private static double ParseLength(string line)
		{
			// Find a number immediately followed by 'm' (e.g. 3054m)
			for (int i = 0; i < line.Length; i++)
			{
				if ((line[i] == 'm' || line[i] == 'M') && i > 0 && line[i - 1] >= '0' && line[i - 1] <= '9')
				{
					int j = i - 1;
					while (j >= 0 && line[j] >= '0' && line[j] <= '9') j--;
					string num = line.Substring(j + 1, i - j - 1);
					double val;
					if (Double.TryParse(num, System.Globalization.NumberStyles.Integer,
						System.Globalization.CultureInfo.InvariantCulture, out val))
						return val;
				}
			}
			return 0;
		}

		void Filter_Notams()
		{
			tabPage1.VerticalScroll.Value = 0;
			ClearTaggedControls(tabPage1);

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();

			string AP = "";
			OleDbDataReader APreader = new OleDbCommand(
				"SELECT location FROM filteredNotams_table WHERE (Checked='N') ORDER BY location DESC", conn).ExecuteReader();
			while (APreader.Read())
				if (!APreader.IsDBNull(0)) AP = APreader.GetString(0);
			conn.Close();

			OleDbConnection connOCC = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			connOCC.Open();
			OleDbCommand cmdOCC = new OleDbCommand("SELECT * FROM Stations_ICAO_IATA WHERE ICAO=?", connOCC);
			cmdOCC.Parameters.AddWithValue("?", AP);
			OleDbDataReader OCCreader = cmdOCC.ExecuteReader();
			string RWYs = "";
			while (OCCreader.Read())
				if (!OCCreader.IsDBNull(6)) RWYs = OCCreader.GetString(6);
			connOCC.Close();

			// Header WebBrowser: Option D style (dark anthracite)
			string[] rwyList = RWYs.Split('/');
			int rwyCount = 0;
			foreach (string r in rwyList) if (r.Trim() != "") rwyCount++;
			int headerHeight = 0;

			string iata = GetIATA(AP);
			string[] rwyLines = RWYs.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
			System.Collections.Generic.List<string> rwyClean = new System.Collections.Generic.List<string>();
			foreach (string rl in rwyLines) if (rl.Trim() != "") rwyClean.Add(rl.Trim());
			int pairRows = (rwyClean.Count + 1) / 2;
			headerHeight = Math.Max(60 + pairRows * 22 + 12, 60 + 120);
			Web_FilterHeader.Size = new Size(490, headerHeight);

			string iataLine = (iata != "" && iata != AP) ? "<div class=\"sub\">IATA: " + iata + "</div>" : "";
			string leftCol = "", rightCol = "";
			for (int i = 0; i < rwyClean.Count; i++)
			{
				string cell = "<div class=\"rwyline\">" +
					rwyClean[i].Replace("&", "&amp;").Replace("<", "&lt;") + "</div>";
				if (i % 2 == 0) leftCol += cell; else rightCol += cell;
			}
			string rwySvg = BuildRwySvg(rwyClean);
			Web_FilterHeader.DocumentText =
				"<html xmlns:v=\"urn:schemas-microsoft-com:vml\"><head><style>" +
				"v\\:*{behavior:url(#default#VML)}" +
				"body{margin:0;padding:10px 14px;background:#263238;font-family:'Courier New',monospace;overflow:hidden}" +
				".icao{font-size:18px;font-weight:bold;color:#eceff1;letter-spacing:3px}" +
				".sub{font-size:11px;color:#78909c;margin-top:1px;margin-bottom:8px}" +
				".blk{float:left;font-size:11px;color:#b0bec5;background:#37474f;border-left:2px solid #546e7a;padding:5px 12px;margin-top:8px;margin-right:10px}" +
				".rwyline{white-space:nowrap;line-height:1.9}" +
				".diagram{float:right;margin-top:4px}" +
				"</style></head><body>" +
				"<div class=\"diagram\">" + rwySvg + "</div>" +
				"<div class=\"icao\">" + AP + "</div>" +
				iataLine +
				"<div style=\"clear:left\"></div>" +
				"<div class=\"blk\">" + leftCol + "</div>" +
				"<div class=\"blk\">" + rightCol + "</div>" +
				"</body></html>";

			// Per-NOTAM RTBs with colored left border strip (Option B)
			int keptTop = Web_FilterHeader.Bottom + 8;
			conn.Open();
			OleDbCommand cmdKept = new OleDbCommand(
				"SELECT * FROM filteredNotams_table WHERE (Status='K') AND (location=?)", conn);
			cmdKept.Parameters.AddWithValue("?", AP);
			OleDbDataReader keptReader = cmdKept.ExecuteReader();
			while (keptReader.Read())
			{
				string fromDate = !keptReader.IsDBNull(5)  ? keptReader.GetString(5)  : "";
				string tillDate = !keptReader.IsDBNull(6)  ? keptReader.GetString(6)  : "";
				string text     = !keptReader.IsDBNull(7)  ? keptReader.GetString(7)  : "";
				string notamKey = !keptReader.IsDBNull(10) ? keptReader.GetString(10) : "";
				string Impact   = !keptReader.IsDBNull(13) ? keptReader.GetString(13) : "";
				string Remark   = !keptReader.IsDBNull(14) ? keptReader.GetString(14) : "";
				text = text.Replace("(char)39", "'");

				Color ic = ImpactColor(Impact);
				string ilabel = ImpactLabel(Impact);

				RichTextBox rtb = new RichTextBox
				{
					Tag = "dispose", Left = 6, Top = 0, Width = 482, Height = 2000,
					BorderStyle = BorderStyle.None, ReadOnly = true,
					BackColor = SystemColors.Window, ScrollBars = RichTextBoxScrollBars.None
				};
				AppendRtb(rtb, notamKey, ic, true, 11f);
				if (ilabel != "") AppendRtb(rtb, "  [" + ilabel + "]", ic, false, 9f);
				AppendRtb(rtb, "\n", Color.Black, false);
				AppendRtb(rtb, FormatDate(fromDate) + "  →  " + FormatDate(tillDate) + "\n", Color.DimGray, false, 9f);
				AppendRtb(rtb, text + "\n", Color.Black, false, 10f);
				if (Remark != "") AppendRtb(rtb, "▶ " + Remark + "\n", ic, false, 9f);
				HighlightKeywords(rtb);

				int totalLines = rtb.GetLineFromCharIndex(rtb.TextLength) + 1;
				int rtbHeight  = totalLines * 17 + 12;
				rtb.Height = rtbHeight;

				Panel container = new Panel { Tag = "dispose", Left = 7, Top = keptTop, Width = 490, Height = rtbHeight };
				Panel strip     = new Panel { Left = 0, Top = 0, Width = 4, Height = rtbHeight, BackColor = ic };
				container.Controls.Add(strip);
				container.Controls.Add(rtb);
				tabPage1.Controls.Add(container);
				keptTop += rtbHeight + 6;
			}
			conn.Close();

			// New unchecked NOTAMs
			FontFamily courier = new FontFamily("Courier New");
			conn.Open();
			OleDbCommand cmdNew = new OleDbCommand(
				"SELECT * FROM filteredNotams_table WHERE (Checked='N') AND (location=?)", conn);
			cmdNew.Parameters.AddWithValue("?", AP);
			OleDbDataReader dBreader = cmdNew.ExecuteReader();

			int nbNotams = 0;
			int Top = 100;

			Dictionary<int, Button>      keep_Buttons       = new Dictionary<int, Button>();
			Dictionary<int, RichTextBox> RchTxt_notam_text  = new Dictionary<int, RichTextBox>();
			Dictionary<int, CheckBox>    apt_CLSD_Chckbox   = new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_CATI_Chckbox   = new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_NILS_Chckbox   = new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_NOALTN_Chckbox = new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_FUEL_Chckbox   = new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_MISC_Chckbox   = new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_AIPSUP_Chckbox = new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_RWYCLSD_Chckbox= new Dictionary<int, CheckBox>();
			Dictionary<int, TextBox>     remark_Txtbox      = new Dictionary<int, TextBox>();
			Dictionary<int, Button>      remark_Buttons     = new Dictionary<int, Button>();

			while (dBreader.Read())
			{
				int    notam_ID   = !dBreader.IsDBNull(0)  ? dBreader.GetInt32(0)   : 0;
				string fromDate   = !dBreader.IsDBNull(5)  ? dBreader.GetString(5)  : "";
				string tillDate   = !dBreader.IsDBNull(6)  ? dBreader.GetString(6)  : "";
				string notam_text = !dBreader.IsDBNull(7)  ? dBreader.GetString(7)  : "";
				string notam_key  = !dBreader.IsDBNull(10) ? dBreader.GetString(10) : "";
				string Status     = !dBreader.IsDBNull(12) ? dBreader.GetString(12) : "";
				string Impact     = !dBreader.IsDBNull(13) ? dBreader.GetString(13) : "";
				string Remark     = !dBreader.IsDBNull(14) ? dBreader.GetString(14) : "";
				notam_text = notam_text.Replace("(char)39", "'");

				Color keyColor = HasImpact(Impact) ? ImpactColor(Impact) : Color.MidnightBlue;

				AddNotamLabel(tabPage1, courier, notam_key, Top, 510, 140, keyColor, true, 11f);
				AddNotamLabel(tabPage1, courier, FormatDate(fromDate), Top,    655, 190, Color.DimGray, false, 10f);
				AddNotamLabel(tabPage1, courier, FormatDate(tillDate), Top,    850, 190, Color.DimGray, false, 10f);

				int height = Math.Max(notam_text.Length / 50 * 20, 80);
				RchTxt_notam_text[notam_ID] = MakeNotamRichText(courier, notam_text, Top+20, 510, height, Status == "K");
				HighlightKeywords(RchTxt_notam_text[notam_ID]);
				tabPage1.Controls.Add(RchTxt_notam_text[notam_ID]);

				keep_Buttons[notam_ID] = MakeKeepButton(courier, Status, new Point(1070, Top+20));
				int nid = notam_ID;
				if (Status == "")  keep_Buttons[notam_ID].Click += (s, e) => Keep_Notam(nid);
				if (Status == "K") keep_Buttons[notam_ID].Click += (s, e) => Ignore_Notam(nid);
				tabPage1.Controls.Add(keep_Buttons[notam_ID]);

				if (Status == "K")
				{
					AddImpactCheckboxes(tabPage1, notam_ID, Impact, Top, 1070, 1150, 1240, 1330,
						apt_CLSD_Chckbox, apt_CATI_Chckbox, apt_NILS_Chckbox,
						apt_NOALTN_Chckbox, apt_FUEL_Chckbox, apt_MISC_Chckbox,
						apt_AIPSUP_Chckbox, apt_RWYCLSD_Chckbox);

					if (HasImpact(Impact))
					{
						remark_Txtbox[notam_ID]  = new TextBox { Tag="dispose", Top=Top+94, Left=1070, Size=new Size(250,24), Text=Remark };
						remark_Buttons[notam_ID] = new Button  { Tag="dispose", Top=Top+92, Left=1320, Size=new Size(40,24),  Text="OK" };
						int ri = notam_ID;
						remark_Buttons[notam_ID].Click += (s, e) => Remark_Notam(ri, remark_Txtbox[ri].Text);
						tabPage1.Controls.Add(remark_Txtbox[notam_ID]);
						tabPage1.Controls.Add(remark_Buttons[notam_ID]);
					}
				}

				Top = Top + height + 30;
				nbNotams++;
			}
			conn.Close();

			Lbl_location.Text        = AP;
			Lbl_notamsUnchecked.Text = "Notams Unchecked : " + nbNotams;
			Btn_submitNotams.Top     = Top + 30;
		}

		void ICAO_Notams()
		{
			tabPage2.VerticalScroll.Value = 0;
			ClearTaggedControls(tabPage2);

			string AP = TxtBox_ICAO.Text;

			OleDbConnection connOCC = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			connOCC.Open();
			OleDbCommand cmdOCC = new OleDbCommand("SELECT * FROM Stations_ICAO_IATA WHERE ICAO=?", connOCC);
			cmdOCC.Parameters.AddWithValue("?", AP);
			OleDbDataReader OCCreader = cmdOCC.ExecuteReader();
			string RWYs = "";
			while (OCCreader.Read())
				if (!OCCreader.IsDBNull(6)) RWYs = OCCreader.GetString(6);
			connOCC.Close();

			string stationHTML = "<html><body style=\"font-family:Courier New;font-size:12px;\">" +
				"<b><u style=\"color:MidnightBlue;font-size:14px;\">" + AP + "</u></b><br/>" +
				"<span style=\"color:DarkGreen;\">" + RWYs.Replace("/", "<br/>") + "</span>" +
				"</body></html>";
			Web_ICAONotams.DocumentText = stationHTML;

			FontFamily courier = new FontFamily("Courier New");
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			string query2 = ChckBox_SeeIgnored.Checked
				? "SELECT * FROM filteredNotams_table WHERE location=?"
				: "SELECT * FROM filteredNotams_table WHERE (Status='K') AND (location=?)";
			OleDbCommand cmd = new OleDbCommand(query2, conn);
			cmd.Parameters.AddWithValue("?", AP);
			OleDbDataReader dBreader = cmd.ExecuteReader();

			int Top = 100;
			Dictionary<int, Button>      keep_Buttons       = new Dictionary<int, Button>();
			Dictionary<int, RichTextBox> RchTxt_notam_text  = new Dictionary<int, RichTextBox>();
			Dictionary<int, CheckBox>    apt_CLSD_Chckbox   = new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_CATI_Chckbox   = new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_NILS_Chckbox   = new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_NOALTN_Chckbox = new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_FUEL_Chckbox   = new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_MISC_Chckbox   = new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_AIPSUP_Chckbox = new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_RWYCLSD_Chckbox= new Dictionary<int, CheckBox>();
			Dictionary<int, TextBox>     remark_Txtbox      = new Dictionary<int, TextBox>();
			Dictionary<int, Button>      remark_Buttons     = new Dictionary<int, Button>();

			while (dBreader.Read())
			{
				int    notam_ID   = !dBreader.IsDBNull(0)  ? dBreader.GetInt32(0)   : 0;
				string fromDate   = !dBreader.IsDBNull(5)  ? dBreader.GetString(5)  : "";
				string tillDate   = !dBreader.IsDBNull(6)  ? dBreader.GetString(6)  : "";
				string notam_text = !dBreader.IsDBNull(7)  ? dBreader.GetString(7)  : "";
				string notam_key  = !dBreader.IsDBNull(10) ? dBreader.GetString(10) : "";
				string Status     = !dBreader.IsDBNull(12) ? dBreader.GetString(12) : "";
				string Impact     = !dBreader.IsDBNull(13) ? dBreader.GetString(13) : "";
				string Remark     = !dBreader.IsDBNull(14) ? dBreader.GetString(14) : "";
				notam_text = notam_text.Replace("(char)39", "'");

				Color keyColor = HasImpact(Impact) ? ImpactColor(Impact) : Color.MidnightBlue;

				AddNotamLabel(tabPage2, courier, notam_key, Top, 210, 140, keyColor, true, 11f);
				AddNotamLabel(tabPage2, courier, FormatDate(fromDate), Top, 355, 190, Color.DimGray, false, 10f);
				AddNotamLabel(tabPage2, courier, FormatDate(tillDate), Top, 550, 190, Color.DimGray, false, 10f);

				int height = Math.Max(notam_text.Length / 50 * 20, 80);
				RchTxt_notam_text[notam_ID] = MakeNotamRichText(courier, notam_text, Top+20, 210, height, Status == "K");
				HighlightKeywords(RchTxt_notam_text[notam_ID]);
				tabPage2.Controls.Add(RchTxt_notam_text[notam_ID]);

				keep_Buttons[notam_ID] = MakeKeepButton(courier, Status, new Point(770, Top+20));
				int nid = notam_ID;
				if (Status == "")  keep_Buttons[notam_ID].Click += (s, e) => Keep_Notam(nid);
				if (Status == "K") keep_Buttons[notam_ID].Click += (s, e) => Ignore_Notam(nid);
				tabPage2.Controls.Add(keep_Buttons[notam_ID]);

				if (Status == "K")
				{
					AddImpactCheckboxes(tabPage2, notam_ID, Impact, Top, 770, 850, 940, 1030,
						apt_CLSD_Chckbox, apt_CATI_Chckbox, apt_NILS_Chckbox,
						apt_NOALTN_Chckbox, apt_FUEL_Chckbox, apt_MISC_Chckbox,
						apt_AIPSUP_Chckbox, apt_RWYCLSD_Chckbox);

					if (HasImpact(Impact))
					{
						remark_Txtbox[notam_ID]  = new TextBox { Tag="dispose", Top=Top+94, Left=770,  Size=new Size(250,24), Text=Remark };
						remark_Buttons[notam_ID] = new Button  { Tag="dispose", Top=Top+92, Left=1020, Size=new Size(40,24),  Text="OK" };
						int ri = notam_ID;
						remark_Buttons[notam_ID].Click += (s, e) => Remark_Notam(ri, remark_Txtbox[ri].Text);
						tabPage2.Controls.Add(remark_Txtbox[notam_ID]);
						tabPage2.Controls.Add(remark_Buttons[notam_ID]);
					}
				}
				Top = Top + height + 30;
			}
			conn.Close();
		}

		void Keep_Notam(int notam_ID)
		{
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand cmd = new OleDbCommand("UPDATE filteredNotams_table SET Status='K' WHERE ID=?", conn);
			cmd.Parameters.AddWithValue("?", notam_ID);
			cmd.ExecuteNonQuery();
			conn.Close();
			Filter_Notams();
			ICAO_Notams();
		}

		void Ignore_Notam(int notam_ID)
		{
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand cmd = new OleDbCommand("UPDATE filteredNotams_table SET Status='' WHERE ID=?", conn);
			cmd.Parameters.AddWithValue("?", notam_ID);
			cmd.ExecuteNonQuery();
			conn.Close();
			Filter_Notams();
			ICAO_Notams();
		}

		void Impact_Notam(int notam_ID, string I)
		{
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand qry = new OleDbCommand("SELECT * FROM filteredNotams_table WHERE ID=?", conn);
			qry.Parameters.AddWithValue("?", notam_ID);
			OleDbDataReader r = qry.ExecuteReader();
			string Impact = "";
			while (r.Read())
				if (!r.IsDBNull(13)) Impact = r.GetString(13);
			conn.Close();

			conn.Open();
			OleDbCommand upd;
			if (Impact == "")
			{
				upd = new OleDbCommand("UPDATE filteredNotams_table SET Impact=? WHERE ID=?", conn);
				upd.Parameters.AddWithValue("?", I);
				upd.Parameters.AddWithValue("?", notam_ID);
			}
			else
			{
				upd = new OleDbCommand("UPDATE filteredNotams_table SET Impact='' WHERE ID=?", conn);
				upd.Parameters.AddWithValue("?", notam_ID);
			}
			upd.ExecuteNonQuery();
			conn.Close();
			ICAO_Notams();
			Filter_Notams();
		}

		void Remark_Notam(int notam_ID, string remark)
		{
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand cmd = new OleDbCommand("UPDATE filteredNotams_table SET Remark=? WHERE ID=?", conn);
			cmd.Parameters.AddWithValue("?", remark);
			cmd.Parameters.AddWithValue("?", notam_ID);
			cmd.ExecuteNonQuery();
			conn.Close();
			ICAO_Notams();
		}

		void Btn_submitNotamsClick(object sender, EventArgs e)
		{
			string AP = Lbl_location.Text;
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand cmd = new OleDbCommand(
				"UPDATE filteredNotams_table SET Checked='Y' WHERE (location=?) AND (Checked='N')", conn);
			cmd.Parameters.AddWithValue("?", AP);
			cmd.ExecuteNonQuery();
			conn.Close();
			Filter_Notams();
		}

		void Btn_ICAOClick(object sender, EventArgs e)                        { ICAO_Notams(); }
		void ChckBox_SeeIgnoredCheckedChanged(object sender, EventArgs e)     { ICAO_Notams(); }

		public void ShowAutoPopup(string message, int durationMs = 1200)
		{
			Form popup = new Form
			{
				StartPosition    = FormStartPosition.CenterScreen,
				FormBorderStyle  = FormBorderStyle.FixedToolWindow,
				Width = 350, Height = 120, TopMost = true, ControlBox = false
			};
			Label lbl = new Label
			{
				Dock = DockStyle.Fill,
				TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
				Text = message, Font = new Font("Segoe UI", 10)
			};
			popup.Controls.Add(lbl);
			popup.Shown += (s, e) =>
			{
				var timer = new Timer { Interval = durationMs };
				timer.Tick += (s2, e2) => { timer.Stop(); popup.Close(); };
				timer.Start();
			};
			popup.Show();
			popup.Refresh();
		}

		// ── private helpers ──────────────────────────────────────────────────

		private string FormatDate(string raw)
		{
			if (raw.Length < 16) return raw;
			string d = raw.Substring(0, 16);
			return d.Substring(8,2) + " " + MonthAbbrev(d.Substring(5,2)) + " " + d.Substring(0,4) + " " + d.Substring(11,5) + "z";
		}

		private void AddNotamLabel(Control parent, FontFamily font, string text, int top, int left, int width,
			Color color, bool bold = false, float size = 10f)
		{
			Label lbl = new Label
			{
				Font      = new Font(font, size, bold ? FontStyle.Bold : FontStyle.Regular),
				Tag       = "dispose",
				Top       = top, Left = left,
				Size      = new Size(width, 18),
				ForeColor = color, Text = text
			};
			parent.Controls.Add(lbl);
		}

		private RichTextBox MakeNotamRichText(FontFamily font, string text, int top, int left, int height, bool kept)
		{
			return new RichTextBox
			{
				Font      = new Font(font, 10, FontStyle.Regular), Tag = "dispose",
				Top       = top, Left = left, Size = new Size(550, height),
				ForeColor = Color.Black,
				BackColor = kept ? Color.FromArgb(220, 230, 255) : SystemColors.Window,
				Text      = text, ReadOnly = true
			};
		}

		private Button MakeKeepButton(FontFamily font, string status, Point location)
		{
			return new Button
			{
				Tag       = "dispose", Location = location,
				Size      = status == "K" ? new Size(50,25) : new Size(40,25),
				Text      = status == "K" ? "Ignore" : "Keep",
				BackColor = status == "K" ? Color.LightCoral : Color.CornflowerBlue,
				Font      = new Font(font, 7)
			};
		}

		private bool HasImpact(string impact)
		{
			return impact == "A" || impact == "C" || impact == "N" || impact == "D" ||
			       impact == "F" || impact == "M" || impact == "AS" || impact == "R";
		}

		private void AddImpactCheckboxes(Control parent, int notam_ID, string Impact, int Top,
			int col1, int col2, int col3, int col4,
			Dictionary<int,CheckBox> CLSD,   Dictionary<int,CheckBox> CATI,
			Dictionary<int,CheckBox> NILS,   Dictionary<int,CheckBox> NOALTN,
			Dictionary<int,CheckBox> FUEL,   Dictionary<int,CheckBox> MISC,
			Dictionary<int,CheckBox> AIPSUP, Dictionary<int,CheckBox> RWYCLSD)
		{
			int nid = notam_ID;
			CLSD[notam_ID]    = AddImpactChk(parent, "APT CLSD", Top+44, col1, Impact=="A",  (s,e) => Impact_Notam(nid,"A"));
			CATI[notam_ID]    = AddImpactChk(parent, "APT CATI", Top+44, col2, Impact=="C",  (s,e) => Impact_Notam(nid,"C"));
			NILS[notam_ID]    = AddImpactChk(parent, "No ILS",   Top+44, col3, Impact=="N",  (s,e) => Impact_Notam(nid,"N"));
			NOALTN[notam_ID]  = AddImpactChk(parent, "Not ALTN", Top+68, col1, Impact=="D",  (s,e) => Impact_Notam(nid,"D"));
			FUEL[notam_ID]    = AddImpactChk(parent, "Fuel",     Top+68, col2, Impact=="F",  (s,e) => Impact_Notam(nid,"F"));
			MISC[notam_ID]    = AddImpactChk(parent, "MISC",     Top+68, col3, Impact=="M",  (s,e) => Impact_Notam(nid,"M"));
			AIPSUP[notam_ID]  = AddImpactChk(parent, "SUP",      Top+44, col4, Impact=="AS", (s,e) => Impact_Notam(nid,"AS"));
			RWYCLSD[notam_ID] = AddImpactChk(parent, "RWY",      Top+68, col4, Impact=="R",  (s,e) => Impact_Notam(nid,"R"));
		}

		private CheckBox AddImpactChk(Control parent, string text, int top, int left, bool chked, EventHandler handler)
		{
			CheckBox chk = new CheckBox { Tag="dispose", Top=top, Left=left, Text=text, Size=new Size(80,25), Checked=chked };
			chk.CheckedChanged += handler;
			parent.Controls.Add(chk);
			return chk;
		}
	}
}
