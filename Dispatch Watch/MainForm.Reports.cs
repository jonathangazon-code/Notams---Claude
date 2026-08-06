using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		// The report auto-loads when the tab is selected (TabPage.Enter fires whenever this
		// page becomes the active tab), and re-runs whenever the dispatcher changes the
		// time-window filter — there's no "Report !" button anymore.
		void TabPage4Enter(object sender, EventArgs e) { Report(); }

		void RadBtn_reportWindowCheckedChanged(object sender, EventArgs e)
		{
			if (((RadioButton)sender).Checked) Report();
		}

		public void Report()
		{
			// Impact rows per category per operator
			string APClsdLH="",RWYClsdLH="",CatILH="",NilsLH="",NoAltnLH="",FuelLH="",MiscLH="";
			string APClsdSH="",RWYClsdSH="",CatISH="",NilsSH="",NoAltnSH="",FuelSH="",MiscSH="";
			string APClsdC ="",RWYClsdC ="",CatIC ="",NilsC ="",NoAltnC ="",FuelC ="",MiscC ="";
			string t_APClsdLH="",t_RWYClsdLH="",t_CatILH="",t_NilsLH="",t_NoAltnLH="",t_FuelLH="",t_MiscLH="";
			string t_APClsdSH="",t_RWYClsdSH="",t_CatISH="",t_NilsSH="",t_NoAltnSH="",t_FuelSH="",t_MiscSH="";
			string t_APClsdC ="",t_RWYClsdC ="",t_CatIC ="",t_NilsC ="",t_NoAltnC ="",t_FuelC ="",t_MiscC ="";

			DateTime today         = DateTime.Now;
			string todayStr        = today.ToString("yyyyMMdd");
			string tomorrowStr     = today.AddDays(1).ToString("yyyyMMdd");
			string sevenDaysStr    = today.AddDays(7).ToString("yyyyMMdd");
			string thirtyOneDaysStr= today.AddDays(31).ToString("yyyyMMdd");
			int tomorrowInt = Int32.Parse(tomorrowStr);
			int todayInt = 0, endWindowInt = 30000101;

			if (radBtn_24Hrs.Checked)   { todayInt = Int32.Parse(todayStr); endWindowInt = Int32.Parse(tomorrowStr); }
			if (radBtn_7days.Checked)   { todayInt = Int32.Parse(todayStr); endWindowInt = Int32.Parse(sevenDaysStr); }
			if (radBtn_31days.Checked)  { todayInt = Int32.Parse(todayStr); endWindowInt = Int32.Parse(thirtyOneDaysStr); }

			// Airports referenced by at least one summary row, in first-seen order — each gets
			// a detail section (airport card + Kept NOTAMs) appended below the summary table,
			// and the row's NOTAM key links to it.
			HashSet<string> seenIcaos = new HashSet<string>();
			List<string> orderedIcaos = new List<string>();
			HashSet<string> knownImpactCodes = new HashSet<string> { "A", "R", "C", "N", "D", "F", "M" };

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbDataReader dBreader = new OleDbCommand("SELECT * FROM filteredNotams_table ORDER BY LOCATION", conn).ExecuteReader();

			while (dBreader.Read())
			{
				if (dBreader.IsDBNull(12) || dBreader.GetString(12) != "K") continue;

				string fromDate = !dBreader.IsDBNull(5) ? dBreader.GetString(5) : "";
				string tillDate = !dBreader.IsDBNull(6) ? dBreader.GetString(6) : "";
				int fromInt = Int32.Parse(fromDate.Substring(0,4) + fromDate.Substring(5,2) + fromDate.Substring(8,2));
				int tillInt = Int32.Parse(tillDate.Substring(0,4) + tillDate.Substring(5,2) + tillDate.Substring(8,2));

				if (!(tillInt > todayInt && fromInt < endWindowInt)) continue;

				string ICAO    = !dBreader.IsDBNull(8)  ? dBreader.GetString(8)  : "";
				string key     = !dBreader.IsDBNull(10) ? dBreader.GetString(10) : "";
				string Impact  = !dBreader.IsDBNull(13) ? dBreader.GetString(13) : "";
				string Remark  = !dBreader.IsDBNull(14) ? dBreader.GetString(14) : "";
				string loc     = GetIATA(ICAO);
				string from    = dateTransformation(fromDate);
				string till    = dateTransformation(tillDate);
				bool   isNext24 = (fromInt <= tomorrowInt);

				string lhColor  = "RoyalBlue";
				string shColor  = "#663399";
				string chColor  = "SeaGreen";

				bool anyOp = IsOpsType("LH", ICAO) == "Yes" || IsOpsType("FedEx", ICAO) == "Yes" || IsOpsType("Charters", ICAO) == "Yes";
				if (anyOp && ICAO != "" && knownImpactCodes.Contains(Impact) && seenIcaos.Add(ICAO))
					orderedIcaos.Add(ICAO);

				if (IsOpsType("LH", ICAO) == "Yes")
				{
					if (isNext24) AppendImpactRow(Impact, loc, ICAO, key, from, till, Remark, "Yellow", lhColor,
						ref t_APClsdLH, ref t_RWYClsdLH, ref t_CatILH, ref t_NilsLH, ref t_NoAltnLH, ref t_FuelLH, ref t_MiscLH);
					else AppendImpactRow(Impact, loc, ICAO, key, from, till, Remark, "", lhColor,
						ref APClsdLH, ref RWYClsdLH, ref CatILH, ref NilsLH, ref NoAltnLH, ref FuelLH, ref MiscLH);
				}
				if (IsOpsType("FedEx", ICAO) == "Yes")
				{
					if (isNext24) AppendImpactRow(Impact, loc, ICAO, key, from, till, Remark, "Yellow", shColor,
						ref t_APClsdSH, ref t_RWYClsdSH, ref t_CatISH, ref t_NilsSH, ref t_NoAltnSH, ref t_FuelSH, ref t_MiscSH);
					else AppendImpactRow(Impact, loc, ICAO, key, from, till, Remark, "", shColor,
						ref APClsdSH, ref RWYClsdSH, ref CatISH, ref NilsSH, ref NoAltnSH, ref FuelSH, ref MiscSH);
				}
				if (IsOpsType("Charters", ICAO) == "Yes")
				{
					if (isNext24) AppendImpactRow(Impact, loc, ICAO, key, from, till, Remark, "Yellow", chColor,
						ref t_APClsdC, ref t_RWYClsdC, ref t_CatIC, ref t_NilsC, ref t_NoAltnC, ref t_FuelC, ref t_MiscC);
					else AppendImpactRow(Impact, loc, ICAO, key, from, till, Remark, "", chColor,
						ref APClsdC, ref RWYClsdC, ref CatIC, ref NilsC, ref NoAltnC, ref FuelC, ref MiscC);
				}
			}
			conn.Close();

			string window = radBtn_noFilter.Checked ? "COMPLETE"
				: radBtn_24Hrs.Checked   ? "Next 24 Hours"
				: radBtn_7days.Checked   ? "Next 7 days"
				: "Next 31 days";

			string reportDate = DateTime.Now.ToString("ddMMMMyyyy HHmm") + "CET";
			string report = "<html><head><title>NOTAM REPORT</title><style>" +
				".notamlink{color:#003399;text-decoration:underline}" +
				BuildAirportDetailStyleBlock() +
				"</style><body style=\"font-family:Calibri\">" +
				"<h1>Notam Report - " + window + "</h1><p>" + reportDate + "</p>" +
				"<table border=\"1\" style=\"width:700px;text-align:left;font-family:Calibri;font-size:12px;border:1px solid black;border-collapse:collapse;\">" +
				Section("RoyalBlue",    "Long Haul",  t_APClsdLH,t_RWYClsdLH,t_CatILH,t_NilsLH,t_NoAltnLH,t_FuelLH,t_MiscLH, APClsdLH,RWYClsdLH,CatILH,NilsLH,NoAltnLH,FuelLH,MiscLH) +
				Section("#663399","Short Haul",  t_APClsdSH,t_RWYClsdSH,t_CatISH,t_NilsSH,t_NoAltnSH,t_FuelSH,t_MiscSH, APClsdSH,RWYClsdSH,CatISH,NilsSH,NoAltnSH,FuelSH,MiscSH) +
				Section("SeaGreen",     "Charters",    t_APClsdC, t_RWYClsdC, t_CatIC, t_NilsC, t_NoAltnC, t_FuelC, t_MiscC,  APClsdC, RWYClsdC, CatIC, NilsC, NoAltnC, FuelC, MiscC) +
				"</table>" +
				BuildAirportDetailSectionsHtml(orderedIcaos) +
				"</body></html>";

			// Save report to OCC.mdb (SQL concatenation kept here as the column names are dynamic fields)
			OleDbConnection conn2 = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			conn2.Open();
			string update = "UPDATE Notams_ICAO_CSV SET " +
				"APClsdLH='" + APClsdLH + "',RWYClsdLH='" + RWYClsdLH + "',CatILH='" + CatILH + "'," +
				"NilsLH='" + NilsLH + "',NoAltnLH='" + NoAltnLH + "',FuelLH='" + FuelLH + "',MiscLH='" + MiscLH + "'," +
				"APClsdSH='" + APClsdSH + "',RWYClsdSH='" + RWYClsdSH + "',CatISH='" + CatISH + "'," +
				"NilsSH='" + NilsSH + "',NoAltnSH='" + NoAltnSH + "',FuelSH='" + FuelSH + "',MiscSH='" + MiscSH + "'," +
				"APClsdCharters='" + APClsdC + "',RWYClsdCharters='" + RWYClsdC + "',CatICharters='" + CatIC + "'," +
				"NilsCharters='" + NilsC + "',NoAltnCharters='" + NoAltnC + "',FuelCharters='" + FuelC + "',MiscCharters='" + MiscC + "' WHERE ID=1";
			new OleDbCommand(update, conn2).ExecuteNonQuery();
			conn2.Close();

			Web_report.DocumentText = report;
			string dir = @"V:\TAY Ops Control Centre\Flight Dispatch\AIP SUP -  Notams report\Reports";
			Directory.CreateDirectory(dir);
			File.WriteAllText(System.IO.Path.Combine(dir, DateTime.Now.ToString("yyMMdd") + "-Complete_NOTAMS_report.html"), report);
		}

		private static string Section(string color, string title,
			string t_A, string t_R, string t_C, string t_N, string t_D, string t_F, string t_M,
			string A,   string R,   string C,   string N,   string D,   string F,   string M)
		{
			string hdr = "<tr><th colspan=\"4\" bgcolor=\"" + color + "\" style=\"font-size:16px;color:white;font-weight:bold;\">" + title + ":</th></tr>";
			return hdr +
				SubSection("AP Closed",                   t_A + A) +
				SubSection("RWY/TWY Closure impacting Perfos", t_R + R) +
				SubSection("CAT I",                       t_C + C) +
				SubSection("No ILS",                      t_N + N) +
				SubSection("Fuel",                        t_F + F) +
				SubSection("Not as Altn",                 t_D + D) +
				SubSection("Misc",                        t_M + M);
		}

		private static string SubSection(string label, string rows)
		{
			return "<tr><th colspan=\"4\" bgcolor=\"LightSlateGrey\" style=\"color:white;\"><b>" + label + "</b></th></tr>" + rows;
		}

		private static void AppendImpactRow(string impact, string loc, string icao, string key, string from, string till, string remark,
			string bg, string color,
			ref string A, ref string R, ref string C, ref string N, ref string D, ref string F, ref string M)
		{
			string bgAttr = bg != "" ? " bgcolor=\"" + bg + "\"" : "";
			string keyCell = key != ""
				? "<a href=\"#ap-" + icao + "\" class=\"notamlink\">" + key + "</a>"
				: key;
			string row = "<tr><th" + bgAttr + " style=\"color:" + color + ";\">" + loc + "</th>" +
				"<th style=\"font-family:Courier New;\">" + keyCell + "</th>" +
				"<th style=\"font-family:Courier New;\">" + from + "-" + till + "</th>" +
				"<th" + bgAttr + " style=\"width:400px;font-weight:normal;\">" + remark + "</th></tr>";
			if (impact == "A") A += row;
			if (impact == "R") R += row;
			if (impact == "C") C += row;
			if (impact == "N") N += row;
			if (impact == "D") D += row;
			if (impact == "F") F += row;
			if (impact == "M") M += row;
		}

		// Reuses the same dark-card visual language as the Conflict tab (MainForm.Conflict.cs)
		// so a NOTAM key link's destination looks consistent with the rest of the app.
		private static string BuildAirportDetailStyleBlock()
		{
			return
				".apSection{margin:24px 0 0 0}" +
				".ahead{background:#263238;padding:14px 18px;position:relative;border-radius:8px 8px 0 0}" +
				".icao{font-size:18px;font-weight:bold;color:#eceff1;letter-spacing:3px}" +
				".sub{font-size:13px;color:#78909c;margin-top:2px}" +
				".apname{font-size:13px;color:#90a4ae;margin-top:1px}" +
				".blk{font-size:12px;color:#b0bec5;background:#37474f;border-left:2px solid #546e7a;padding:6px 12px;margin-right:8px;vertical-align:top}" +
				".rwytable{margin-top:8px}" +
				".rwyline{white-space:nowrap;line-height:1.7}" +
				".diagram{position:absolute;top:10px;right:60px}" +
				".keptWrap{border:1px solid #cfd8dc;border-top:none;border-radius:0 0 8px 8px;padding:12px 18px;background:#fafafa}" +
				".keptCard{position:relative;border:1px solid #e0e0e0;border-radius:4px;background:#fff;padding:8px 10px 8px 14px;margin:0 0 10px 0}" +
				".keptStrip{position:absolute;left:0;top:0;bottom:0;width:4px;border-radius:4px 0 0 4px}" +
				".keptKey{font-family:'Courier New',monospace;font-weight:bold;font-size:12.5px}" +
				".keptDates{font-size:11.5px;color:#607d8b;margin:2px 0 4px 0}" +
				".keptText{font-family:'Courier New',monospace;font-size:12px;white-space:pre-wrap;line-height:1.5;color:#222}" +
				".keptRemark{font-size:12px;margin-top:4px}" +
				".noKept{font-size:12.5px;color:#78909c;font-style:italic}";
		}

		// One <div class="apSection"> per airport referenced by a NOTAM key link in the
		// summary table above — the airport header (ICAO/IATA/name/RWY/diagram, shared with
		// the Conflict tab via BuildAirportHeaderHtml) followed by every Kept NOTAM at that
		// station, styled like RenderKeptCard's WinForms layout in the NOTAM Filter tab.
		private string BuildAirportDetailSectionsHtml(List<string> icaos)
		{
			StringBuilder sb = new StringBuilder();
			foreach (string icao in icaos)
			{
				sb.Append("<div id=\"ap-" + icao + "\" class=\"apSection\">");
				sb.Append(BuildAirportHeaderHtml(icao));
				sb.Append("<div class=\"keptWrap\">");

				List<string> keptCards = new List<string>();
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				OleDbCommand cmd = new OleDbCommand("SELECT * FROM filteredNotams_table WHERE (Status='K') AND (location=?)", conn);
				cmd.Parameters.AddWithValue("?", icao);
				OleDbDataReader r = cmd.ExecuteReader();
				while (r.Read())
				{
					string fromDate = !r.IsDBNull(5)  ? r.GetString(5)  : "";
					string tillDate = !r.IsDBNull(6)  ? r.GetString(6)  : "";
					string text     = !r.IsDBNull(7)  ? r.GetString(7).Replace((char)39, '\'') : "";
					string key      = !r.IsDBNull(10) ? r.GetString(10) : "";
					string impact   = !r.IsDBNull(13) ? r.GetString(13) : "";
					string remark   = !r.IsDBNull(14) ? r.GetString(14) : "";
					keptCards.Add(BuildKeptNotamCardHtml(key, fromDate, tillDate, text, impact, remark));
				}
				conn.Close();

				if (keptCards.Count == 0) sb.Append("<div class=\"noKept\">No Kept NOTAMs at this station.</div>");
				else foreach (string card in keptCards) sb.Append(card);

				sb.Append("</div></div>");
			}
			return sb.ToString();
		}

		private string BuildKeptNotamCardHtml(string notamKey, string fromDate, string tillDate, string text, string impact, string remark)
		{
			string hex = System.Drawing.ColorTranslator.ToHtml(ImpactColor(impact));
			string label = ImpactLabel(impact);
			string labelSuffix = label != "" ? "  [" + label + "]" : "";
			string from = dateTransformation(fromDate);
			string till = dateTransformation(tillDate);
			string remarkLine = remark != ""
				? "<div class=\"keptRemark\" style=\"color:" + hex + "\">&#9654; " + remark.Replace("&", "&amp;").Replace("<", "&lt;") + "</div>"
				: "";

			return
				"<div class=\"keptCard\">" +
				"<div class=\"keptStrip\" style=\"background:" + hex + "\"></div>" +
				"<div class=\"keptKey\" style=\"color:" + hex + "\">" + notamKey.Replace("&", "&amp;").Replace("<", "&lt;") + labelSuffix + "</div>" +
				"<div class=\"keptDates\">" + from + "  &#8594;  " + till + "</div>" +
				"<div class=\"keptText\">" + HighlightKeywordsHtml(text) + "</div>" +
				remarkLine +
				"</div>";
		}
	}
}
