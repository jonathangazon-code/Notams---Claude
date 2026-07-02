using System;
using System.IO;
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

				if (IsOpsType("LH", ICAO) == "Yes")
				{
					if (isNext24) AppendImpactRow(Impact, loc, key, from, till, Remark, "Yellow", lhColor,
						ref t_APClsdLH, ref t_RWYClsdLH, ref t_CatILH, ref t_NilsLH, ref t_NoAltnLH, ref t_FuelLH, ref t_MiscLH);
					else AppendImpactRow(Impact, loc, key, from, till, Remark, "", lhColor,
						ref APClsdLH, ref RWYClsdLH, ref CatILH, ref NilsLH, ref NoAltnLH, ref FuelLH, ref MiscLH);
				}
				if (IsOpsType("FedEx", ICAO) == "Yes")
				{
					if (isNext24) AppendImpactRow(Impact, loc, key, from, till, Remark, "Yellow", shColor,
						ref t_APClsdSH, ref t_RWYClsdSH, ref t_CatISH, ref t_NilsSH, ref t_NoAltnSH, ref t_FuelSH, ref t_MiscSH);
					else AppendImpactRow(Impact, loc, key, from, till, Remark, "", shColor,
						ref APClsdSH, ref RWYClsdSH, ref CatISH, ref NilsSH, ref NoAltnSH, ref FuelSH, ref MiscSH);
				}
				if (IsOpsType("Charters", ICAO) == "Yes")
				{
					if (isNext24) AppendImpactRow(Impact, loc, key, from, till, Remark, "Yellow", chColor,
						ref t_APClsdC, ref t_RWYClsdC, ref t_CatIC, ref t_NilsC, ref t_NoAltnC, ref t_FuelC, ref t_MiscC);
					else AppendImpactRow(Impact, loc, key, from, till, Remark, "", chColor,
						ref APClsdC, ref RWYClsdC, ref CatIC, ref NilsC, ref NoAltnC, ref FuelC, ref MiscC);
				}
			}
			conn.Close();

			string window = radBtn_noFilter.Checked ? "COMPLETE"
				: radBtn_24Hrs.Checked   ? "Next 24 Hours"
				: radBtn_7days.Checked   ? "Next 7 days"
				: "Next 31 days";

			string reportDate = DateTime.Now.ToString("ddMMMMyyyy HHmm") + "CET";
			string report = "<html><head><title>NOTAM REPORT</title><body style=\"font-family:Calibri\">" +
				"<h1>Notam Report - " + window + "</h1><p>" + reportDate + "</p>" +
				"<table border=\"1\" style=\"width:700px;text-align:left;font-family:Calibri;font-size:12px;border:1px solid black;border-collapse:collapse;\">" +
				Section("RoyalBlue",    "Long Haul",  t_APClsdLH,t_RWYClsdLH,t_CatILH,t_NilsLH,t_NoAltnLH,t_FuelLH,t_MiscLH, APClsdLH,RWYClsdLH,CatILH,NilsLH,NoAltnLH,FuelLH,MiscLH) +
				Section("#663399","Short Haul",  t_APClsdSH,t_RWYClsdSH,t_CatISH,t_NilsSH,t_NoAltnSH,t_FuelSH,t_MiscSH, APClsdSH,RWYClsdSH,CatISH,NilsSH,NoAltnSH,FuelSH,MiscSH) +
				Section("SeaGreen",     "Charters",    t_APClsdC, t_RWYClsdC, t_CatIC, t_NilsC, t_NoAltnC, t_FuelC, t_MiscC,  APClsdC, RWYClsdC, CatIC, NilsC, NoAltnC, FuelC, MiscC) +
				"</table></body></html>";

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
				SubSection("Not as Altn",                 t_D + D) +
				SubSection("Fuel",                        t_F + F) +
				SubSection("Misc",                        t_M + M);
		}

		private static string SubSection(string label, string rows)
		{
			return "<tr><th colspan=\"4\" bgcolor=\"LightSlateGrey\" style=\"color:white;\"><b>" + label + "</b></th></tr>" + rows;
		}

		private static void AppendImpactRow(string impact, string loc, string key, string from, string till, string remark,
			string bg, string color,
			ref string A, ref string R, ref string C, ref string N, ref string D, ref string F, ref string M)
		{
			string bgAttr = bg != "" ? " bgcolor=\"" + bg + "\"" : "";
			string row = "<tr><th" + bgAttr + " style=\"color:" + color + ";\">" + loc + "</th>" +
				"<th style=\"font-family:Courier New;\">" + key + "</th>" +
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
	}
}
