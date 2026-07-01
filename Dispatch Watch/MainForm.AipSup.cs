using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ICAO_CSV
{
	[ComVisible(true)]
	public class AviobookScriptBridge
	{
		private MainForm _form;
		public AviobookScriptBridge(MainForm form) { _form = form; }
		public void UpdateAviobook(int id) { _form.update_SUP_Avio(id); }
	}

	public partial class MainForm
	{
		public void update_SUP_Avio(int i)
		{
			string status = "";
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand cmd = new OleDbCommand("SELECT * FROM filteredNotams_table WHERE ID=?", conn);
			cmd.Parameters.AddWithValue("?", i);
			OleDbDataReader reader = cmd.ExecuteReader();
			while (reader.Read())
				if (!reader.IsDBNull(15)) status = reader.GetString(15);
			conn.Close();

			string newStatus = (status == "Yes") ? "" : "Yes";

			conn.Open();
			OleDbCommand upd = new OleDbCommand("UPDATE filteredNotams_table SET Loaded_Aviobook=? WHERE ID=?", conn);
			upd.Parameters.AddWithValue("?", newStatus);
			upd.Parameters.AddWithValue("?", i);
			upd.ExecuteNonQuery();
			conn.Close();

			// Update only the checkbox in the HTML — no full redraw
			if (Web_Sup_report.Document != null)
				Web_Sup_report.Document.InvokeScript("updateCheckbox",
					new object[] { i, newStatus == "Yes" });
		}

		public void AIP_SUP_Checklist() { } // replaced by HTML checkboxes in Sup_Report()

		// The report auto-loads when the tab is selected, and re-runs whenever the
		// dispatcher changes the time-window filter — there's no "Report !" button anymore.
		void RadBtn_supReportWindowCheckedChanged(object sender, EventArgs e)
		{
			if (((RadioButton)sender).Checked) Sup_Report();
		}

		public void Sup_Report()
		{
			string AIP_Sup_list       = "";
			string twenty4H_AIP_Sup_list = "";

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbDataReader dBreader = new OleDbCommand("SELECT * FROM filteredNotams_table ORDER BY location", conn).ExecuteReader();

			DateTime today         = DateTime.Now;
			DateTime tomorrow      = today.AddDays(1);
			DateTime sevenDays     = today.AddDays(7);
			DateTime thirtyOneDays = today.AddDays(31);

			string todayString        = today.ToString("yyyyMMdd");
			string tomorrowString     = tomorrow.ToString("yyyyMMdd");
			string sevenDaysString    = sevenDays.ToString("yyyyMMdd");
			string thirtyOneDaysString= thirtyOneDays.ToString("yyyyMMdd");
			int tomorrowInt = Int32.Parse(tomorrowString);

			int todayInt    = 0;
			int endWindowInt = 30000101;
			if (radBtn_Sup_24Hrs.Checked)  { todayInt = Int32.Parse(todayString); endWindowInt = Int32.Parse(tomorrowString); }
			if (radBtn_Sup_7days.Checked)  { todayInt = Int32.Parse(todayString); endWindowInt = Int32.Parse(sevenDaysString); }
			if (radBtn_Sup_31days.Checked) { todayInt = Int32.Parse(todayString); endWindowInt = Int32.Parse(thirtyOneDaysString); }

			if (dBreader.HasRows)
			{
				while (dBreader.Read())
				{
					if (dBreader.IsDBNull(12)) continue;
					if (dBreader.GetString(12) != "K") continue;

					string fromDate = !dBreader.IsDBNull(5) ? dBreader.GetString(5) : "";
					string tillDate = !dBreader.IsDBNull(6) ? dBreader.GetString(6) : "";

					string fromCheck = fromDate.Substring(0,4) + fromDate.Substring(5,2) + fromDate.Substring(8,2);
					string tillCheck = tillDate.Substring(0,4) + tillDate.Substring(5,2) + tillDate.Substring(8,2);
					int fromDateInt  = Int32.Parse(fromCheck);
					int tillDateInt  = Int32.Parse(tillCheck);

					if (!(tillDateInt > todayInt && fromDateInt < endWindowInt)) continue;

					int supOrd = dBreader.GetOrdinal("Sup");
					string Sup = !dBreader.IsDBNull(supOrd) ? dBreader.GetString(supOrd) : "";
					if (Sup != "Yes") continue;

					int    int_APT_ID  = !dBreader.IsDBNull(0)  ? dBreader.GetInt32(0)   : 0;
					string ICAO        = !dBreader.IsDBNull(8)  ? dBreader.GetString(8)  : "";
					string key         = !dBreader.IsDBNull(10) ? dBreader.GetString(10) : "";
					int    supRefOrd   = dBreader.GetOrdinal("SupRef");
					string SupRef      = !dBreader.IsDBNull(supRefOrd) ? dBreader.GetString(supRefOrd) : "";
					string Remark      = !dBreader.IsDBNull(14) ? dBreader.GetString(14) : "";
					string supText     = (SupRef != "") ? SupRef : Remark;
					string loaded_avio = !dBreader.IsDBNull(15) ? dBreader.GetString(15) : "";
					string location    = GetIATA(ICAO);
					string fromTxt     = dateTransformation(fromDate);
					string tillTxt     = dateTransformation(tillDate);

					string thBg  = (loaded_avio == "Yes") ? "" : " bgcolor=\"Red\"";
					string chkd  = (loaded_avio == "Yes") ? " checked" : "";
					string chk_status = "<th id=\"th_" + int_APT_ID + "\"" + thBg + " style=\"width:30px\">" +
						"<input type=\"checkbox\" id=\"chk_" + int_APT_ID + "\"" + chkd +
						" onclick=\"window.external.UpdateAviobook(" + int_APT_ID + ")\"></th>";

					string row = "<tr><th style=\"color:SaddleBrown;\">" + location + "</th>" +
						"<th style=\"width:100px;font-family:Courier New;\">" + key + "</th>" +
						"<th style=\"width:140px;font-family:Courier New;padding-right:10px;\">" + fromTxt + "-" + tillTxt + "</th>" +
						chk_status + "<th style=\"font-weight:normal;\">" + supText + "</th></tr>";

					if (fromDateInt <= tomorrowInt)
						twenty4H_AIP_Sup_list += "<tr><th bgcolor=\"Yellow\" " + row.Substring(4);
					else
						AIP_Sup_list += row;
				}
			}
			conn.Close();

			string window = radBtn_Sup_noFilter.Checked ? "COMPLETE"
				: radBtn_Sup_24Hrs.Checked  ? "Next 24 Hours"
				: radBtn_Sup_7days.Checked  ? "Next 7 days"
				: "Next 31 days";

			string reportDate = DateTime.Now.ToString("ddMMMMyyyy HHmm") + "CET";
			string js = "<script type=\"text/javascript\">" +
				"function updateCheckbox(id,isYes){" +
				"var chk=document.getElementById('chk_'+id);" +
				"var th=document.getElementById('th_'+id);" +
				"if(chk)chk.checked=isYes;" +
				"if(th)th.bgColor=isYes?'':'Red';}" +
				"</script>";
			string report = "<html><head><title>AIP SUP Listing</title>" + js + "</head>" +
				"<body style=\"font-family:Calibri\">" +
				"<h1>AIP SUP Listing - " + window + "</h1>" +
				"<p>" + reportDate + "</p>" +
				"<table border=\"1\" style=\"width:700px;text-align:left;font-family:Calibri;font-size:12px;border:1px solid black;border-collapse:collapse\">" +
				"<tr bgcolor=\"Black\" style=\"color:white;font-size:14px;\"><th>IATA</th><th>Notam Ref</th><th>From-Till</th><th>Avio</th><th>AIP SUP Ref</th></tr>" +
				twenty4H_AIP_Sup_list + AIP_Sup_list +
				"</table></body></html>";

			Web_Sup_report.ObjectForScripting = new AviobookScriptBridge(this);
			Web_Sup_report.DocumentText = report;

			string directoryPath = @"V:\TAY Ops Control Centre\Flight Dispatch\AIP SUP -  Notams report\Reports";
			Directory.CreateDirectory(directoryPath);
			File.WriteAllText(System.IO.Path.Combine(directoryPath, DateTime.Now.ToString("yyMMdd") + "-AIP_SUP_report.html"), report);
		}
	}
}
