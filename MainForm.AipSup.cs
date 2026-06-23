using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		void update_SUP_Avio(int i)
		{
			RchTxtBox_Test.Text = "le num ID est : " + i;

			string status = "";
			OleDbConnection conn2 = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn2.Open();

			OleDbCommand cmd = new OleDbCommand("SELECT * FROM filteredNotams_table WHERE ID=?", conn2);
			cmd.Parameters.AddWithValue("?", i);
			OleDbDataReader reader = cmd.ExecuteReader();
			while (reader.Read())
				if (!reader.IsDBNull(15)) status = reader.GetString(15);
			conn2.Close();

			string newStatus = (status == "Yes") ? "" : "Yes";
			RchTxtBox_Test.Text += " " + newStatus;

			conn2.Open();
			OleDbCommand upd = new OleDbCommand("UPDATE filteredNotams_table SET Loaded_Aviobook=? WHERE ID=?", conn2);
			upd.Parameters.AddWithValue("?", newStatus);
			upd.Parameters.AddWithValue("?", i);
			upd.ExecuteNonQuery();
			conn2.Close();

			AIP_SUP_Checklist();
		}

		public void AIP_SUP_Checklist()
		{
			foreach (var ctrl in new System.Windows.Forms.Control.ControlCollection(AIP_Sup))
			{
			}
			// Clear tagged controls
			var labelsToRemove = new List<Label>();
			foreach (Label lbl in AIP_Sup.Controls.OfType<Label>())
				if (lbl.Tag != null && lbl.Tag.ToString() == "dispose") labelsToRemove.Add(lbl);
			foreach (Label lbl in labelsToRemove) { AIP_Sup.Controls.Remove(lbl); lbl.Dispose(); }

			var chkToRemove = new List<CheckBox>();
			foreach (CheckBox chk in AIP_Sup.Controls.OfType<CheckBox>())
				if (chk.Tag != null && chk.Tag.ToString() == "dispose") chkToRemove.Add(chk);
			foreach (CheckBox chk in chkToRemove) { AIP_Sup.Controls.Remove(chk); chk.Dispose(); }

			AIP_Sup.AutoScroll = true;
			AIP_Sup.VerticalScroll.Value = 0;
			int baseTop = 10;
			int i = 0;

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbDataReader dBreader = new OleDbCommand("SELECT * FROM filteredNotams_table ORDER BY location", conn).ExecuteReader();

			Dictionary<int, CheckBox> ChckBx_Loaded_Aviobook = new Dictionary<int, CheckBox>();

			if (dBreader.HasRows)
			{
				while (dBreader.Read())
				{
					if (dBreader.IsDBNull(12)) continue;
					string Status = dBreader.GetString(12);
					if (Status != "K") continue;

					string Impact = !dBreader.IsDBNull(13) ? dBreader.GetString(13) : "";
					if (Impact != "AS") continue;

					string fromDate      = !dBreader.IsDBNull(5)  ? dBreader.GetString(5)  : "";
					string tillDate      = !dBreader.IsDBNull(6)  ? dBreader.GetString(6)  : "";
					int    int_APT_ID    = !dBreader.IsDBNull(0)  ? dBreader.GetInt32(0)   : 0;
					string ICAO          = !dBreader.IsDBNull(8)  ? dBreader.GetString(8)  : "";
					string key           = !dBreader.IsDBNull(10) ? dBreader.GetString(10) : "";
					string Remark        = !dBreader.IsDBNull(14) ? dBreader.GetString(14) : "";
					string loaded_avio   = !dBreader.IsDBNull(15) ? dBreader.GetString(15) : "";

					string location = GetIATA(ICAO);
					string fromTxt  = dateTransformation(fromDate);
					string tillTxt  = dateTransformation(tillDate);

					FontFamily family = new FontFamily("Courier New");

					AddLabel(AIP_Sup, family, location, baseTop + 20 * i, 28,  45,  Color.BlueViolet);
					AddLabel(AIP_Sup, family, key,      baseTop + 20 * i, 80,  130, Color.Black);
					AddLabel(AIP_Sup, family, fromTxt + "-" + tillTxt, baseTop + 20 * i, 230, 200, Color.Black);
					AddLabel(AIP_Sup, family, Remark,   baseTop + 20 * i, 450, 300, Color.Black);

					ChckBx_Loaded_Aviobook[int_APT_ID] = new CheckBox();
					ChckBx_Loaded_Aviobook[int_APT_ID].Tag      = "dispose";
					ChckBx_Loaded_Aviobook[int_APT_ID].Top      = baseTop + 20 * i;
					ChckBx_Loaded_Aviobook[int_APT_ID].Size     = new Size(20, 16);
					ChckBx_Loaded_Aviobook[int_APT_ID].Left     = 430;
					ChckBx_Loaded_Aviobook[int_APT_ID].Checked  = (loaded_avio == "Yes");
					ChckBx_Loaded_Aviobook[int_APT_ID].BackColor = (loaded_avio == "Yes") ? Color.DimGray : Color.Red;
					int id = int_APT_ID;
					ChckBx_Loaded_Aviobook[int_APT_ID].Click += (s, e) => update_SUP_Avio(id);
					AIP_Sup.Controls.Add(ChckBx_Loaded_Aviobook[int_APT_ID]);
					i++;
				}
			}
			conn.Close();
		}

		private static void AddLabel(System.Windows.Forms.Control parent, FontFamily family, string text, int top, int left, int width, System.Drawing.Color color)
		{
			Label lbl = new Label();
			lbl.Font      = new Font(family, 11.0f, FontStyle.Bold);
			lbl.Tag       = "dispose";
			lbl.Top       = top;
			lbl.Left      = left;
			lbl.Size      = new Size(width, 16);
			lbl.ForeColor = color;
			lbl.Text      = text;
			parent.Controls.Add(lbl);
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

					string Impact  = !dBreader.IsDBNull(13) ? dBreader.GetString(13) : "";
					if (Impact != "AS") continue;

					string ICAO        = !dBreader.IsDBNull(8)  ? dBreader.GetString(8)  : "";
					string key         = !dBreader.IsDBNull(10) ? dBreader.GetString(10) : "";
					string Remark      = !dBreader.IsDBNull(14) ? dBreader.GetString(14) : "";
					string loaded_avio = !dBreader.IsDBNull(15) ? dBreader.GetString(15) : "";
					string location    = GetIATA(ICAO);
					string fromTxt     = dateTransformation(fromDate);
					string tillTxt     = dateTransformation(tillDate);

					string chk_status = (loaded_avio == "Yes")
						? "<th style=\"width:30px\"><input type=\"checkbox\" checked></th>"
						: "<th bgcolor=\"Red\" style=\"width:30px\"><input type=\"checkbox\"></th>";

					string row = "<tr><th style=\"color:SaddleBrown;\">" + location + "</th>" +
						"<th style=\"width:100px;font-family:Courier New;\">" + key + "</th>" +
						"<th style=\"width:140px;font-family:Courier New;padding-right:10px;\">" + fromTxt + "-" + tillTxt + "</th>" +
						chk_status + "<th style=\"font-weight:normal;\">" + Remark + "</th></tr>";

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
			string report = "<html><head><title>AIP SUP Listing</title><body style=\"font-family:Calibri\">" +
				"<h1>AIP SUP Listing - " + window + "</h1>" +
				"<p>" + reportDate + "</p>" +
				"<table border=\"1\" style=\"width:700px;text-align:left;font-family:Calibri;font-size:12px;border:1px solid black;border-collapse:collapse\">" +
				"<tr bgcolor=\"Black\" style=\"color:white;font-size:14px;\"><th>IATA</th><th>Notam Ref</th><th>From-Till</th><th>Avio</th><th>AIP SUP Ref</th></tr>" +
				twenty4H_AIP_Sup_list + AIP_Sup_list +
				"</table></body></html>";

			Web_Sup_report.DocumentText = report;

			string directoryPath = @"V:\TAY Ops Control Centre\Flight Dispatch\AIP SUP -  Notams report\Reports";
			Directory.CreateDirectory(directoryPath);
			File.WriteAllText(System.IO.Path.Combine(directoryPath, DateTime.Now.ToString("yyMMdd") + "-AIP_SUP_report.html"), report);
		}
	}
}
