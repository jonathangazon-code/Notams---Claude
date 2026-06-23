using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		void Filter_Notams()
		{
			tabPage1.VerticalScroll.Value = 0;
			ClearTaggedControls(tabPage1);

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();

			// Last unchecked airport
			string AP = "";
			OleDbDataReader APreader = new OleDbCommand(
				"SELECT location FROM filteredNotams_table WHERE (Checked='N') ORDER BY location DESC", conn).ExecuteReader();
			while (APreader.Read())
				if (!APreader.IsDBNull(0)) AP = APreader.GetString(0);
			conn.Close();

			// RWY info for this airport
			OleDbConnection connOCC = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			connOCC.Open();
			OleDbCommand cmdOCC = new OleDbCommand("SELECT * FROM Stations_ICAO_IATA WHERE ICAO=?", connOCC);
			cmdOCC.Parameters.AddWithValue("?", AP);
			OleDbDataReader OCCreader = cmdOCC.ExecuteReader();
			string RWYs = "";
			while (OCCreader.Read())
				if (!OCCreader.IsDBNull(6)) RWYs = OCCreader.GetString(6);
			connOCC.Close();

			string richText = AP + "\n";
			foreach (string rwy in RWYs.Split('/')) richText += rwy + "\n";
			richText += "_____________________________\n\n";

			// Previously kept NOTAMs
			conn.Open();
			OleDbCommand cmdKept = new OleDbCommand(
				"SELECT * FROM filteredNotams_table WHERE (Status='K') AND (location=?)", conn);
			cmdKept.Parameters.AddWithValue("?", AP);
			OleDbDataReader keptReader = cmdKept.ExecuteReader();
			while (keptReader.Read())
			{
				string fromDate  = !keptReader.IsDBNull(5)  ? keptReader.GetString(5)  : "";
				string tillDate  = !keptReader.IsDBNull(6)  ? keptReader.GetString(6)  : "";
				string text      = !keptReader.IsDBNull(7)  ? keptReader.GetString(7)  : "";
				string notamKey  = !keptReader.IsDBNull(10) ? keptReader.GetString(10) : "";
				string Impact    = !keptReader.IsDBNull(13) ? keptReader.GetString(13) : "";
				string Remark    = !keptReader.IsDBNull(14) ? keptReader.GetString(14) : "";
				text = text.Replace("(char)39", "'");
				richText += notamKey + "\n" + FormatDate(fromDate) + " - " + FormatDate(tillDate) + "\n" + text + "\n" + Impact + ": " + Remark + "\n\n";
			}
			conn.Close();
			RchTxt_FilterNotams.Text = richText;

			// New unchecked NOTAMs
			conn.Open();
			OleDbCommand cmdNew = new OleDbCommand(
				"SELECT * FROM filteredNotams_table WHERE (Checked='N') AND (location=?)", conn);
			cmdNew.Parameters.AddWithValue("?", AP);
			OleDbDataReader dBreader = cmdNew.ExecuteReader();

			int nbNotams = 0;
			int Top = 100;

			Dictionary<int, Button>      keep_Buttons      = new Dictionary<int, Button>();
			Dictionary<int, RichTextBox> RchTxt_notam_text = new Dictionary<int, RichTextBox>();
			Dictionary<int, CheckBox>    apt_CLSD_Chckbox  = new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_CATI_Chckbox  = new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_NILS_Chckbox  = new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_NOALTN_Chckbox= new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_FUEL_Chckbox  = new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_MISC_Chckbox  = new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_AIPSUP_Chckbox= new Dictionary<int, CheckBox>();
			Dictionary<int, CheckBox>    apt_RWYCLSD_Chckbox=new Dictionary<int, CheckBox>();
			Dictionary<int, TextBox>     remark_Txtbox     = new Dictionary<int, TextBox>();
			Dictionary<int, Button>      remark_Buttons    = new Dictionary<int, Button>();

			while (dBreader.Read())
			{
				int    notam_ID  = !dBreader.IsDBNull(0)  ? dBreader.GetInt32(0)  : 0;
				string fromDate  = !dBreader.IsDBNull(5)  ? dBreader.GetString(5)  : "";
				string tillDate  = !dBreader.IsDBNull(6)  ? dBreader.GetString(6)  : "";
				string notam_text= !dBreader.IsDBNull(7)  ? dBreader.GetString(7)  : "";
				string notam_key = !dBreader.IsDBNull(10) ? dBreader.GetString(10) : "";
				string Status    = !dBreader.IsDBNull(12) ? dBreader.GetString(12) : "";
				string Impact    = !dBreader.IsDBNull(13) ? dBreader.GetString(13) : "";
				string Remark    = !dBreader.IsDBNull(14) ? dBreader.GetString(14) : "";
				notam_text = notam_text.Replace("(char)39", "'");

				FontFamily courier = new FontFamily("Courier New");

				AddNotamLabel(tabPage1, courier, notam_key, Top, 510, 125);
				AddNotamLabel(tabPage1, courier, "From : " + FormatDate(fromDate), Top, 635, 200);
				AddNotamLabel(tabPage1, courier, "Till : " + FormatDate(tillDate), Top, 840, 200);

				int height = Math.Max(notam_text.Length / 50 * 20, 80);
				RchTxt_notam_text[notam_ID] = MakeNotamRichText(courier, notam_text, Top+20, 510, height, Status == "K");
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
						remark_Txtbox[notam_ID] = new TextBox { Tag="dispose", Top=Top+94, Left=1070, Size=new Size(250,24), Text=Remark };
						remark_Buttons[notam_ID] = new Button  { Tag="dispose", Top=Top+92, Left=1320, Size=new Size(40,24), Text="OK" };
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

			Lbl_location.Text = AP;
			Lbl_notamsUnchecked.Text = "Notams Unchecked : " + nbNotams;
			Btn_submitNotams.Top = Top + 30;
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

			string stationHTML = "<p style=\"font:Courier New;\"><b><u>" + AP + "</u></b><br />" + RWYs + "</p>";
			Web_ICAONotams.DocumentText = stationHTML;

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
				int    notam_ID   = !dBreader.IsDBNull(0)  ? dBreader.GetInt32(0)  : 0;
				string fromDate   = !dBreader.IsDBNull(5)  ? dBreader.GetString(5)  : "";
				string tillDate   = !dBreader.IsDBNull(6)  ? dBreader.GetString(6)  : "";
				string notam_text = !dBreader.IsDBNull(7)  ? dBreader.GetString(7)  : "";
				string notam_key  = !dBreader.IsDBNull(10) ? dBreader.GetString(10) : "";
				string Status     = !dBreader.IsDBNull(12) ? dBreader.GetString(12) : "";
				string Impact     = !dBreader.IsDBNull(13) ? dBreader.GetString(13) : "";
				string Remark     = !dBreader.IsDBNull(14) ? dBreader.GetString(14) : "";
				notam_text = notam_text.Replace("(char)39", "'");

				FontFamily courier = new FontFamily("Courier New");

				AddNotamLabel(tabPage2, courier, notam_key, Top, 210, 125);
				AddNotamLabel(tabPage2, courier, "From : " + FormatDate(fromDate), Top, 335, 200);
				AddNotamLabel(tabPage2, courier, "Till : " + FormatDate(tillDate), Top, 540, 200);

				int height = Math.Max(notam_text.Length / 50 * 20, 80);
				RchTxt_notam_text[notam_ID] = MakeNotamRichText(courier, notam_text, Top+20, 210, height, Status == "K");
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
						remark_Txtbox[notam_ID] = new TextBox { Tag="dispose", Top=Top+94, Left=770, Size=new Size(250,24), Text=Remark };
						remark_Buttons[notam_ID] = new Button  { Tag="dispose", Top=Top+92, Left=1020, Size=new Size(40,24), Text="OK" };
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
			OleDbDataReader reader = new OleDbCommand(
				"SELECT * FROM filteredNotams_table WHERE ID=?", conn) { }.ExecuteReader();

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
			OleDbCommand cmd = new OleDbCommand("UPDATE filteredNotams_table SET Checked='Y' WHERE (location=?) AND (Checked='N')", conn);
			cmd.Parameters.AddWithValue("?", AP);
			cmd.ExecuteNonQuery();
			conn.Close();
			Filter_Notams();
		}

		void Btn_ICAOClick(object sender, EventArgs e)            { ICAO_Notams(); }
		void ChckBox_SeeIgnoredCheckedChanged(object sender, EventArgs e) { ICAO_Notams(); }

		public void ShowAutoPopup(string message, int durationMs = 1200)
		{
			Form popup = new Form
			{
				StartPosition = FormStartPosition.CenterScreen,
				FormBorderStyle = FormBorderStyle.FixedToolWindow,
				Width = 350, Height = 120, TopMost = true, ControlBox = false
			};
			Label lbl = new Label { Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleCenter, Text = message, Font = new Font("Segoe UI", 10) };
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
			return d.Substring(8,2) + MonthAbbrev(d.Substring(5,2)) + d.Substring(0,4) + "(" + d.Substring(11,5) + ")";
		}

		private void AddNotamLabel(Control parent, FontFamily font, string text, int top, int left, int width)
		{
			Label lbl = new Label { Font = new Font(font, 10, FontStyle.Regular), Tag = "dispose", Top = top, Left = left, Size = new Size(width, 16), ForeColor = Color.Black, Text = text };
			parent.Controls.Add(lbl);
		}

		private RichTextBox MakeNotamRichText(FontFamily font, string text, int top, int left, int height, bool kept)
		{
			return new RichTextBox
			{
				Font = new Font(font, 10, FontStyle.Regular), Tag = "dispose",
				Top = top, Left = left, Size = new Size(550, height),
				ForeColor = Color.Black,
				BackColor = kept ? Color.CornflowerBlue : Color.LightCoral,
				Text = text, ReadOnly = true
			};
		}

		private Button MakeKeepButton(FontFamily font, string status, Point location)
		{
			return new Button
			{
				Tag = "dispose", Location = location,
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
			Dictionary<int,CheckBox> CLSD, Dictionary<int,CheckBox> CATI,
			Dictionary<int,CheckBox> NILS, Dictionary<int,CheckBox> NOALTN,
			Dictionary<int,CheckBox> FUEL, Dictionary<int,CheckBox> MISC,
			Dictionary<int,CheckBox> AIPSUP, Dictionary<int,CheckBox> RWYCLSD)
		{
			int nid = notam_ID;
			CLSD[notam_ID]   = AddImpactChk(parent, "APT CLSD", Top+44, col1, Impact=="A",  (s,e) => Impact_Notam(nid,"A"));
			CATI[notam_ID]   = AddImpactChk(parent, "APT CATI", Top+44, col2, Impact=="C",  (s,e) => Impact_Notam(nid,"C"));
			NILS[notam_ID]   = AddImpactChk(parent, "No ILS",   Top+44, col3, Impact=="N",  (s,e) => Impact_Notam(nid,"N"));
			NOALTN[notam_ID] = AddImpactChk(parent, "Not ALTN", Top+68, col1, Impact=="D",  (s,e) => Impact_Notam(nid,"D"));
			FUEL[notam_ID]   = AddImpactChk(parent, "Fuel",     Top+68, col2, Impact=="F",  (s,e) => Impact_Notam(nid,"F"));
			MISC[notam_ID]   = AddImpactChk(parent, "MISC",     Top+68, col3, Impact=="M",  (s,e) => Impact_Notam(nid,"M"));
			AIPSUP[notam_ID] = AddImpactChk(parent, "SUP",      Top+44, col4, Impact=="AS", (s,e) => Impact_Notam(nid,"AS"));
			RWYCLSD[notam_ID]= AddImpactChk(parent, "RWY",      Top+68, col4, Impact=="R",  (s,e) => Impact_Notam(nid,"R"));
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
