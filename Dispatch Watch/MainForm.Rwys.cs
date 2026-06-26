using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		void tab_RWYs()
		{
			tabPage5.VerticalScroll.Value = 0;
			ClearTaggedControls(tabPage5);

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			conn.Open();
			OleDbDataReader dBreader = new OleDbCommand("SELECT * FROM Stations_ICAO_IATA ORDER BY ICAO", conn).ExecuteReader();

			Dictionary<int, Button>      update_Buttons = new Dictionary<int, Button>();
			Dictionary<int, RichTextBox> RchTxt_RWYs   = new Dictionary<int, RichTextBox>();
			FontFamily courier = new FontFamily("Courier New");
			int Top = 100;

			while (dBreader.Read())
			{
				int    AP_ID = !dBreader.IsDBNull(0) ? dBreader.GetInt32(0)  : 0;
				string ICAO  = !dBreader.IsDBNull(1) ? dBreader.GetString(1) : "";
				string RWYs  = !dBreader.IsDBNull(6) ? dBreader.GetString(6) : "";

				Label lbl = new Label { Font = new Font(courier, 10, FontStyle.Regular), Tag = "dispose", Top = Top, Left = 210, Size = new Size(125, 16), ForeColor = Color.Black, Text = ICAO };
				tabPage5.Controls.Add(lbl);

				RchTxt_RWYs[AP_ID] = new RichTextBox
				{
					Font = new Font(courier, 10, FontStyle.Regular), Tag = "dispose",
					Top = Top+20, Left = 210, Size = new Size(550, 100),
					ForeColor = Color.Black, BackColor = System.Drawing.Color.White,
					Text = RWYs, ReadOnly = true
				};
				tabPage5.Controls.Add(RchTxt_RWYs[AP_ID]);

				update_Buttons[AP_ID] = new Button { Tag = "dispose", Size = new Size(40,25), Location = new Point(770, Top+20), Text = "Edit", Font = new Font(courier, 7) };
				int id = AP_ID;
				update_Buttons[AP_ID].Click += (s, e) => Update_RWYs(id);
				tabPage5.Controls.Add(update_Buttons[AP_ID]);

				Top = Top + 130;
				RWYs = "";
			}
			conn.Close();
		}

		void Update_RWYs(int AP_ID)
		{
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			conn.Open();
			OleDbCommand cmd = new OleDbCommand("SELECT * FROM Stations_ICAO_IATA WHERE ID=?", conn);
			cmd.Parameters.AddWithValue("?", AP_ID);
			OleDbDataReader reader = cmd.ExecuteReader();
			while (reader.Read())
			{
				Lbl_ICAO_RWYs.Text    = !reader.IsDBNull(1) ? reader.GetString(1) : "";
				RchTxt_updateRWYs.Text = !reader.IsDBNull(6) ? reader.GetString(6) : "";
			}
			conn.Close();
			tab_RWYs();
		}

		void Btn_updateRWysClick(object sender, System.EventArgs e)
		{
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			conn.Open();
			OleDbCommand cmd = new OleDbCommand("UPDATE Stations_ICAO_IATA SET RWYs=? WHERE ICAO=?", conn);
			cmd.Parameters.AddWithValue("?", RchTxt_updateRWYs.Text);
			cmd.Parameters.AddWithValue("?", Lbl_ICAO_RWYs.Text);
			cmd.ExecuteNonQuery();
			conn.Close();
			tab_RWYs();
		}
	}
}
