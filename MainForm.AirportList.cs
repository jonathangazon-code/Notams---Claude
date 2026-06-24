using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		public void Airport_List()
		{
			ClearTaggedControls(APT_List);
			APT_List.AutoScroll = true;

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			conn.Open();
			OleDbDataReader reader = new OleDbCommand("Select * From Stations_ICAO_IATA ORDER BY ICAO", conn).ExecuteReader();

			TxtBox_APT_ICAO.Text     = "";
			TxtBox_APT_IATA.Text     = "";
			ChckBx_APT_LH.Checked   = false;
			ChckBx_APT_FedEx.Checked = false;
			ChckBx_APT_Charters.Checked = false;

			Dictionary<int, Button> del_Buttons  = new Dictionary<int, Button>();
			Dictionary<int, Button> edit_Buttons = new Dictionary<int, Button>();
			int i = 0;
			Top = 70;

			while (reader.Read())
			{
				int    id       = !reader.IsDBNull(0) ? reader.GetInt32(0)  : 0;
				string icao     = !reader.IsDBNull(1) ? reader.GetString(1) : "";
				string iata     = !reader.IsDBNull(2) ? reader.GetString(2) : "";
				string lh       = !reader.IsDBNull(3) ? reader.GetString(3) : "";
				string fedex    = !reader.IsDBNull(4) ? reader.GetString(4) : "";
				string charters = !reader.IsDBNull(5) ? reader.GetString(5) : "";

				FontFamily family = new FontFamily("Courier New");

				AddDisposableLabel(APT_List, family, icao,       Top + 20*i, 28,  45, Color.OrangeRed);
				AddDisposableLabel(APT_List, family, " - "+iata, Top + 20*i, 65,  65, Color.CornflowerBlue);

				AddDisposableCheckBox(APT_List, Top + 20*i, 150, lh       == "Yes");
				AddDisposableCheckBox(APT_List, Top + 20*i, 190, fedex    == "Yes");
				AddDisposableCheckBox(APT_List, Top + 20*i, 230, charters == "Yes");

				int capturedId = id;
				del_Buttons[id] = MakeSmallButton("Del", Color.Red,       new System.Drawing.Point(310, Top-3+20*i));
				del_Buttons[id].Click += (s, e) => Delete_APT(capturedId);
				APT_List.Controls.Add(del_Buttons[id]);

				edit_Buttons[id] = MakeSmallButton("Edit", Color.LightBlue, new System.Drawing.Point(270, Top-3+20*i));
				edit_Buttons[id].Click += (s, e) => Edit_APT(capturedId);
				APT_List.Controls.Add(edit_Buttons[id]);

				i++;
			}
			conn.Close();
		}

		void Delete_APT(int i)
		{
			if (MessageBox.Show("Are you sure that you want to delete ?", "Delete Airport", MessageBoxButtons.YesNo) != System.Windows.Forms.DialogResult.Yes)
				return;

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			conn.Open();
			OleDbCommand cmd = new OleDbCommand("DELETE From Stations_ICAO_IATA WHERE ID=?", conn);
			cmd.Parameters.AddWithValue("?", i);
			cmd.ExecuteNonQuery();
			conn.Close();
			LoadStationsCache();
			Airport_List();
		}

		void Edit_APT(int i)
		{
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			conn.Open();
			OleDbCommand cmd = new OleDbCommand("Select * From Stations_ICAO_IATA WHERE ID=?", conn);
			cmd.Parameters.AddWithValue("?", i);
			OleDbDataReader reader = cmd.ExecuteReader();

			while (reader.Read())
			{
				TxtBox_APT_ICAO.Text        = !reader.IsDBNull(1) ? reader.GetString(1) : "";
				TxtBox_APT_IATA.Text        = !reader.IsDBNull(2) ? reader.GetString(2) : "";
				ChckBx_APT_LH.Checked      = !reader.IsDBNull(3) && reader.GetString(3) == "Yes";
				ChckBx_APT_FedEx.Checked   = !reader.IsDBNull(4) && reader.GetString(4) == "Yes";
				ChckBx_APT_Charters.Checked = !reader.IsDBNull(5) && reader.GetString(5) == "Yes";
			}
			conn.Close();

			Btn_addAPT.Text = "Edit";
			Btn_addAPT.Tag  = i.ToString();
		}

		void Btn_addAPTClick(object sender, System.EventArgs e)
		{
			string icao     = TxtBox_APT_ICAO.Text;
			string iata     = TxtBox_APT_IATA.Text;
			string lh       = ChckBx_APT_LH.Checked       ? "Yes" : "No";
			string fedex    = ChckBx_APT_FedEx.Checked     ? "Yes" : "No";
			string charters = ChckBx_APT_Charters.Checked  ? "Yes" : "No";

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			conn.Open();

			if (Btn_addAPT.Text == "Edit")
			{
				Btn_addAPT.Text = "Add Airport !";
				int intID = int.Parse(Btn_addAPT.Tag.ToString());
				OleDbCommand cmd = new OleDbCommand("UPDATE Stations_ICAO_IATA SET ICAO=?,IATA=?,LH=?,FedEx=?,Charters=? WHERE ID=?", conn);
				cmd.Parameters.AddWithValue("?", icao);
				cmd.Parameters.AddWithValue("?", iata);
				cmd.Parameters.AddWithValue("?", lh);
				cmd.Parameters.AddWithValue("?", fedex);
				cmd.Parameters.AddWithValue("?", charters);
				cmd.Parameters.AddWithValue("?", intID);
				cmd.ExecuteNonQuery();
			}
			else
			{
				OleDbCommand cmd = new OleDbCommand("INSERT INTO Stations_ICAO_IATA ([ICAO],[IATA],[LH],[FedEx],[Charters]) VALUES (?,?,?,?,?)", conn);
				cmd.Parameters.AddWithValue("?", icao);
				cmd.Parameters.AddWithValue("?", iata);
				cmd.Parameters.AddWithValue("?", lh);
				cmd.Parameters.AddWithValue("?", fedex);
				cmd.Parameters.AddWithValue("?", charters);
				cmd.ExecuteNonQuery();
			}

			conn.Close();
			LoadStationsCache();
			Airport_List();
		}

		// ── helpers ──────────────────────────────────────────────────────────

		private void ClearTaggedControls(System.Windows.Forms.Control panel)
		{
			var labels   = new List<Label>();
			var txtboxes = new List<TextBox>();
			var richtxt  = new List<RichTextBox>();
			var chkboxes = new List<CheckBox>();
			var buttons  = new List<Button>();
			var panels   = new List<Panel>();

			foreach (Label       c in panel.Controls.OfType<Label>())       if (c.Tag != null && c.Tag.ToString() == "dispose") labels.Add(c);
			foreach (TextBox     c in panel.Controls.OfType<TextBox>())     if (c.Tag != null && c.Tag.ToString() == "dispose") txtboxes.Add(c);
			foreach (RichTextBox c in panel.Controls.OfType<RichTextBox>()) if (c.Tag != null && c.Tag.ToString() == "dispose") richtxt.Add(c);
			foreach (CheckBox    c in panel.Controls.OfType<CheckBox>())    if (c.Tag != null && c.Tag.ToString() == "dispose") chkboxes.Add(c);
			foreach (Button      c in panel.Controls.OfType<Button>())      if (c.Tag != null && c.Tag.ToString() == "dispose") buttons.Add(c);
			foreach (Panel       c in panel.Controls.OfType<Panel>())       if (c.Tag != null && c.Tag.ToString() == "dispose") panels.Add(c);

			foreach (var c in labels)   { panel.Controls.Remove(c); c.Dispose(); }
			foreach (var c in txtboxes) { panel.Controls.Remove(c); c.Dispose(); }
			foreach (var c in richtxt)  { panel.Controls.Remove(c); c.Dispose(); }
			foreach (var c in chkboxes) { panel.Controls.Remove(c); c.Dispose(); }
			foreach (var c in buttons)  { panel.Controls.Remove(c); c.Dispose(); }
			foreach (var c in panels)   { panel.Controls.Remove(c); c.Dispose(); }
		}

		private void AddDisposableLabel(System.Windows.Forms.Control parent, FontFamily family, string text, int top, int left, int width, System.Drawing.Color color)
		{
			Label lbl = new Label { Font = new Font(family, 11.0f, FontStyle.Bold), Tag = "dispose", Top = top, Left = left, Size = new Size(width, 16), ForeColor = color, Text = text };
			parent.Controls.Add(lbl);
		}

		private void AddDisposableCheckBox(System.Windows.Forms.Control parent, int top, int left, bool chked)
		{
			CheckBox chk = new CheckBox { Enabled = false, Tag = "dispose", Top = top, Left = left, Size = new Size(20, 16), ForeColor = Color.DimGray, Checked = chked };
			parent.Controls.Add(chk);
		}

		private Button MakeSmallButton(string text, System.Drawing.Color color, System.Drawing.Point location)
		{
			return new Button { Tag = "dispose", Size = new Size(35, 20), Location = location, Text = text, BackColor = color, Font = new Font(System.Drawing.SystemFonts.DefaultFont.FontFamily, 7) };
		}
	}
}
