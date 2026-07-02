using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		private TextBox      _aptSearch;
		private DataGridView _aptDgv;

		// ── UI ───────────────────────────────────────────────────────────────
		public void Airport_List()
		{
			APT_List.VerticalScroll.Value = 0;
			ClearTaggedControls(APT_List);

			Label hdr = new Label { Tag = "dispose", Top = 18, Left = 20, AutoSize = true,
				Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold), Text = "Airports" };
			APT_List.Controls.Add(hdr);

			_aptSearch = new TextBox { Tag = "dispose", Top = 50, Left = 20, Width = 200 };
			APT_List.Controls.Add(_aptSearch);
			Label searchHint = new Label { Tag = "dispose", Top = 54, Left = 226, AutoSize = true,
				ForeColor = Color.Gray, Font = new Font("Microsoft Sans Serif", 8.5f), Text = "Search ICAO / IATA" };
			APT_List.Controls.Add(searchHint);
			_aptSearch.TextChanged += (s, e) => FilterAptGrid();

			Button bCopy = new Button { Tag = "dispose", Top = 49, Left = 380, Size = new Size(110, 26),
				Text = "Copy ICAO List", BackColor = Color.FromArgb(38, 50, 56), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
			bCopy.Click += (s, e) => CopyAptList();
			APT_List.Controls.Add(bCopy);

			_aptDgv = new DataGridView { Tag = "dispose", Top = 85, Left = 20, Size = new Size(700, 900),
				AllowUserToAddRows = true, RowHeadersWidth = 28, BackgroundColor = Color.White,
				AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None };
			_aptDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "ICAO", HeaderText = "ICAO", Width = 80 });
			_aptDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "IATA", HeaderText = "IATA", Width = 80 });
			_aptDgv.Columns.Add(new DataGridViewCheckBoxColumn { Name = "LH",       HeaderText = "Long Haul", Width = 90 });
			_aptDgv.Columns.Add(new DataGridViewCheckBoxColumn { Name = "FedEx",    HeaderText = "FedEx",     Width = 70 });
			_aptDgv.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Charters", HeaderText = "Charters",  Width = 80 });
			_aptDgv.Columns.Add(new DataGridViewButtonColumn { Name = "Del", HeaderText = "", Text = "Del",
				UseColumnTextForButtonValue = true, Width = 55, DefaultCellStyle = { BackColor = Color.MistyRose } });

			// Checkbox columns commit (and save) on click instead of waiting for focus to
			// leave the cell — matches the AIP SUP report's "click to toggle" feel.
			_aptDgv.CellContentClick += (s, e) =>
			{
				if (e.RowIndex < 0) return;
				if (_aptDgv.Columns[e.ColumnIndex].Name == "Del") { DeleteAptRow(e.RowIndex); return; }
				if (_aptDgv.Columns[e.ColumnIndex] is DataGridViewCheckBoxColumn) _aptDgv.EndEdit();
			};
			_aptDgv.CellValueChanged += (s, e) => { if (e.RowIndex >= 0) SaveAptRow(e.RowIndex); };

			APT_List.Controls.Add(_aptDgv);

			LoadAptGrid();
		}

		private void LoadAptGrid()
		{
			_aptDgv.Rows.Clear();
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			conn.Open();
			OleDbDataReader reader = new OleDbCommand("SELECT * FROM Stations_ICAO_IATA ORDER BY ICAO", conn).ExecuteReader();
			while (reader.Read())
			{
				int    id       = !reader.IsDBNull(0) ? reader.GetInt32(0)  : 0;
				string icao     = !reader.IsDBNull(1) ? reader.GetString(1) : "";
				string iata     = !reader.IsDBNull(2) ? reader.GetString(2) : "";
				string lh       = !reader.IsDBNull(3) ? reader.GetString(3) : "";
				string fedex    = !reader.IsDBNull(4) ? reader.GetString(4) : "";
				string charters = !reader.IsDBNull(5) ? reader.GetString(5) : "";

				int rowIndex = _aptDgv.Rows.Add(icao, iata, lh == "Yes", fedex == "Yes", charters == "Yes");
				_aptDgv.Rows[rowIndex].Tag = id;
			}
			conn.Close();
		}

		private void FilterAptGrid()
		{
			string term = (_aptSearch.Text ?? "").Trim().ToUpper();
			foreach (DataGridViewRow row in _aptDgv.Rows)
			{
				if (row.IsNewRow) { row.Visible = true; continue; }
				string icao = Cell(row, "ICAO"), iata = Cell(row, "IATA");
				row.Visible = term == "" || icao.ToUpper().Contains(term) || iata.ToUpper().Contains(term);
			}
		}

		private void CopyAptList()
		{
			List<string> icaos = new List<string>();
			foreach (DataGridViewRow row in _aptDgv.Rows)
				if (!row.IsNewRow && row.Visible && Cell(row, "ICAO") != "") icaos.Add(Cell(row, "ICAO"));
			if (icaos.Count > 0) Clipboard.SetText(string.Join(Environment.NewLine, icaos.ToArray()));
		}

		// Insert (Tag still null, ICAO now filled) or update (Tag holds the DB ID) a row
		// as soon as any of its cells commits — the grid's own "Copy List"/checkbox/text
		// edits all funnel through here instead of a separate Add/Edit form.
		private void SaveAptRow(int rowIndex)
		{
			DataGridViewRow row = _aptDgv.Rows[rowIndex];
			if (row.IsNewRow) return;

			string icao = Cell(row, "ICAO").Trim().ToUpper();
			if (icao == "") return;   // wait until the dispatcher has typed an ICAO
			string iata     = Cell(row, "IATA").Trim().ToUpper();
			string lh       = CheckedCell(row, "LH")       ? "Yes" : "No";
			string fedex    = CheckedCell(row, "FedEx")    ? "Yes" : "No";
			string charters = CheckedCell(row, "Charters") ? "Yes" : "No";

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			conn.Open();

			bool isNew = row.Tag == null;
			if (isNew)
			{
				OleDbCommand ins = new OleDbCommand(
					"INSERT INTO Stations_ICAO_IATA ([ICAO],[IATA],[LH],[FedEx],[Charters]) VALUES (?,?,?,?,?)", conn);
				ins.Parameters.AddWithValue("?", icao);
				ins.Parameters.AddWithValue("?", iata);
				ins.Parameters.AddWithValue("?", lh);
				ins.Parameters.AddWithValue("?", fedex);
				ins.Parameters.AddWithValue("?", charters);
				ins.ExecuteNonQuery();

				OleDbCommand idQuery = new OleDbCommand("SELECT @@IDENTITY", conn);
				row.Tag = Convert.ToInt32(idQuery.ExecuteScalar());
			}
			else
			{
				OleDbCommand upd = new OleDbCommand(
					"UPDATE Stations_ICAO_IATA SET ICAO=?,IATA=?,LH=?,FedEx=?,Charters=? WHERE ID=?", conn);
				upd.Parameters.AddWithValue("?", icao);
				upd.Parameters.AddWithValue("?", iata);
				upd.Parameters.AddWithValue("?", lh);
				upd.Parameters.AddWithValue("?", fedex);
				upd.Parameters.AddWithValue("?", charters);
				upd.Parameters.AddWithValue("?", (int)row.Tag);
				upd.ExecuteNonQuery();
			}
			conn.Close();

			// New airport -> pre-encode its runways from the CSV
			if (isNew)
			{
				ImportRunwaysFromCsv(icao);
				RegenerateRwyMemo(icao);
			}

			LoadStationsCache();
		}

		private void DeleteAptRow(int rowIndex)
		{
			DataGridViewRow row = _aptDgv.Rows[rowIndex];
			if (row.IsNewRow || row.Tag == null) return;

			string icao = Cell(row, "ICAO");
			if (MessageBox.Show("Delete " + icao + " ?", "Delete Airport", MessageBoxButtons.YesNo) != DialogResult.Yes)
				return;

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			conn.Open();
			OleDbCommand cmd = new OleDbCommand("DELETE From Stations_ICAO_IATA WHERE ID=?", conn);
			cmd.Parameters.AddWithValue("?", (int)row.Tag);
			cmd.ExecuteNonQuery();
			conn.Close();

			_aptDgv.Rows.Remove(row);
			LoadStationsCache();
		}

		// ── helpers ──────────────────────────────────────────────────────────

		private static string Cell(DataGridViewRow row, string columnName)
		{
			object v = row.Cells[columnName].Value;
			return v == null ? "" : v.ToString().Trim();
		}

		private static bool CheckedCell(DataGridViewRow row, string columnName)
		{
			object v = row.Cells[columnName].Value;
			return v != null && v is bool && (bool)v;
		}

		private void ClearTaggedControls(System.Windows.Forms.Control panel)
		{
			var labels   = new List<Label>();
			var txtboxes = new List<TextBox>();
			var richtxt  = new List<RichTextBox>();
			var chkboxes = new List<CheckBox>();
			var buttons  = new List<Button>();
			var panels   = new List<Panel>();
			var browsers = new List<WebBrowser>();
			var grids    = new List<DataGridView>();
			var combos   = new List<ComboBox>();

			foreach (Label       c in panel.Controls.OfType<Label>())       if (c.Tag != null && c.Tag.ToString() == "dispose") labels.Add(c);
			foreach (TextBox     c in panel.Controls.OfType<TextBox>())     if (c.Tag != null && c.Tag.ToString() == "dispose") txtboxes.Add(c);
			foreach (RichTextBox c in panel.Controls.OfType<RichTextBox>()) if (c.Tag != null && c.Tag.ToString() == "dispose") richtxt.Add(c);
			foreach (CheckBox    c in panel.Controls.OfType<CheckBox>())    if (c.Tag != null && c.Tag.ToString() == "dispose") chkboxes.Add(c);
			foreach (Button      c in panel.Controls.OfType<Button>())      if (c.Tag != null && c.Tag.ToString() == "dispose") buttons.Add(c);
			foreach (Panel       c in panel.Controls.OfType<Panel>())       if (c.Tag != null && c.Tag.ToString() == "dispose") panels.Add(c);
			foreach (WebBrowser  c in panel.Controls.OfType<WebBrowser>())  if (c.Tag != null && c.Tag.ToString() == "dispose") browsers.Add(c);
			foreach (DataGridView c in panel.Controls.OfType<DataGridView>()) if (c.Tag != null && c.Tag.ToString() == "dispose") grids.Add(c);
			foreach (ComboBox    c in panel.Controls.OfType<ComboBox>())    if (c.Tag != null && c.Tag.ToString() == "dispose") combos.Add(c);

			foreach (var c in labels)   { panel.Controls.Remove(c); c.Dispose(); }
			foreach (var c in txtboxes) { panel.Controls.Remove(c); c.Dispose(); }
			foreach (var c in richtxt)  { panel.Controls.Remove(c); c.Dispose(); }
			foreach (var c in chkboxes) { panel.Controls.Remove(c); c.Dispose(); }
			foreach (var c in buttons)  { panel.Controls.Remove(c); c.Dispose(); }
			foreach (var c in panels)   { panel.Controls.Remove(c); c.Dispose(); }
			foreach (var c in browsers) { panel.Controls.Remove(c); c.Dispose(); }
			foreach (var c in grids)    { panel.Controls.Remove(c); c.Dispose(); }
			foreach (var c in combos)   { panel.Controls.Remove(c); c.Dispose(); }
		}
	}
}
