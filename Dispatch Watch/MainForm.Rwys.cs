using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		private ComboBox     _rwyCmb;
		private DataGridView _rwyDgv;

		// ── Schema ───────────────────────────────────────────────────────────
		// Structured runway table in OCC.mdb (one row per threshold). The legacy
		// Stations_ICAO_IATA.RWYs memo is regenerated from it so ParseRunways /
		// BuildRwySvg / the airport card keep working unchanged.
		public void EnsureRunwaysTable()
		{
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
				conn.Open();
				try { new OleDbCommand("CREATE TABLE Runways ([ICAO] TEXT(4), [QFU] TEXT(8), [Cat] TEXT(30), [DistM] LONG, [Hdg] DOUBLE, [ThrLat] DOUBLE, [ThrLon] DOUBLE, [Ord] LONG)", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }
				conn.Close();
			}
			catch { }
		}

		// ── CSV helpers ──────────────────────────────────────────────────────
		private static string[] CsvSplit(string line)
		{
			List<string> fields = new List<string>();
			bool inQ = false;
			System.Text.StringBuilder cur = new System.Text.StringBuilder();
			for (int i = 0; i < line.Length; i++)
			{
				char ch = line[i];
				if (ch == '"') inQ = !inQ;
				else if (ch == ',' && !inQ) { fields.Add(cur.ToString()); cur.Length = 0; }
				else cur.Append(ch);
			}
			fields.Add(cur.ToString());
			return fields.ToArray();
		}

		private static double ParseD(string s)
		{
			double v;
			return Double.TryParse((s ?? "").Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out v) ? v : 0;
		}

		private static bool IsRunwayIdent(string s)
		{
			return Regex.IsMatch((s ?? "").Trim(), @"^\d{1,2}[LCRG]?$", RegexOptions.IgnoreCase);
		}

		// Insert one Runways row.
		private void InsertRwy(OleDbConnection conn, string icao, string qfu, string cat, int distM, double hdg, double lat, double lon, int ord)
		{
			OleDbCommand ins = new OleDbCommand(
				"INSERT INTO Runways ([ICAO],[QFU],[Cat],[DistM],[Hdg],[ThrLat],[ThrLon],[Ord]) VALUES (?,?,?,?,?,?,?,?)", conn);
			ins.Parameters.AddWithValue("?", icao);
			ins.Parameters.AddWithValue("?", qfu);
			ins.Parameters.AddWithValue("?", cat);
			ins.Parameters.AddWithValue("?", distM);
			ins.Parameters.AddWithValue("?", hdg);
			ins.Parameters.AddWithValue("?", lat);
			ins.Parameters.AddWithValue("?", lon);
			ins.Parameters.AddWithValue("?", ord);
			ins.ExecuteNonQuery();
		}

		private static double DefaultHdg(string qfu)
		{
			Match m = Regex.Match(qfu, @"\d{1,2}");
			return m.Success ? (Int32.Parse(m.Value) * 10) % 360 : 0;
		}

		// Single pass over runways.csv keeping only the wanted ICAOs -> ICAO -> CSV rows.
		private Dictionary<string, List<string[]>> ScanCsvFor(HashSet<string> wanted)
		{
			Dictionary<string, List<string[]>> d = new Dictionary<string, List<string[]>>(StringComparer.OrdinalIgnoreCase);
			string csv = Path.Combine(Application.StartupPath, "runways.csv");
			if (!File.Exists(csv)) return d;
			using (StreamReader sr = new StreamReader(csv))
			{
				string line;
				while ((line = sr.ReadLine()) != null)
				{
					string[] f = CsvSplit(line);
					if (f.Length < 19) continue;
					string id = f[2].Trim().ToUpper();
					if (id == "" || !wanted.Contains(id)) continue;
					if (!d.ContainsKey(id)) d[id] = new List<string[]>();
					d[id].Add(f);
				}
			}
			return d;
		}

		// Insert runway threshold rows from CSV field-arrays. Returns next Ord.
		private int ImportRows(OleDbConnection conn, string icao, List<string[]> rows, int ord)
		{
			if (rows == null) return ord;
			foreach (string[] f in rows)
			{
				if (f.Length < 19 || f[7].Trim() == "1") continue;   // closed
				int distM = (int)Math.Round(ParseD(f[3]) * 0.3048);
				string le = f[8].Trim().ToUpper(), he = f[14].Trim().ToUpper();
				if (IsRunwayIdent(le)) InsertRwy(conn, icao, le, "", distM, ParseD(f[12]) != 0 ? ParseD(f[12]) : DefaultHdg(le), ParseD(f[9]),  ParseD(f[10]), ord++);
				if (IsRunwayIdent(he)) InsertRwy(conn, icao, he, "", distM, ParseD(f[18]) != 0 ? ParseD(f[18]) : DefaultHdg(he), ParseD(f[15]), ParseD(f[16]), ord++);
			}
			return ord;
		}

		// Fill Hdg / ThrLat / ThrLon of existing Runways rows from CSV field-arrays (by QFU).
		private void EnrichRows(OleDbConnection conn, string icao, List<string[]> rows)
		{
			if (rows == null) return;
			Dictionary<string, double[]> byQfu = new Dictionary<string, double[]>();
			foreach (string[] f in rows)
			{
				if (f.Length < 19) continue;
				string le = f[8].Trim().ToUpper(), he = f[14].Trim().ToUpper();
				if (le != "") byQfu[le] = new double[] { ParseD(f[12]), ParseD(f[9]),  ParseD(f[10]) };
				if (he != "") byQfu[he] = new double[] { ParseD(f[18]), ParseD(f[15]), ParseD(f[16]) };
			}
			if (byQfu.Count == 0) return;
			List<string> qfus = new List<string>();
			OleDbCommand sel = new OleDbCommand("SELECT QFU FROM Runways WHERE ICAO=?", conn);
			sel.Parameters.AddWithValue("?", icao);
			OleDbDataReader r = sel.ExecuteReader();
			while (r.Read()) if (!r.IsDBNull(0)) qfus.Add(r.GetString(0).Trim().ToUpper());
			r.Close();
			foreach (string qfu in qfus)
			{
				if (!byQfu.ContainsKey(qfu)) continue;
				double[] d = byQfu[qfu];
				OleDbCommand cmd = new OleDbCommand("UPDATE Runways SET Hdg=?, ThrLat=?, ThrLon=? WHERE ICAO=? AND QFU=?", conn);
				cmd.Parameters.AddWithValue("?", d[0] != 0 ? d[0] : DefaultHdg(qfu));
				cmd.Parameters.AddWithValue("?", d[1]);
				cmd.Parameters.AddWithValue("?", d[2]);
				cmd.Parameters.AddWithValue("?", icao);
				cmd.Parameters.AddWithValue("?", qfu);
				cmd.ExecuteNonQuery();
			}
		}

		// Parse a legacy "QFU: CAT distM" memo into Runways rows (preserving the manual CAT).
		private void InsertMemoRows(OleDbConnection conn, string icao, string memo)
		{
			int ord = 0;
			foreach (string raw in memo.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n'))
			{
				string line = raw.Trim();
				if (line == "") continue;
				int colon = line.IndexOf(':');
				string qfu = (colon > 0 ? line.Substring(0, colon) : line).Trim().ToUpper();
				string rest = colon > 0 ? line.Substring(colon + 1) : "";
				int distM = 0; string cat = rest.Trim();
				Match dm = Regex.Match(rest, @"(\d+)\s*m\b", RegexOptions.IgnoreCase);
				if (dm.Success) { distM = Int32.Parse(dm.Groups[1].Value); cat = rest.Substring(0, dm.Index).Trim(); }
				InsertRwy(conn, icao, qfu, cat, distM, DefaultHdg(qfu), 0, 0, ord++);
			}
		}

		// Single-airport import (airport add / Re-import button).
		public int ImportRunwaysFromCsv(string icao)
		{
			icao = (icao ?? "").Trim().ToUpper();
			if (icao == "") return 0;
			HashSet<string> w = new HashSet<string>(StringComparer.OrdinalIgnoreCase); w.Add(icao);
			Dictionary<string, List<string[]>> csv = ScanCsvFor(w);
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			conn.Open();
			OleDbCommand del = new OleDbCommand("DELETE FROM Runways WHERE ICAO=?", conn);
			del.Parameters.AddWithValue("?", icao); del.ExecuteNonQuery();
			int ord = ImportRows(conn, icao, csv.ContainsKey(icao) ? csv[icao] : null, 0);
			conn.Close();
			return ord;
		}

		// Single-airport legacy migration (RwyLoadGrid fallback).
		private void MigrateLegacyMemo(string icao, string memo)
		{
			HashSet<string> w = new HashSet<string>(StringComparer.OrdinalIgnoreCase); w.Add(icao);
			Dictionary<string, List<string[]>> csv = ScanCsvFor(w);
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			conn.Open();
			OleDbCommand del = new OleDbCommand("DELETE FROM Runways WHERE ICAO=?", conn);
			del.Parameters.AddWithValue("?", icao); del.ExecuteNonQuery();
			InsertMemoRows(conn, icao, memo);
			EnrichRows(conn, icao, csv.ContainsKey(icao) ? csv[icao] : null);
			conn.Close();
		}

		// One-time bulk migration of every station (single CSV scan) — runs at startup
		// while the Runways table is still empty, so the diagram works everywhere.
		public void MigrateAllRunwaysIfNeeded()
		{
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
				conn.Open();
				int total = Convert.ToInt32(new OleDbCommand("SELECT COUNT(*) FROM Runways", conn).ExecuteScalar());
				if (total > 0) { conn.Close(); return; }

				List<string[]> stations = new List<string[]>();
				OleDbDataReader sr = new OleDbCommand("SELECT ICAO, RWYs FROM Stations_ICAO_IATA", conn).ExecuteReader();
				while (sr.Read())
					stations.Add(new string[] { sr.IsDBNull(0) ? "" : sr.GetString(0), sr.IsDBNull(1) ? "" : sr.GetString(1) });
				sr.Close();

				HashSet<string> wanted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
				foreach (string[] s in stations) if (s[0].Trim() != "") wanted.Add(s[0].Trim().ToUpper());
				Dictionary<string, List<string[]>> csv = ScanCsvFor(wanted);   // single pass

				foreach (string[] s in stations)
				{
					string icao = s[0].Trim().ToUpper();
					if (icao == "") continue;
					List<string[]> rows = csv.ContainsKey(icao) ? csv[icao] : null;
					if (s[1].Trim() != "") { InsertMemoRows(conn, icao, s[1]); EnrichRows(conn, icao, rows); }
					else                   { ImportRows(conn, icao, rows, 0); }
				}
				conn.Close();
			}
			catch { }
		}

		// Rebuild the legacy RWYs memo ("QFU: Cat DistMm") from the Runways table.
		public void RegenerateRwyMemo(string icao)
		{
			_geoCache.Remove(icao);   // runways changed -> drop cached diagram geometry
			System.Text.StringBuilder sb = new System.Text.StringBuilder();
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			conn.Open();
			OleDbCommand q = new OleDbCommand("SELECT QFU, Cat, DistM FROM Runways WHERE ICAO=? ORDER BY Ord", conn);
			q.Parameters.AddWithValue("?", icao);
			OleDbDataReader r = q.ExecuteReader();
			while (r.Read())
			{
				string qfu = !r.IsDBNull(0) ? r.GetString(0) : "";
				string cat = !r.IsDBNull(1) ? r.GetString(1).Trim() : "";
				int dist   = !r.IsDBNull(2) ? Convert.ToInt32(r.GetValue(2)) : 0;
				if (sb.Length > 0) sb.Append("\r\n");
				sb.Append(qfu).Append(": ").Append(cat != "" ? cat + " " : "").Append(dist).Append("m");
			}
			r.Close();
			OleDbCommand upd = new OleDbCommand("UPDATE Stations_ICAO_IATA SET RWYs=? WHERE ICAO=?", conn);
			upd.Parameters.AddWithValue("?", sb.ToString());
			upd.Parameters.AddWithValue("?", icao);
			upd.ExecuteNonQuery();
			conn.Close();
		}

		// ── UI ───────────────────────────────────────────────────────────────
		void tab_RWYs()
		{
			tabPage5.VerticalScroll.Value = 0;
			ClearTaggedControls(tabPage5);

			Label hdr = new Label { Tag = "dispose", Top = 18, Left = 20, AutoSize = true,
				Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold), Text = "Runways" };
			tabPage5.Controls.Add(hdr);

			_rwyCmb = new ComboBox { Tag = "dispose", Top = 50, Left = 20, Width = 160, DropDownStyle = ComboBoxStyle.DropDownList };
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			conn.Open();
			OleDbDataReader r = new OleDbCommand("SELECT ICAO FROM Stations_ICAO_IATA ORDER BY ICAO", conn).ExecuteReader();
			while (r.Read()) if (!r.IsDBNull(0)) _rwyCmb.Items.Add(r.GetString(0));
			conn.Close();
			_rwyCmb.SelectedIndexChanged += (s, e) => RwyLoadGrid(_rwyCmb.Text);
			tabPage5.Controls.Add(_rwyCmb);

			Button bReimport = new Button { Tag = "dispose", Top = 49, Left = 190, Size = new Size(150, 26),
				Text = "↻ Re-import CSV", BackColor = Color.SeaGreen, ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
			bReimport.Click += (s, e) => { if (_rwyCmb.Text != "") { ImportRunwaysFromCsv(_rwyCmb.Text); RegenerateRwyMemo(_rwyCmb.Text); RwyLoadGrid(_rwyCmb.Text); } };
			tabPage5.Controls.Add(bReimport);

			Button bAdd = new Button { Tag = "dispose", Top = 49, Left = 350, Size = new Size(90, 26), Text = "+ Add RWY" };
			bAdd.Click += (s, e) => { if (_rwyDgv != null) _rwyDgv.Rows.Add(); };
			tabPage5.Controls.Add(bAdd);

			Button bSave = new Button { Tag = "dispose", Top = 49, Left = 450, Size = new Size(110, 26),
				Text = "Save", BackColor = Color.FromArgb(38,50,56), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
			bSave.Click += (s, e) => { if (_rwyCmb.Text != "") { RwySaveGrid(_rwyCmb.Text); RegenerateRwyMemo(_rwyCmb.Text); } };
			tabPage5.Controls.Add(bSave);

			_rwyDgv = new DataGridView { Tag = "dispose", Top = 85, Left = 20, Size = new Size(820, 460),
				AllowUserToAddRows = true, RowHeadersWidth = 28, BackgroundColor = Color.White,
				AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None };
			AddCol("QFU", 70); AddCol("Max APPCH (CAT)", 150); AddCol("Dist (m)", 80);
			AddCol("Hdg", 70); AddCol("Thr Lat", 130); AddCol("Thr Lon", 130);
			tabPage5.Controls.Add(_rwyDgv);

			if (_rwyCmb.Items.Count > 0) { _rwyCmb.SelectedIndex = 0; RwyLoadGrid(_rwyCmb.Text); }
		}

		private void AddCol(string header, int width)
		{
			DataGridViewTextBoxColumn c = new DataGridViewTextBoxColumn { HeaderText = header, Width = width };
			if (header.StartsWith("Max APPCH")) c.DefaultCellStyle.BackColor = Color.FromArgb(255, 248, 225);
			_rwyDgv.Columns.Add(c);
		}

		void RwyLoadGrid(string icao)
		{
			if (_rwyDgv == null || icao == "") return;
			_rwyDgv.Rows.Clear();

			// Auto-populate the structured table on first view: migrate the legacy memo
			// (keeps the manual CAT) if present, otherwise import fresh from the CSV.
			int count = 0;
			OleDbConnection c1 = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			c1.Open();
			OleDbCommand cc = new OleDbCommand("SELECT COUNT(*) FROM Runways WHERE ICAO=?", c1);
			cc.Parameters.AddWithValue("?", icao);
			count = Convert.ToInt32(cc.ExecuteScalar());
			string memo = "";
			if (count == 0)
			{
				OleDbCommand mq = new OleDbCommand("SELECT RWYs FROM Stations_ICAO_IATA WHERE ICAO=?", c1);
				mq.Parameters.AddWithValue("?", icao);
				object mv = mq.ExecuteScalar();
				memo = (mv != null && mv != DBNull.Value) ? mv.ToString() : "";
			}
			c1.Close();
			if (count == 0)
			{
				if (memo.Trim() != "") MigrateLegacyMemo(icao, memo);
				else                   ImportRunwaysFromCsv(icao);
				RegenerateRwyMemo(icao);
			}

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			conn.Open();
			OleDbCommand q = new OleDbCommand("SELECT QFU, Cat, DistM, Hdg, ThrLat, ThrLon FROM Runways WHERE ICAO=? ORDER BY Ord", conn);
			q.Parameters.AddWithValue("?", icao);
			OleDbDataReader r = q.ExecuteReader();
			while (r.Read())
			{
				_rwyDgv.Rows.Add(
					!r.IsDBNull(0) ? r.GetString(0) : "",
					!r.IsDBNull(1) ? r.GetString(1) : "",
					!r.IsDBNull(2) ? Convert.ToInt32(r.GetValue(2)).ToString() : "",
					!r.IsDBNull(3) ? Convert.ToDouble(r.GetValue(3)).ToString("0.#", CultureInfo.InvariantCulture) : "",
					!r.IsDBNull(4) ? Convert.ToDouble(r.GetValue(4)).ToString("0.#####", CultureInfo.InvariantCulture) : "",
					!r.IsDBNull(5) ? Convert.ToDouble(r.GetValue(5)).ToString("0.#####", CultureInfo.InvariantCulture) : "");
			}
			r.Close();
			conn.Close();
		}

		void RwySaveGrid(string icao)
		{
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= OCC.mdb");
			conn.Open();
			OleDbCommand del = new OleDbCommand("DELETE FROM Runways WHERE ICAO=?", conn);
			del.Parameters.AddWithValue("?", icao);
			del.ExecuteNonQuery();
			int ord = 0;
			foreach (DataGridViewRow row in _rwyDgv.Rows)
			{
				if (row.IsNewRow) continue;
				string qfu = Cell(row, 0);
				if (qfu == "") continue;
				InsertRwy(conn, icao, qfu, Cell(row, 1),
					(int)ParseD(Cell(row, 2)), ParseD(Cell(row, 3)), ParseD(Cell(row, 4)), ParseD(Cell(row, 5)), ord++);
			}
			conn.Close();
		}

		private static string Cell(DataGridViewRow row, int i)
		{
			object v = row.Cells[i].Value;
			return v == null ? "" : v.ToString().Trim();
		}
	}
}
