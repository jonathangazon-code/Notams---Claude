using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Windows.Forms;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		// Creates the Keywords table if missing and seeds it with the defaults (idempotent).
		public void EnsureKeywordsTable()
		{
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				try { new OleDbCommand("CREATE TABLE Keywords ([Word] TEXT(50))", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }

				int count = 0;
				try { count = Convert.ToInt32(new OleDbCommand("SELECT COUNT(*) FROM Keywords", conn).ExecuteScalar()); }
				catch { }
				if (count == 0)
				{
					foreach (string w in _defaultKeywords)
					{
						OleDbCommand ins = new OleDbCommand("INSERT INTO Keywords ([Word]) VALUES (?)", conn);
						ins.Parameters.AddWithValue("?", w);
						ins.ExecuteNonQuery();
					}
				}
				conn.Close();
			}
			catch { /* DB not ready; keep defaults */ }
		}

		// Loads the keyword list from the table into _notamKeywords.
		public void LoadKeywords()
		{
			List<string> list = new List<string>();
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				OleDbDataReader r = new OleDbCommand("SELECT Word FROM Keywords ORDER BY Word", conn).ExecuteReader();
				while (r.Read())
					if (!r.IsDBNull(0) && r.GetString(0).Trim() != "") list.Add(r.GetString(0).Trim());
				conn.Close();
			}
			catch { }
			if (list.Count > 0) _notamKeywords = list.ToArray();
		}

		// Fills the ListBox on the Keywords tab.
		void Keywords_Refresh()
		{
			Lst_Keywords.Items.Clear();
			foreach (string w in _notamKeywords) Lst_Keywords.Items.Add(w);
		}

		void TxtBox_keywordKeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				Btn_addKeywordClick(sender, e);
				e.SuppressKeyPress = true;   // avoid the Windows "ding"
			}
		}

		void Btn_addKeywordClick(object sender, EventArgs e)
		{
			string w = TxtBox_keyword.Text.Trim().ToUpper();
			if (w == "") return;

			// Reject duplicates (case-insensitive)
			foreach (string k in _notamKeywords)
				if (k.Equals(w, StringComparison.OrdinalIgnoreCase)) { TxtBox_keyword.Clear(); return; }

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand ins = new OleDbCommand("INSERT INTO Keywords ([Word]) VALUES (?)", conn);
			ins.Parameters.AddWithValue("?", w);
			ins.ExecuteNonQuery();
			conn.Close();

			TxtBox_keyword.Clear();
			LoadKeywords();
			Keywords_Refresh();
		}

		void Btn_removeKeywordClick(object sender, EventArgs e)
		{
			if (Lst_Keywords.SelectedItem == null) return;
			string w = Lst_Keywords.SelectedItem.ToString();

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbCommand del = new OleDbCommand("DELETE FROM Keywords WHERE Word=?", conn);
			del.Parameters.AddWithValue("?", w);
			del.ExecuteNonQuery();
			conn.Close();

			LoadKeywords();
			Keywords_Refresh();
		}
	}
}
