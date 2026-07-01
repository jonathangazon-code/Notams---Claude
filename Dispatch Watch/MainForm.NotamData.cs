using System;
using System.Drawing;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		public void Reload_text() { }

		public void GetXML()
		{
			if (_stationsCache == null) LoadStationsCache();
			string APList = "";
			foreach (var entry in _stationsCache)
			{
				string[] row = entry.Value;
				if (row[1] == "Yes" || row[2] == "Yes" || row[3] == "Yes")
					APList += entry.Key + "-";
			}
			APList = APList.TrimEnd('-');

			OleDbConnection connNotams = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			connNotams.Open();
			new OleDbCommand("DELETE FROM storedNotams_table", connNotams).ExecuteNonQuery();
			connNotams.Close();

			DateTime today       = DateTime.Now;
			DateTime threeMonths = today.AddDays(91);
			string todayStr  = today.ToString("yyyyMMdd");
			string threeStr  = threeMonths.ToString("yyyyMMdd");

			string stringToday       = todayStr.Substring(6,2)  + MonthAbbrev(todayStr.Substring(4,2))  + todayStr.Substring(2,2);
			string stringThreeMonths = threeStr.Substring(6,2)  + MonthAbbrev(threeStr.Substring(4,2))  + threeStr.Substring(2,2);

string xmlNotams = "";
			using (WebClient wc = new WebClient())
				xmlNotams = wc.DownloadString("http://10.48.12.43:5455/BriefingService.svc/web/?METHOD=getAdHocNOTAM&AIRPORTS=" + APList + "&PERIODSTART=" + stringToday + "&PERIODEND=" + stringThreeMonths);

			xmlNotams = Regex.Replace(xmlNotams, @"\t|\n|\r", " ");

			string key="", series="", serial="", year="", all="", result="";
			string startDate="", endDate="", creationDate="", revisionDate="", location="", NOTAMInfo="";

			connNotams.Open();
			foreach (string search1 in Regex.Split(xmlNotams, "<NOTAM issuer="))
			{
				foreach (string search2 in Regex.Split(search1, "series=\""))
				{
					series = search2.Substring(0, 1);
					if (series == "<" || series == "\"" || series == "0") continue;

					foreach (string search3 in Regex.Split(search2, "serial=\""))
					{
						if (search3.Length <= 4) continue;
						serial = search3.Substring(0, 4);

						foreach (string search4 in Regex.Split(search3, "year=\""))
						{
							if (search4.Length <= 2) continue;
							year = search4.Substring(0, 2);

							foreach (string search5 in Regex.Split(search4, "startValidTime=\""))
							{
								if (search5.Length <= 19) continue;
								startDate = search5.Substring(0, 19) + ".000Z";

								foreach (string search6 in Regex.Split(search5, "endValidTime=\""))
								{
									if (search6.Length <= 19) continue;
									endDate = search6.Substring(0, 19) + ".000Z";

									foreach (string search7 in Regex.Split(search6, "creationTime=\""))
									{
										if (search7.Length <= 19) continue;
										creationDate = search7.Substring(0, 19) + ".000Z";

										foreach (string search8 in Regex.Split(search7, "creationTime=\""))
										{
											if (search8.Length <= 19) continue;
											revisionDate = search8.Substring(0, 19) + ".000Z";

											foreach (string search9 in Regex.Split(search8, "<NOTAMText><Paragraph><Text>"))
											{
												if (search9.Length <= 2) continue;
												all = Regex.Split(search9, "</Text>")[0];

												foreach (string search10 in Regex.Split(search9, "<AirportICAOCode>"))
												{
													if (search10.Length <= 4) continue;
													location = search10.Substring(0, 4);
													NOTAMInfo = "No";
													if (Regex.IsMatch(search10, "</ItemC><ItemD>"))
													{
														NOTAMInfo = "Yes";
														foreach (string search11 in Regex.Split(search10, "</ItemC><ItemD>"))
															if (search11.Length > 4)
																NOTAMInfo = Regex.Split(search11, "</ItemD>")[0];
													}
												}
											}
										}
									}
								}
							}
						}
					}

					key     = series + serial + "/" + year + "-" + location;
					result += key + "\n" + startDate + " - " + endDate + "\n" + all + "\nCreated: " + creationDate + "\nRevised:" + revisionDate + "\n\n";

					if (NOTAMInfo != "") all = NOTAMInfo + "\n" + all;
					all = all.Replace("'", "(char)39").Replace("\"", "(char)34");

					OleDbCommand ins = new OleDbCommand(
						"INSERT INTO storedNotams_table ([Notam_id],[startdate],[enddate],[all],[location],[Created],[key]) VALUES (?,?,?,?,?,?,?)",
						connNotams);
					ins.Parameters.AddWithValue("?", key);
					ins.Parameters.AddWithValue("?", startDate);
					ins.Parameters.AddWithValue("?", endDate);
					ins.Parameters.AddWithValue("?", all);
					ins.Parameters.AddWithValue("?", location);
					ins.Parameters.AddWithValue("?", creationDate + " // " + revisionDate);
					ins.Parameters.AddWithValue("?", key);
					ins.ExecuteNonQuery();
				}
			}
			connNotams.Close();
			SetLog(result);
		}

		public void GetCSV()
		{
			if (_stationsCache == null) LoadStationsCache();
			string APList = "";
			foreach (var entry in _stationsCache)
			{
				string[] row = entry.Value;
				if (row[1] == "Yes" || row[2] == "Yes" || row[3] == "Yes")
					APList += entry.Key + ",";
			}
			APList = APList.TrimEnd(',');

			ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
			string url = @"https://applications.icao.int/dataservices/api/notams-list?api_key=c62b1e08-c60e-41ba-a654-30cb2807a682&format=csv&type=&Qcode=&locations=" + APList + "&qstring=&states=&ICAOonly=";
			using (WebClient wc = new WebClient())
				wc.DownloadFile(url, "ICAO-CSV.csv");

			RchTxtCSV.Text = "CSV Downloaded";
		}

		public void Split()
		{
			OleDbConnection connNotams = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			connNotams.Open();
			new OleDbCommand("DELETE FROM storedNotams_table", connNotams).ExecuteNonQuery();
			connNotams.Close();

			string[] lines = File.ReadAllLines("ICAO-CSV.csv");
			string reader  = "";
			Regex CSVParser = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))", RegexOptions.Compiled);

			for (int i = 0; i < lines.Length; i++)
			{
				if (i == 0) { reader += lines[i] + "\n"; continue; }

				string[] subs = CSVParser.Split(lines[i]);
				string[] fields = new string[20];

				for (int j = 0; j < subs.Length && j < 20; j++)
				{
					string v = subs[j];
					if (v.Length > 2) v = v.Substring(1, v.Length - 2);
					if (v.Length > 2 && v.Substring(0, 2) == "\"\"") v = v.Substring(2, v.Length - 4);
					v = v.Replace("'", "(char)39").Replace("\"", "(char)34");
					fields[j] = v;
					reader += v + (j == 19 ? "\n \n ------------------------------- \n \n" : "\n");
				}

				connNotams.Open();
				OleDbCommand ins = new OleDbCommand(
					"INSERT INTO storedNotams_table ([StateName],[StateCode],[Notam_id],[entity],[status],[Qcode],[Area],[SubArea],[Condition],[Subject],[Modifier],[message],[startdate],[enddate],[all],[location],[isICAO],[Created],[key],[type]) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
					connNotams);
				for (int k = 0; k < 20; k++)
					ins.Parameters.AddWithValue("?", fields[k] ?? "");
				ins.ExecuteNonQuery();
				connNotams.Close();
			}

			reader = reader.Replace("(char)39", "'").Replace("(char)34", "\"");
			RchTxtCSV.Text = reader;
		}

		// Thread-safe write to the debug/log textbox — the DB-update pipeline runs these
		// methods on a background thread, so a direct ".Text =" would throw a cross-thread
		// exception.
		private void SetLog(string text)
		{
			if (RchTxtCSV.InvokeRequired) RchTxtCSV.Invoke((MethodInvoker)delegate { RchTxtCSV.Text = text; });
			else RchTxtCSV.Text = text;
		}

		// Records when the DB Update pipeline last completed successfully. Read back by
		// RefreshLastDbUpdateLabel (MainForm.NotamFilter.cs) — deliberately separate from the
		// .mdb file's own last-write time, which also changes on V: drive sync.
		private void SaveLastDbUpdateTimestamp()
		{
			try
			{
				string tsPath = System.IO.Path.Combine(Application.StartupPath, "last_db_update.txt");
				System.IO.File.WriteAllText(tsPath, DateTime.Now.ToString("o", System.Globalization.CultureInfo.InvariantCulture));
			}
			catch { /* non-critical: only the "last updated" label depends on this */ }
		}

		// onProgress reports (percent 0-100 within this phase, status message) so the caller
		// can weight it into the overall DB-update progress bar. Left null for standalone use
		// (e.g. the individual buttons on the DB Update tab).
		public void deleteWithdrawnedNotams(Action<int, string> onProgress = null)
		{
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();   // single connection for the whole method — was 3 opens before

			System.Collections.Generic.List<string> filtered = new System.Collections.Generic.List<string>();
			OleDbDataReader fReader = new OleDbCommand("SELECT key FROM filteredNotams_table", conn).ExecuteReader();
			while (fReader.Read()) if (!fReader.IsDBNull(0)) filtered.Add(fReader.GetString(0));
			fReader.Close();

			// One pass to snapshot every still-stored key into a set, instead of one
			// SELECT COUNT(*) query per filtered NOTAM (was O(N) queries, now O(1)).
			System.Collections.Generic.HashSet<string> storedKeys = new System.Collections.Generic.HashSet<string>(StringComparer.Ordinal);
			OleDbDataReader sReader = new OleDbCommand("SELECT key FROM storedNotams_table", conn).ExecuteReader();
			while (sReader.Read()) if (!sReader.IsDBNull(0)) storedKeys.Add(sReader.GetString(0));
			sReader.Close();

			System.Collections.Generic.List<string> withdrawn = new System.Collections.Generic.List<string>();
			foreach (string key in filtered) if (!storedKeys.Contains(key)) withdrawn.Add(key);

			string monitor = "";
			int total = withdrawn.Count, done = 0;
			foreach (string key in withdrawn)
			{
				OleDbCommand del = new OleDbCommand("DELETE FROM filteredNotams_table WHERE key=?", conn);
				del.Parameters.AddWithValue("?", key);
				del.ExecuteNonQuery();
				monitor += key + "\n";
				done++;
				if (onProgress != null) onProgress(total == 0 ? 100 : done * 100 / total, "Removing withdrawn NOTAM " + done + "/" + total + "...");
			}
			conn.Close();
			SetLog(monitor);
			if (onProgress != null && total == 0) onProgress(100, "No withdrawn NOTAMs to remove.");
		}

		public void NewNotams(Action<int, string> onProgress = null)
		{
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();   // single connection for the whole method — was 2 + 2*N opens before

			// Snapshot of already-filtered keys, loaded once instead of one COUNT(*) query
			// per candidate row.
			System.Collections.Generic.HashSet<string> existingFiltered = new System.Collections.Generic.HashSet<string>(StringComparer.Ordinal);
			OleDbDataReader fReader = new OleDbCommand("SELECT key FROM filteredNotams_table", conn).ExecuteReader();
			while (fReader.Read()) if (!fReader.IsDBNull(0)) existingFiltered.Add(fReader.GetString(0));
			fReader.Close();

			// Single full-table scan collects the fields for every row that isn't filtered
			// yet, instead of one "SELECT * WHERE key=?" round-trip per new NOTAM. Columns
			// are looked up by name (GetOrdinal) rather than hard-coded position, so this
			// doesn't depend on the table's physical column order.
			System.Collections.Generic.List<string[]> newRows = new System.Collections.Generic.List<string[]>();
			OleDbDataReader sReader = new OleDbCommand("SELECT * FROM storedNotams_table", conn).ExecuteReader();
			int ordKey        = sReader.GetOrdinal("key");
			int ordStateName  = sReader.GetOrdinal("StateName");
			int ordSubject    = sReader.GetOrdinal("Subject");
			int ordModifier   = sReader.GetOrdinal("Modifier");
			int ordMessage    = sReader.GetOrdinal("message");
			int ordStartdate  = sReader.GetOrdinal("startdate");
			int ordEnddate    = sReader.GetOrdinal("enddate");
			int ordAll        = sReader.GetOrdinal("all");
			int ordLocation   = sReader.GetOrdinal("location");
			int ordCreated    = sReader.GetOrdinal("Created");
			while (sReader.Read())
			{
				string key = !sReader.IsDBNull(ordKey) ? sReader.GetString(ordKey) : "";
				if (key == "" || existingFiltered.Contains(key)) continue;

				string[] row = new string[10];
				row[0] = key;
				row[1] = !sReader.IsDBNull(ordStateName) ? sReader.GetString(ordStateName) : "";
				row[2] = !sReader.IsDBNull(ordSubject)   ? sReader.GetString(ordSubject)   : "";
				row[3] = !sReader.IsDBNull(ordModifier)  ? sReader.GetString(ordModifier)  : "";
				row[4] = !sReader.IsDBNull(ordMessage)   ? sReader.GetString(ordMessage)   : "";
				row[5] = !sReader.IsDBNull(ordStartdate) ? sReader.GetString(ordStartdate) : "";
				row[6] = !sReader.IsDBNull(ordEnddate)   ? sReader.GetString(ordEnddate)   : "";
				row[7] = !sReader.IsDBNull(ordAll)       ? sReader.GetString(ordAll)       : "";
				row[8] = !sReader.IsDBNull(ordLocation)  ? sReader.GetString(ordLocation)  : "";
				row[9] = !sReader.IsDBNull(ordCreated)   ? sReader.GetString(ordCreated)   : "";
				newRows.Add(row);
			}
			sReader.Close();

			string testFilter = "";
			int total = newRows.Count, done = 0;
			foreach (string[] row in newRows)
			{
				string key = row[0], StateName = row[1], Subject = row[2], Modifier = row[3], message = row[4],
					startdate = row[5], enddate = row[6], all = row[7], location = row[8], created = row[9];

				OleDbCommand ins = new OleDbCommand(
					"INSERT INTO filteredNotams_table ([StateName],[Subject],[Modifier],[message],[startdate],[enddate],[all],[location],[Created],[key],[Checked]) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
					conn);
				ins.Parameters.AddWithValue("?", StateName);
				ins.Parameters.AddWithValue("?", Subject);
				ins.Parameters.AddWithValue("?", Modifier);
				ins.Parameters.AddWithValue("?", message);
				ins.Parameters.AddWithValue("?", startdate);
				ins.Parameters.AddWithValue("?", enddate);
				ins.Parameters.AddWithValue("?", all);
				ins.Parameters.AddWithValue("?", location);
				ins.Parameters.AddWithValue("?", created);
				ins.Parameters.AddWithValue("?", key);
				ins.Parameters.AddWithValue("?", "N");
				ins.ExecuteNonQuery();

				testFilter += StateName + "||" + Subject + "||" + Modifier + "||" + message + "||" +
					startdate + "||" + enddate + "||" + all + "||" + location + "||" + created + "||N\n";
				done++;
				if (onProgress != null) onProgress(total == 0 ? 100 : done * 100 / total, "Adding new NOTAM " + done + "/" + total + "...");
			}
			conn.Close();
			SetLog(testFilter);
			if (onProgress != null && total == 0) onProgress(100, "No new NOTAMs to add.");
		}

		public void DelOld(Action<int, string> onProgress = null)
		{
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbDataReader reader = new OleDbCommand("SELECT enddate,key FROM filteredNotams_table", conn).ExecuteReader();

			DateTime _date = DateTime.Now.ToUniversalTime();
			string today = _date.ToString("yyyyMMddHHmm");
			long intToday = Int64.Parse(today);

			System.Collections.Generic.List<string> toDelete = new System.Collections.Generic.List<string>();
			string endDateList = "", todayList = today + "\n";

			while (reader.Read())
			{
				string endDate = !reader.IsDBNull(0) ? reader.GetString(0) : "";
				endDate = endDate.Replace("-","").Replace("T","").Replace(":","").Replace(".","").Replace("Z","");
				if (endDate.Length <= 11) continue;
				endDate = endDate.Substring(0, 12);
				long intEnd = Int64.Parse(endDate);
				endDateList += endDate + "\n";
				if (intToday > intEnd && !reader.IsDBNull(1))
					toDelete.Add(reader.GetString(1));
			}
			reader.Close();

			string deletedNotams = "";
			int total = toDelete.Count, done = 0;
			foreach (string key in toDelete)
			{
				if (string.IsNullOrEmpty(key)) { done++; continue; }
				OleDbCommand del = new OleDbCommand("DELETE FROM filteredNotams_table WHERE key=?", conn);
				del.Parameters.AddWithValue("?", key);
				del.ExecuteNonQuery();
				deletedNotams += key;
				done++;
				if (onProgress != null) onProgress(total == 0 ? 100 : done * 100 / total, "Removing expired NOTAM " + done + "/" + total + "...");
			}
			conn.Close();

			SetLog(todayList + "\n\n" + deletedNotams + "\n\n" + endDateList);
			if (onProgress != null && total == 0) onProgress(100, "No expired NOTAMs to remove.");
		}

		// ── DB update pipeline: runs on a background thread with a live progress dialog ──

		private System.ComponentModel.BackgroundWorker _dbUpdateWorker;
		private Form _dbProgressForm;
		private Label _dbProgressStatus;
		private ProgressBar _dbProgressBar;

		void Btn_updateDBClick(object sender, EventArgs e) { RunDbUpdatePipeline(null); }

		// Runs GetXML -> deleteWithdrawnedNotams -> NewNotams -> DelOld on a background
		// thread and reports weighted progress into a single dialog, replacing the old
		// sequence of blocking popups + a final MessageBox. onCompleted (optional) runs on
		// the UI thread once the pipeline finishes successfully — used by the NOTAM Filter
		// tab's quick button to refresh its view and the "last update" label afterwards.
		public void RunDbUpdatePipeline(Action onCompleted)
		{
			if (_dbUpdateWorker != null && _dbUpdateWorker.IsBusy) return;   // already running

			Btn_updateDB.Enabled = false;
			if (Btn_dbUpdateQuick != null) Btn_dbUpdateQuick.Enabled = false;
			ShowDbProgressForm();

			_dbUpdateWorker = new System.ComponentModel.BackgroundWorker { WorkerReportsProgress = true };
			_dbUpdateWorker.DoWork += delegate(object s, System.ComponentModel.DoWorkEventArgs e)
			{
				System.ComponentModel.BackgroundWorker w = (System.ComponentModel.BackgroundWorker)s;

				// Phase weights: download has no measurable sub-progress, so it just jumps
				// to 30%; the three DB-diff phases report fine-grained progress based on the
				// actual NOTAM counts they process.
				w.ReportProgress(2, "Downloading NOTAMs from web service...");
				GetXML();
				w.ReportProgress(30, "Checking for withdrawn NOTAMs...");
				deleteWithdrawnedNotams(delegate(int pct, string msg) { w.ReportProgress(30 + pct * 15 / 100, msg); });
				w.ReportProgress(45, "Adding new NOTAMs...");
				NewNotams(delegate(int pct, string msg) { w.ReportProgress(45 + pct * 45 / 100, msg); });
				w.ReportProgress(90, "Removing expired NOTAMs...");
				DelOld(delegate(int pct, string msg) { w.ReportProgress(90 + pct * 10 / 100, msg); });
			};
			_dbUpdateWorker.ProgressChanged += delegate(object s, System.ComponentModel.ProgressChangedEventArgs e)
			{
				UpdateDbProgress(e.ProgressPercentage, e.UserState as string);
			};
			_dbUpdateWorker.RunWorkerCompleted += delegate(object s, System.ComponentModel.RunWorkerCompletedEventArgs e)
			{
				Btn_updateDB.Enabled = true;
				if (Btn_dbUpdateQuick != null) Btn_dbUpdateQuick.Enabled = true;

				if (e.Error != null)
				{
					CloseDbProgressForm();
					MessageBox.Show("Error while updating the database:\n" + e.Error.Message, "DB Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				SaveLastDbUpdateTimestamp();
				RefreshLastDbUpdateLabel();
				UpdateDbProgress(100, "Database successfully updated!");
				if (onCompleted != null) onCompleted();
				Timer closeTimer = new Timer { Interval = 900 };
				closeTimer.Tick += delegate(object s2, EventArgs e2) { closeTimer.Stop(); CloseDbProgressForm(); };
				closeTimer.Start();
			};
			_dbUpdateWorker.RunWorkerAsync();
		}

		// Dark, non-modal progress dialog matching the app's dark-card styling (used
		// elsewhere for the airport header card). Lives only for the duration of a DB update.
		private void ShowDbProgressForm()
		{
			_dbProgressForm = new Form
			{
				StartPosition   = FormStartPosition.CenterScreen,
				FormBorderStyle = FormBorderStyle.FixedDialog,
				ControlBox      = false,
				MinimizeBox     = false,
				MaximizeBox     = false,
				TopMost         = true,
				ShowInTaskbar   = false,
				Width = 420, Height = 130,
				BackColor = Color.FromArgb(38, 50, 56),
				Text = "Dispatch Watch"
			};

			Label title = new Label
			{
				Text = "UPDATING DATABASE",
				ForeColor = Color.White,
				Font = new Font("Segoe UI", 10, FontStyle.Bold),
				Dock = DockStyle.Top, Height = 32,
				TextAlign = ContentAlignment.MiddleCenter
			};
			_dbProgressStatus = new Label
			{
				Text = "Starting...",
				ForeColor = Color.FromArgb(207, 216, 220),
				Font = new Font("Segoe UI", 9),
				Dock = DockStyle.Top, Height = 26,
				TextAlign = ContentAlignment.MiddleCenter
			};
			_dbProgressBar = new ProgressBar
			{
				Minimum = 0, Maximum = 100, Value = 0,
				Style = ProgressBarStyle.Continuous,
				Dock = DockStyle.Bottom, Height = 20
			};
			Panel pad = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(38, 50, 56) };

			_dbProgressForm.Controls.Add(pad);
			_dbProgressForm.Controls.Add(_dbProgressBar);
			_dbProgressForm.Controls.Add(_dbProgressStatus);
			_dbProgressForm.Controls.Add(title);
			_dbProgressForm.Show(this);
		}

		private void UpdateDbProgress(int percent, string status)
		{
			if (_dbProgressForm == null || _dbProgressForm.IsDisposed) return;
			_dbProgressBar.Value = Math.Max(0, Math.Min(100, percent));
			if (status != null) _dbProgressStatus.Text = status;
		}

		private void CloseDbProgressForm()
		{
			if (_dbProgressForm != null && !_dbProgressForm.IsDisposed) _dbProgressForm.Close();
			_dbProgressForm = null;
		}

		// Event to track CSV download progress
		void wc_DownloadProgressChanged(object sender, System.Net.DownloadProgressChangedEventArgs e)
		{
			progressBar.Value = e.ProgressPercentage;
		}
	}
}
