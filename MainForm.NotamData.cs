using System;
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
			RchTxtCSV.Text = result;
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

		public void deleteWithdrawnedNotams()
		{
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbDataReader reader = new OleDbCommand("SELECT key FROM filteredNotams_table", conn).ExecuteReader();

			System.Collections.Generic.List<string> filtered = new System.Collections.Generic.List<string>();
			while (reader.Read())
				if (!reader.IsDBNull(0)) filtered.Add(reader.GetString(0));
			conn.Close();

			System.Collections.Generic.List<string> withdrawn = new System.Collections.Generic.List<string>();
			conn.Open();
			foreach (string key in filtered)
			{
				OleDbCommand chk = new OleDbCommand("SELECT COUNT(*) FROM storedNotams_table WHERE key=?", conn);
				chk.Parameters.AddWithValue("?", key);
				if ((int)chk.ExecuteScalar() == 0)
					withdrawn.Add(key);
			}
			conn.Close();

			string monitor = "";
			conn.Open();
			foreach (string key in withdrawn)
			{
				OleDbCommand del = new OleDbCommand("DELETE FROM filteredNotams_table WHERE key=?", conn);
				del.Parameters.AddWithValue("?", key);
				del.ExecuteNonQuery();
				monitor += key + "\n";
			}
			conn.Close();
			RchTxtCSV.Text = monitor;
		}

		public void NewNotams()
		{
			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbDataReader reader = new OleDbCommand("SELECT key FROM storedNotams_table", conn).ExecuteReader();

			System.Collections.Generic.List<string> stored = new System.Collections.Generic.List<string>();
			while (reader.Read())
				if (!reader.IsDBNull(0)) stored.Add(reader.GetString(0));
			conn.Close();

			System.Collections.Generic.List<string> newKeys = new System.Collections.Generic.List<string>();
			conn.Open();
			foreach (string key in stored)
			{
				OleDbCommand chk = new OleDbCommand("SELECT COUNT(*) FROM filteredNotams_table WHERE key=?", conn);
				chk.Parameters.AddWithValue("?", key);
				if ((int)chk.ExecuteScalar() == 0)
					newKeys.Add(key);
			}
			conn.Close();

			string testFilter = "";
			foreach (string key in newKeys)
			{
				string StateName="", Subject="", Modifier="", message="", startdate="", enddate="", all="", location="", created="";
				conn.Open();
				OleDbCommand qry = new OleDbCommand("SELECT * FROM storedNotams_table WHERE key=?", conn);
				qry.Parameters.AddWithValue("?", key);
				OleDbDataReader r = qry.ExecuteReader();
				while (r.Read())
				{
					if (!r.IsDBNull(1))  StateName  = r.GetString(1);
					if (!r.IsDBNull(10)) Subject    = r.GetString(10);
					if (!r.IsDBNull(11)) Modifier   = r.GetString(11);
					if (!r.IsDBNull(12)) message    = r.GetString(12);
					if (!r.IsDBNull(13)) startdate  = r.GetString(13);
					if (!r.IsDBNull(14)) enddate    = r.GetString(14);
					if (!r.IsDBNull(15)) all        = r.GetString(15);
					if (!r.IsDBNull(16)) location   = r.GetString(16);
					if (!r.IsDBNull(18)) created    = r.GetString(18);
				}
				conn.Close();

				conn.Open();
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
				conn.Close();

				testFilter += StateName + "||" + Subject + "||" + Modifier + "||" + message + "||" +
					startdate + "||" + enddate + "||" + all + "||" + location + "||" + created + "||N\n";
			}
			RchTxtCSV.Text = testFilter;
		}

		public void DelOld()
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
			conn.Close();

			string deletedNotams = "";
			conn.Open();
			foreach (string key in toDelete)
			{
				if (string.IsNullOrEmpty(key)) continue;
				OleDbCommand del = new OleDbCommand("DELETE FROM filteredNotams_table WHERE key=?", conn);
				del.Parameters.AddWithValue("?", key);
				del.ExecuteNonQuery();
				deletedNotams += key;
			}
			conn.Close();

			RchTxtCSV.Text = todayList + "\n\n" + deletedNotams + "\n\n" + endDateList;
		}

		void Btn_updateDBClick(object sender, EventArgs e)
		{
			ShowAutoPopup("Downloading XML Notams from Web Service...");
			GetXML();
			ShowAutoPopup("Deleting withdrawn NOTAMs...");
			deleteWithdrawnedNotams();
			ShowAutoPopup("Adding new NOTAMs...");
			NewNotams();
			ShowAutoPopup("Deleting old NOTAMs...");
			DelOld();
			MessageBox.Show("Database successfully updated!", "DB Updated");
		}

		// Event to track CSV download progress
		void wc_DownloadProgressChanged(object sender, System.Net.DownloadProgressChangedEventArgs e)
		{
			progressBar.Value = e.ProgressPercentage;
		}
	}
}
