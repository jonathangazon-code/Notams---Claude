using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ICAO_CSV
{
	public partial class MainForm
	{
		private DataGridView _fsDgv;
		private System.ComponentModel.BackgroundWorker _fsWorker;
		private Form _fsProgressForm;
		private Label _fsProgressStatus;
		private ProgressBar _fsProgressBar;

		private static readonly XNamespace FsNs = "http://www.fwz.aero/Schemas/StandardInterfaces";

		// Flight schedule cache in ICAO_storedNotams.mdb — one row per Fltleg_ID. STA is
		// the expensive field (one getBriefing call per flight), so it's kept here and
		// only re-fetched when a flight's STD has actually changed since last cached.
		public void EnsureFlightScheduleTable()
		{
			try
			{
				OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
				conn.Open();
				try { new OleDbCommand("CREATE TABLE FlightSchedule ([FltlegID] LONG, [FLDt] TEXT(10), [Callsign] TEXT(20), [Reg] TEXT(10), [STD] TEXT(25), [STA] TEXT(25), [Crew] TEXT(255))", conn).ExecuteNonQuery(); }
				catch { /* already exists */ }
				conn.Close();
			}
			catch { }
		}

		void FlightScheduleTabEnter(object sender, EventArgs e)
		{
			if (_fsDgv == null) BuildFlightScheduleGrid();
			RefreshFlightSchedule();
		}

		private void BuildFlightScheduleGrid()
		{
			ClearTaggedControls(tabPage_FlightSchedule);

			Label hdr = new Label { Tag = "dispose", Top = 18, Left = 20, AutoSize = true,
				Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Bold), Text = "Flight Schedule — next 7 days (UTC)" };
			tabPage_FlightSchedule.Controls.Add(hdr);

			Button refresh = new Button { Tag = "dispose", Top = 12, Left = 400, Width = 90, Height = 28, Text = "Refresh" };
			refresh.Click += delegate { RefreshFlightSchedule(); };
			tabPage_FlightSchedule.Controls.Add(refresh);

			_fsDgv = new DataGridView { Tag = "dispose", Top = 55, Left = 20, Size = new Size(950, 900),
				ReadOnly = true, AllowUserToAddRows = false, RowHeadersWidth = 28, BackgroundColor = Color.White,
				AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None };
			_fsDgv.ColumnHeadersDefaultCellStyle.Font = new Font(_fsDgv.Font, FontStyle.Bold);
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Date",     HeaderText = "Date",      Width = 90 });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Callsign", HeaderText = "Callsign",  Width = 100 });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Reg",      HeaderText = "Reg",       Width = 80 });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "STD",      HeaderText = "STD (UTC)", Width = 140 });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "STA",      HeaderText = "STA (UTC)", Width = 140 });
			_fsDgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Crew",     HeaderText = "Crew",      Width = 380 });
			tabPage_FlightSchedule.Controls.Add(_fsDgv);
		}

		// Downloads getFlightList for today..today+7, resolves STA via a cached getBriefing
		// call per flight (skipped when the flight's STD hasn't moved since the last
		// fetch), stores everything in FlightSchedule and reloads the grid sorted by STD.
		// Runs on a BackgroundWorker with a progress dialog — a 7-day window can mean a few
		// hundred individual getBriefing calls.
		public void RefreshFlightSchedule()
		{
			if (_fsWorker != null && _fsWorker.IsBusy) return;

			EnsureFlightScheduleTable();
			EnsureArchiveConfig();
			ShowFsProgressForm();

			_fsWorker = new System.ComponentModel.BackgroundWorker { WorkerReportsProgress = true };
			_fsWorker.DoWork += delegate(object s, System.ComponentModel.DoWorkEventArgs e)
			{
				System.ComponentModel.BackgroundWorker w = (System.ComponentModel.BackgroundWorker)s;
				FetchFlightSchedule(delegate(int pct, string msg) { w.ReportProgress(pct, msg); });
			};
			_fsWorker.ProgressChanged += delegate(object s, System.ComponentModel.ProgressChangedEventArgs e)
			{
				UpdateFsProgress(e.ProgressPercentage, e.UserState as string);
			};
			_fsWorker.RunWorkerCompleted += delegate(object s, System.ComponentModel.RunWorkerCompletedEventArgs e)
			{
				if (e.Error != null)
				{
					CloseFsProgressForm();
					MessageBox.Show("Error while loading the flight schedule:\n" + e.Error.Message, "Flight Schedule Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
				LoadFlightScheduleGrid();
				CloseFsProgressForm();
			};
			_fsWorker.RunWorkerAsync();
		}

		private static string El(XElement parent, XName name)
		{
			XElement e = parent.Element(name);
			return e != null ? e.Value : "";
		}

		private void FetchFlightSchedule(Action<int, string> onProgress)
		{
			Dictionary<int, string[]> cached = new Dictionary<int, string[]>();   // FltlegID -> [STD, STA]
			OleDbConnection rconn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			rconn.Open();
			OleDbDataReader rdr = new OleDbCommand("SELECT FltlegID, STD, STA FROM FlightSchedule", rconn).ExecuteReader();
			while (rdr.Read())
			{
				if (rdr.IsDBNull(0)) continue;
				int id = Convert.ToInt32(rdr.GetValue(0));
				string std = rdr.IsDBNull(1) ? "" : rdr.GetString(1);
				string sta = rdr.IsDBNull(2) ? "" : rdr.GetString(2);
				cached[id] = new string[] { std, sta };
			}
			rconn.Close();

			List<object[]> flights = new List<object[]>();   // FltlegID, FLDt, Callsign, Reg, STD, Crew

			for (int dayOffset = 0; dayOffset <= 7; dayOffset++)
			{
				DateTime day = DateTime.UtcNow.Date.AddDays(dayOffset);
				string fldt = day.ToString("ddMMMyy", CultureInfo.InvariantCulture).ToUpper();
				onProgress(dayOffset * 10 / 8, "Downloading flight list for " + fldt + "...");

				string xml;
				try
				{
					using (WebClient wc = new WebClient())
						xml = wc.DownloadString(_flightScheduleBaseUrl.TrimEnd('/') + "/?METHOD=getFlightList&FLDT=" + fldt);
				}
				catch { continue; }   // a single day's failure shouldn't abort the whole refresh

				XDocument doc;
				try { doc = XDocument.Parse(xml); } catch { continue; }

				foreach (XElement flight in doc.Descendants(FsNs + "Flight"))
				{
					int fltlegId;
					if (!int.TryParse(El(flight, FsNs + "Fltleg_ID"), out fltlegId) || fltlegId == 0) continue;

					string callsign = El(flight, FsNs + "FPfx") + El(flight, FsNs + "FLNr");
					string reg      = El(flight, FsNs + "ACREG");
					string std      = El(flight, FsNs + "STD");
					string fldtVal  = El(flight, FsNs + "FLDt");

					List<string> crewParts = new List<string>();
					XElement crewMembers = flight.Element(FsNs + "CrewMembers");
					if (crewMembers != null)
						foreach (XElement cm in crewMembers.Elements(FsNs + "crewMember"))
						{
							string name = El(cm, FsNs + "name");
							string function = El(cm, FsNs + "function");
							if (name != "") crewParts.Add(name + " (" + function + ")");
						}
					string crew = string.Join(" / ", crewParts.ToArray());

					flights.Add(new object[] { fltlegId, fldtVal, callsign, reg, std, crew });
				}
			}

			OleDbConnection wconn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			wconn.Open();

			int total = flights.Count, done = 0;
			foreach (object[] f in flights)
			{
				int    fltlegId = (int)f[0];
				string fldtVal  = (string)f[1];
				string callsign = (string)f[2];
				string reg      = (string)f[3];
				string std      = (string)f[4];
				string crew     = (string)f[5];

				string[] cache;
				bool exists = cached.TryGetValue(fltlegId, out cache);
				string sta = (exists && cache[0] == std && cache[1] != "") ? cache[1] : FetchSta(fltlegId);

				if (exists)
				{
					OleDbCommand upd = new OleDbCommand("UPDATE FlightSchedule SET FLDt=?, Callsign=?, Reg=?, STD=?, STA=?, Crew=? WHERE FltlegID=?", wconn);
					upd.Parameters.AddWithValue("?", fldtVal);
					upd.Parameters.AddWithValue("?", callsign);
					upd.Parameters.AddWithValue("?", reg);
					upd.Parameters.AddWithValue("?", std);
					upd.Parameters.AddWithValue("?", sta);
					upd.Parameters.AddWithValue("?", crew);
					upd.Parameters.AddWithValue("?", fltlegId);
					upd.ExecuteNonQuery();
				}
				else
				{
					OleDbCommand ins = new OleDbCommand("INSERT INTO FlightSchedule ([FltlegID],[FLDt],[Callsign],[Reg],[STD],[STA],[Crew]) VALUES (?,?,?,?,?,?,?)", wconn);
					ins.Parameters.AddWithValue("?", fltlegId);
					ins.Parameters.AddWithValue("?", fldtVal);
					ins.Parameters.AddWithValue("?", callsign);
					ins.Parameters.AddWithValue("?", reg);
					ins.Parameters.AddWithValue("?", std);
					ins.Parameters.AddWithValue("?", sta);
					ins.Parameters.AddWithValue("?", crew);
					ins.ExecuteNonQuery();
				}

				done++;
				onProgress(10 + done * 90 / Math.Max(1, total), "Loading briefing " + done + " / " + total + "...");
			}
			wconn.Close();
		}

		// A getBriefing response bundles several BriefingContentProduct blocks (OFP, MET,
		// NOTAM...); ScheduledTimeOfArrival appears once, deep under the OFP's
		// FlightPlanSummary. A plain regex search (same style already used for the NOTAM
		// XML in MainForm.NotamData.cs) is simpler and more robust here than walking the
		// deeply nested, namespaced OFP tree with LINQ to XML.
		private string FetchSta(int fltlegId)
		{
			try
			{
				string xml;
				using (WebClient wc = new WebClient())
					xml = wc.DownloadString(_briefingBaseUrl.TrimEnd('/') + "/?METHOD=getBriefing&FLTLEG_ID=" + fltlegId);
				Match m = Regex.Match(xml, "<ScheduledTimeOfArrival>([^<]*)</ScheduledTimeOfArrival>");
				return m.Success ? m.Groups[1].Value : "";
			}
			catch { return ""; }
		}

		private void LoadFlightScheduleGrid()
		{
			if (_fsDgv == null) return;
			_fsDgv.Rows.Clear();

			OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data source= ICAO_storedNotams.mdb");
			conn.Open();
			OleDbDataReader reader = new OleDbCommand("SELECT FLDt, Callsign, Reg, STD, STA, Crew FROM FlightSchedule ORDER BY STD", conn).ExecuteReader();
			while (reader.Read())
			{
				string fldt     = reader.IsDBNull(0) ? "" : reader.GetString(0);
				string callsign = reader.IsDBNull(1) ? "" : reader.GetString(1);
				string reg      = reader.IsDBNull(2) ? "" : reader.GetString(2);
				string std      = reader.IsDBNull(3) ? "" : reader.GetString(3);
				string sta      = reader.IsDBNull(4) ? "" : reader.GetString(4);
				string crew     = reader.IsDBNull(5) ? "" : reader.GetString(5);
				_fsDgv.Rows.Add(fldt, callsign, reg, std, sta, crew);
			}
			conn.Close();
		}

		// ── progress dialog (dark, matches the DB Update dialog's styling) ────
		private void ShowFsProgressForm()
		{
			_fsProgressForm = new Form
			{
				StartPosition = FormStartPosition.CenterScreen,
				FormBorderStyle = FormBorderStyle.FixedDialog,
				ControlBox = false, MinimizeBox = false, MaximizeBox = false,
				TopMost = true, ShowInTaskbar = false,
				Width = 420, Height = 130,
				BackColor = Color.FromArgb(38, 50, 56),
				Text = "Dispatch Watch"
			};
			Label title = new Label { Text = "LOADING FLIGHT SCHEDULE", ForeColor = Color.White,
				Font = new Font("Segoe UI", 10, FontStyle.Bold), Dock = DockStyle.Top, Height = 32,
				TextAlign = ContentAlignment.MiddleCenter };
			_fsProgressStatus = new Label { Text = "Starting...", ForeColor = Color.FromArgb(207, 216, 220),
				Font = new Font("Segoe UI", 9), Dock = DockStyle.Top, Height = 26, TextAlign = ContentAlignment.MiddleCenter };
			_fsProgressBar = new ProgressBar { Minimum = 0, Maximum = 100, Value = 0, Style = ProgressBarStyle.Continuous, Dock = DockStyle.Bottom, Height = 20 };
			Panel pad = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(38, 50, 56) };
			_fsProgressForm.Controls.Add(pad);
			_fsProgressForm.Controls.Add(_fsProgressBar);
			_fsProgressForm.Controls.Add(_fsProgressStatus);
			_fsProgressForm.Controls.Add(title);
			_fsProgressForm.Show(this);
		}

		private void UpdateFsProgress(int percent, string status)
		{
			if (_fsProgressForm == null || _fsProgressForm.IsDisposed) return;
			_fsProgressBar.Value = Math.Max(0, Math.Min(100, percent));
			if (status != null) _fsProgressStatus.Text = status;
		}

		private void CloseFsProgressForm()
		{
			if (_fsProgressForm != null && !_fsProgressForm.IsDisposed) _fsProgressForm.Close();
			_fsProgressForm = null;
		}
	}
}
