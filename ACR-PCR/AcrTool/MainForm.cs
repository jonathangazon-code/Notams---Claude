using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace AcrTool
{
	/// <summary>
	/// Built entirely in code rather than through the WinForms designer, so the
	/// layout stays readable in a diff and there is no .resx to keep in step.
	/// </summary>
	public class MainForm : Form
	{
		readonly AcrEngine _engine = new AcrEngine();
		readonly List<AircraftSpec> _fleet = AircraftSpec.Fleet();

		TextBox _pcrBox;
		Label _pcrFeedback;
		CheckBox _metricBox;
		DataGridView _grid;
		Label _provenance;

		bool _libraryReady;

		// Column keys
		const string ColAircraft = "aircraft";
		const string ColLibrary = "library";
		const string ColMtow = "mtow";
		const string ColAcrMtow = "acrmtow";
		const string ColMaxWeight = "maxweight";
		const string ColMargin = "margin";
		const string ColTyres = "tyres";
		const string ColVerdict = "verdict";
		const string ColCheckWeight = "checkweight";
		const string ColAcrCheck = "acrcheck";

		public MainForm()
		{
			Text = "ACR / PCR - maximum weight check";
			StartPosition = FormStartPosition.CenterScreen;
			ClientSize = new Size(1180, 460);
			MinimumSize = new Size(900, 400);
			Font = new Font("Segoe UI", 9f);

			BuildTopBar();
			BuildGrid();
			BuildFooter();

			Shown += MainFormShown;
		}

		void BuildTopBar()
		{
			Panel bar = new Panel();
			bar.Dock = DockStyle.Top;
			bar.Height = 76;
			bar.Padding = new Padding(10, 8, 10, 4);

			Label lbl = new Label();
			lbl.Text = "Published PCR";
			lbl.AutoSize = true;
			lbl.Location = new Point(12, 14);
			bar.Controls.Add(lbl);

			_pcrBox = new TextBox();
			_pcrBox.Location = new Point(110, 10);
			_pcrBox.Width = 170;
			_pcrBox.Font = new Font("Consolas", 11f, FontStyle.Bold);
			_pcrBox.Text = "690/R/B/W/T";
			_pcrBox.KeyDown += PcrBoxKeyDown;
			bar.Controls.Add(_pcrBox);

			Button compute = new Button();
			compute.Text = "Evaluate";
			compute.Location = new Point(292, 9);
			compute.Size = new Size(90, 27);
			compute.Click += delegate { Evaluate(); };
			bar.Controls.Add(compute);

			_metricBox = new CheckBox();
			_metricBox.Text = "Show weights in kg";
			_metricBox.Checked = true;
			_metricBox.AutoSize = true;
			_metricBox.Location = new Point(398, 14);
			_metricBox.CheckedChanged += delegate { Evaluate(); };
			bar.Controls.Add(_metricBox);

			Button selfTest = new Button();
			selfTest.Text = "Self-test";
			selfTest.Location = new Point(546, 9);
			selfTest.Size = new Size(90, 27);
			selfTest.Click += delegate { RunSelfTest(); };
			bar.Controls.Add(selfTest);

			_pcrFeedback = new Label();
			_pcrFeedback.AutoSize = true;
			_pcrFeedback.Location = new Point(12, 46);
			_pcrFeedback.MaximumSize = new Size(1140, 0);
			bar.Controls.Add(_pcrFeedback);

			Controls.Add(bar);
		}

		void BuildGrid()
		{
			_grid = new DataGridView();
			_grid.Dock = DockStyle.Fill;
			_grid.AllowUserToAddRows = false;
			_grid.AllowUserToDeleteRows = false;
			_grid.AllowUserToResizeRows = false;
			_grid.RowHeadersVisible = false;
			_grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
			_grid.BackgroundColor = SystemColors.Window;
			_grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
			_grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;

			AddColumn(ColAircraft, "Aircraft", true, 90);
			AddColumn(ColLibrary, "Library entry", true, 105);
			AddColumn(ColMtow, "MTOW", true, 85);
			AddColumn(ColAcrMtow, "ACR @ MTOW", true, 90);
			AddColumn(ColMaxWeight, "Max weight on this PCR", true, 120);
			AddColumn(ColMargin, "Margin", true, 80);
			AddColumn(ColTyres, "Tyres", true, 110);
			AddColumn(ColVerdict, "Verdict", true, 150);
			AddColumn(ColCheckWeight, "Check a weight", false, 95);
			AddColumn(ColAcrCheck, "ACR at that weight", true, 110);

			_grid.CellEndEdit += GridCellEndEdit;

			foreach (AircraftSpec spec in _fleet)
			{
				int i = _grid.Rows.Add();
				_grid.Rows[i].Tag = spec;
				_grid.Rows[i].Cells[ColAircraft].Value = spec.Display;
				_grid.Rows[i].Cells[ColLibrary].Value = spec.LibraryName;
			}

			Controls.Add(_grid);
			_grid.BringToFront();
		}

		void AddColumn(string name, string header, bool readOnly, int fillWeight)
		{
			DataGridViewTextBoxColumn c = new DataGridViewTextBoxColumn();
			c.Name = name;
			c.HeaderText = header;
			c.ReadOnly = readOnly;
			c.FillWeight = fillWeight;
			c.SortMode = DataGridViewColumnSortMode.NotSortable;
			if (!readOnly)
				c.DefaultCellStyle.BackColor = Color.FromArgb(255, 253, 235);
			_grid.Columns.Add(c);
		}

		void BuildFooter()
		{
			Panel foot = new Panel();
			foot.Dock = DockStyle.Bottom;
			foot.Height = 66;
			foot.Padding = new Padding(12, 6, 12, 6);

			_provenance = new Label();
			_provenance.Dock = DockStyle.Fill;
			_provenance.ForeColor = Color.FromArgb(90, 90, 90);
			foot.Controls.Add(_provenance);

			Controls.Add(foot);
		}

		void MainFormShown(object sender, EventArgs e)
		{
			try
			{
				string path = Path.Combine(Application.StartupPath, "aircraft.xml");
				_engine.Load(path);
				_libraryReady = true;
				Evaluate();
			}
			catch (Exception ex)
			{
				_libraryReady = false;
				SetFeedback("Could not load the aircraft library: " + ex.Message, true);
			}
			UpdateProvenance();
		}

		void PcrBoxKeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				e.SuppressKeyPress = true;
				Evaluate();
			}
		}

		void SetFeedback(string text, bool isError)
		{
			_pcrFeedback.Text = text;
			_pcrFeedback.ForeColor = isError ? Color.Firebrick : Color.FromArgb(60, 60, 60);
		}

		bool Metric { get { return _metricBox != null && _metricBox.Checked; } }

		string FormatWeight(float lb)
		{
			float v = Metric ? AcrEngine.LbToKg(lb) : lb;
			return v.ToString("#,##0", CultureInfo.InvariantCulture) + (Metric ? " kg" : " lb");
		}

		void Evaluate()
		{
			if (!_libraryReady) return;

			PcrCode pcr;
			string error;
			if (!PcrCode.TryParse(_pcrBox.Text, out pcr, out error))
			{
				SetFeedback(error, true);
				ClearComputedCells();
				return;
			}

			SetFeedback(string.Format(
				"{0} pavement, subgrade {1}, tyre pressure {2}, {3}.   "
				+ "An aircraft may operate when its ACR does not exceed {4:0.##}.",
				pcr.Pavement == PavementKind.Rigid ? "Rigid" : "Flexible",
				pcr.Subgrade, pcr.TyreCategoryText(), pcr.MethodText(), pcr.Value), false);

			// Every row runs the layered-elastic solver a dozen or more times, so
			// this takes a noticeable moment rather than being instant.
			UseWaitCursor = true;
			try
			{
				foreach (DataGridViewRow row in _grid.Rows)
				{
					AircraftSpec spec = row.Tag as AircraftSpec;
					if (spec == null) continue;

					try
					{
						EvaluateRow(row, spec, pcr);
					}
					catch (Exception ex)
					{
						row.Cells[ColVerdict].Value = "Error: " + ex.Message;
						row.Cells[ColVerdict].Style.ForeColor = Color.Firebrick;
					}
				}
			}
			finally
			{
				UseWaitCursor = false;
			}
		}

		void EvaluateRow(DataGridViewRow row, AircraftSpec spec, PcrCode pcr)
		{
			float mtow = _engine.MaxWeightLb(spec);
			float acrAtMtow = _engine.Acr(spec, mtow, pcr.Pavement).For(pcr.Subgrade);

			bool limited;
			float maxWeight = _engine.MaxAllowableWeightLb(spec, pcr, out limited);

			row.Cells[ColMtow].Value = FormatWeight(mtow);
			row.Cells[ColAcrMtow].Value = acrAtMtow.ToString("0.0", CultureInfo.InvariantCulture);

			// Tyre pressure is a separate limit: an aircraft can pass on ACR and
			// still be refused on tyre pressure.
			float psi = _engine.TyrePressurePsi(spec);
			float psiLimit = pcr.TyrePressureLimitPsi();
			bool tyreOk = psi <= psiLimit;
			row.Cells[ColTyres].Value = psi.ToString("0", CultureInfo.InvariantCulture) + " psi"
				+ (tyreOk ? " - ok" : " - EXCEEDS " + psiLimit.ToString("0", CultureInfo.InvariantCulture));
			row.Cells[ColTyres].Style.ForeColor = tyreOk ? Color.FromArgb(30, 110, 40) : Color.Firebrick;

			if (!limited)
			{
				row.Cells[ColMaxWeight].Value = FormatWeight(mtow);
				row.Cells[ColMargin].Value = (pcr.Value - acrAtMtow).ToString("+0.0;-0.0", CultureInfo.InvariantCulture);
				SetVerdict(row, "No pavement limit at MTOW", Color.FromArgb(30, 110, 40));
			}
			else if (maxWeight > 0f)
			{
				row.Cells[ColMaxWeight].Value = FormatWeight(maxWeight);
				row.Cells[ColMargin].Value = "-" + (acrAtMtow - pcr.Value).ToString("0.0", CultureInfo.InvariantCulture) + " at MTOW";
				SetVerdict(row, "Limited to " + FormatWeight(maxWeight), Color.FromArgb(170, 95, 0));
			}
			else
			{
				row.Cells[ColMaxWeight].Value = "-";
				row.Cells[ColMargin].Value = "-";
				SetVerdict(row, "Not usable", Color.Firebrick);
			}

			if (!tyreOk)
				SetVerdict(row, row.Cells[ColVerdict].Value + " / tyre pressure too high", Color.Firebrick);

			RecomputeCheckWeight(row, spec, pcr);
		}

		void SetVerdict(DataGridViewRow row, object text, Color colour)
		{
			row.Cells[ColVerdict].Value = text;
			row.Cells[ColVerdict].Style.ForeColor = colour;
			row.Cells[ColVerdict].Style.Font = new Font(Font, FontStyle.Bold);
		}

		void GridCellEndEdit(object sender, DataGridViewCellEventArgs e)
		{
			if (_grid.Columns[e.ColumnIndex].Name != ColCheckWeight) return;
			if (!_libraryReady) return;

			PcrCode pcr;
			string error;
			if (!PcrCode.TryParse(_pcrBox.Text, out pcr, out error)) return;

			DataGridViewRow row = _grid.Rows[e.RowIndex];
			AircraftSpec spec = row.Tag as AircraftSpec;
			if (spec != null) RecomputeCheckWeight(row, spec, pcr);
		}

		/// <summary>
		/// Reverse direction: the dispatcher types a weight, the tool reports the
		/// ACR at that weight and whether it fits.
		/// </summary>
		void RecomputeCheckWeight(DataGridViewRow row, AircraftSpec spec, PcrCode pcr)
		{
			object raw = row.Cells[ColCheckWeight].Value;
			string text = raw == null ? "" : raw.ToString().Trim();

			if (text.Length == 0)
			{
				row.Cells[ColAcrCheck].Value = "";
				return;
			}

			float entered;
			if (!float.TryParse(text.Replace(" ", "").Replace(",", ""),
			                    NumberStyles.Float, CultureInfo.InvariantCulture, out entered) || entered <= 0f)
			{
				row.Cells[ColAcrCheck].Value = "not a weight";
				row.Cells[ColAcrCheck].Style.ForeColor = Color.Firebrick;
				return;
			}

			float lb = Metric ? AcrEngine.KgToLb(entered) : entered;
			float mtow = _engine.MaxWeightLb(spec);

			if (lb > mtow)
			{
				row.Cells[ColAcrCheck].Value = "above MTOW";
				row.Cells[ColAcrCheck].Style.ForeColor = Color.Firebrick;
				return;
			}

			float acr = _engine.Acr(spec, lb, pcr.Pavement).For(pcr.Subgrade);
			bool fits = acr <= pcr.Value;
			row.Cells[ColAcrCheck].Value = acr.ToString("0.0", CultureInfo.InvariantCulture)
				+ (fits ? "  fits" : "  over");
			row.Cells[ColAcrCheck].Style.ForeColor = fits ? Color.FromArgb(30, 110, 40) : Color.Firebrick;
		}

		void ClearComputedCells()
		{
			foreach (DataGridViewRow row in _grid.Rows)
			{
				row.Cells[ColMtow].Value = "";
				row.Cells[ColAcrMtow].Value = "";
				row.Cells[ColMaxWeight].Value = "";
				row.Cells[ColMargin].Value = "";
				row.Cells[ColTyres].Value = "";
				row.Cells[ColVerdict].Value = "";
				row.Cells[ColAcrCheck].Value = "";
			}
		}

		void RunSelfTest()
		{
			string report;
			bool ok = SelfTest.Run(out report);
			MessageBox.Show(this, report,
				ok ? "Self-test passed" : "Self-test FAILED",
				MessageBoxButtons.OK,
				ok ? MessageBoxIcon.Information : MessageBoxIcon.Error);
		}

		void UpdateProvenance()
		{
			string libVersion = "unknown";
			try
			{
				if (_engine.LibraryPath != null && File.Exists(_engine.LibraryPath))
				{
					// LibraryVersion sits in the opening tag, so a short read is enough.
					using (StreamReader r = new StreamReader(_engine.LibraryPath))
					{
						char[] buf = new char[2048];
						int n = r.Read(buf, 0, buf.Length);
						string head = new string(buf, 0, n);
						int i = head.IndexOf("LibraryVersion=\"", StringComparison.Ordinal);
						if (i >= 0)
						{
							int j = head.IndexOf('"', i + 16);
							if (j > i) libVersion = head.Substring(i + 16, j - i - 16);
						}
					}
				}
			}
			catch { /* provenance only - never block the tool */ }

			_provenance.Text =
				"ACR computed by the FAA ICAO-ACR engine (ACRClassLib.dll); gear geometry from the FAA aircraft library "
				+ "(aircraft.xml, version " + libVersion + "). The 747 freighters are evaluated on the -400 / -400ER "
				+ "entries, which carry the same gear and the freighter weights.\r\n"
				+ "Planning aid only - check against the AIP. Operation above the published PCR is the aerodrome "
				+ "operator's decision, not a calculation, so no overload allowance is applied here.";
		}
	}
}
