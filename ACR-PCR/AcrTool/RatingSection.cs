using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace AcrTool
{
	/// <summary>
	/// One half of the window: a published code to type in, and the four aircraft
	/// evaluated against it.
	///
	/// Both the PCR and the PCN halves are instances of this, so the two are
	/// rendered by identical code and cannot drift into different presentations.
	/// Everything method-specific comes from the IRatingEngine passed in.
	/// </summary>
	public class RatingSection : Panel
	{
		/// <summary>
		/// One row's numbers, computed away from the grid.
		///
		/// The ACR side runs a layered-elastic solver, which is far too slow to sit
		/// on the UI thread, so the arithmetic happens on a worker and only this
		/// plain data crosses back to be rendered.
		/// </summary>
		class RowData
		{
			public AircraftSpec Spec;
			public float MinWeight, MaxWeight, AtMin, AtMax, Allowed, Psi, TyreLimit;
			public bool Limited, TyreOk;
			public string Error;
		}

		readonly IRatingEngine _engine;
		readonly List<AircraftSpec> _fleet;

		TextBox _codeBox;
		Label _feedback;
		CheckBox _overloadBox;
		ProgressBar _progress;
		Button _go;
		DataGridView _grid;

		BackgroundWorker _worker;
		bool _rerunWanted;

		const string ColAircraft = "aircraft";
		const string ColLibrary = "library";
		const string ColMinWeight = "minweight";
		const string ColAtMin = "atmin";
		const string ColMaxWeight = "maxweight";
		const string ColAtMax = "atmax";
		const string ColAllowed = "allowed";
		const string ColMargin = "margin";
		const string ColTyres = "tyres";
		const string ColVerdict = "verdict";
		const string ColCheckWeight = "checkweight";
		const string ColAtCheck = "atcheck";

		/// <summary>Set by the form; both halves share one unit setting.</summary>
		public bool Metric = true;

		public IRatingEngine Engine { get { return _engine; } }

		public RatingSection(IRatingEngine engine, string title, string placeholder, List<AircraftSpec> fleet)
		{
			_engine = engine;
			_fleet = fleet;

			BuildFeedback();
			BuildTopBar(title, placeholder);
			BuildGrid();
		}

		void BuildFeedback()
		{
			Panel p = new Panel();
			p.Dock = DockStyle.Top;
			p.Height = 30;
			p.Padding = new Padding(10, 2, 10, 2);

			_feedback = new Label();
			_feedback.Dock = DockStyle.Fill;
			_feedback.AutoEllipsis = true;
			p.Controls.Add(_feedback);

			Controls.Add(p);
		}

		/// <summary>
		/// Laid out with a FlowLayoutPanel rather than fixed coordinates: the title
		/// is wider than the gap it was given and sat on top of the text box. Flow
		/// layout cannot overlap, whatever the font or DPI turns out to be.
		/// </summary>
		void BuildTopBar(string title, string placeholder)
		{
			FlowLayoutPanel flow = new FlowLayoutPanel();
			flow.Dock = DockStyle.Top;
			flow.Height = 36;
			flow.WrapContents = false;
			flow.FlowDirection = FlowDirection.LeftToRight;
			flow.Padding = new Padding(8, 4, 8, 0);

			Label lbl = new Label();
			lbl.Text = title;
			lbl.AutoSize = true;
			lbl.Font = new Font(Font.FontFamily, 9.5f, FontStyle.Bold);
			lbl.Margin = new Padding(2, 7, 10, 0);
			flow.Controls.Add(lbl);

			_codeBox = new TextBox();
			_codeBox.Width = 160;
			_codeBox.Font = new Font("Consolas", 11f, FontStyle.Bold);
			_codeBox.Text = placeholder;
			_codeBox.Margin = new Padding(0, 3, 8, 0);
			_codeBox.KeyDown += CodeBoxKeyDown;
			flow.Controls.Add(_codeBox);

			_go = new Button();
			_go.Text = "Evaluate";
			_go.Size = new Size(84, 25);
			_go.Margin = new Padding(0, 3, 14, 0);
			_go.Click += delegate { Evaluate(); };
			flow.Controls.Add(_go);

			// Overload is never applied unless this is ticked: allowing movements
			// above the published rating is the aerodrome operator's call, and the
			// ICAO criteria carry conditions this tool cannot check.
			_overloadBox = new CheckBox();
			_overloadBox.Text = "Allow ICAO overload (+10% flex / +5% rigid)";
			_overloadBox.AutoSize = true;
			_overloadBox.Margin = new Padding(0, 7, 14, 0);
			_overloadBox.ForeColor = Color.FromArgb(170, 95, 0);
			_overloadBox.CheckedChanged += delegate { Evaluate(); };
			flow.Controls.Add(_overloadBox);

			_progress = new ProgressBar();
			_progress.Style = ProgressBarStyle.Blocks;
			_progress.Size = new Size(130, 16);
			_progress.Margin = new Padding(0, 8, 0, 0);
			_progress.Visible = false;
			flow.Controls.Add(_progress);

			Controls.Add(flow);
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

			string r = _engine.RatingName;
			AddColumn(ColAircraft, "Aircraft", true, 88);
			AddColumn(ColLibrary, "Source entry", true, 92);
			AddColumn(ColMinWeight, "Min weight", true, 88);
			AddColumn(ColAtMin, r + " at min", true, 74);
			AddColumn(ColMaxWeight, "Max weight", true, 88);
			AddColumn(ColAtMax, r + " at max", true, 74);
			AddColumn(ColAllowed, "Max weight on this code", true, 112);
			AddColumn(ColMargin, "Margin", true, 80);
			AddColumn(ColTyres, "Tyres", true, 100);
			AddColumn(ColVerdict, "Verdict", true, 140);
			AddColumn(ColCheckWeight, "Check a weight", false, 88);
			AddColumn(ColAtCheck, r + " at that weight", true, 100);

			_grid.CellEndEdit += GridCellEndEdit;

			foreach (AircraftSpec spec in _fleet)
			{
				int i = _grid.Rows.Add();
				_grid.Rows[i].Tag = spec;
				_grid.Rows[i].Cells[ColAircraft].Value = spec.Display;
				_grid.Rows[i].Cells[ColLibrary].Value =
					_engine.Method == RatingMethod.Acr ? spec.LibraryName : spec.Display;
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

		void CodeBoxKeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				e.SuppressKeyPress = true;
				Evaluate();
			}
		}

		void SetFeedback(string text, bool isError)
		{
			_feedback.Text = text;
			_feedback.ForeColor = isError ? Color.Firebrick : Color.FromArgb(60, 60, 60);
		}

		string FormatWeight(float lb)
		{
			float v = Metric ? AcrEngine.LbToKg(lb) : lb;
			return v.ToString("#,##0", CultureInfo.InvariantCulture) + (Metric ? " kg" : " lb");
		}

		public void Evaluate()
		{
			if (!_engine.Ready)
			{
				SetFeedback(_engine.NotReadyReason, true);
				ClearComputedCells();
				return;
			}

			PavementCode code;
			string error;
			if (!PavementCode.TryParse(_codeBox.Text, _engine.Method, out code, out error))
			{
				SetFeedback(error, true);
				ClearComputedCells();
				return;
			}

			// A run already in flight finishes first, then repeats with whatever the
			// controls say by then - so a quick flurry of clicks cannot pile up
			// several solver runs on top of each other.
			if (_worker != null && _worker.IsBusy)
			{
				_rerunWanted = true;
				return;
			}

			bool overload = _overloadBox.Checked;
			float limit = code.EffectiveValue(overload);

			string basis = overload
				? string.Format(CultureInfo.InvariantCulture,
					"   Overload {0} applied: limit used is {1:0.##}, not the published {2:0.##}.",
					code.OverloadText(), limit, code.Value)
				: "";

			SetFeedback(string.Format(CultureInfo.InvariantCulture,
				"{0} pavement, {1}, tyre pressure {2}, {3}.   An aircraft may operate when its {4} does not exceed {5:0.##}.{6}",
				code.Pavement == PavementKind.Rigid ? "Rigid" : "Flexible",
				code.SubgradeText(), code.TyreCategoryText(), code.EvaluationText(),
				_engine.RatingName, code.Value, basis), false);

			if (overload) _feedback.ForeColor = Color.FromArgb(170, 95, 0);

			StartRun(code, limit, overload);
		}

		void StartRun(PavementCode code, float limit, bool overload)
		{
			_progress.Maximum = _fleet.Count;
			_progress.Value = 0;
			_progress.Visible = true;
			_go.Enabled = false;

			// Only ever reached when the previous run has finished, so disposing it
			// here cannot pull the rug from under a live one.
			if (_worker != null) _worker.Dispose();
			_worker = new BackgroundWorker();
			_worker.WorkerReportsProgress = true;

			_worker.DoWork += delegate(object s, DoWorkEventArgs e)
			{
				BackgroundWorker bw = (BackgroundWorker)s;
				for (int i = 0; i < _fleet.Count; i++)
				{
					// Reported one at a time so each aircraft appears as soon as it
					// is solved, rather than the grid staying blank until all four
					// are done.
					RowData d = Compute(_fleet[i], code, limit);
					bw.ReportProgress(i + 1, d);
				}
			};

			_worker.ProgressChanged += delegate(object s, ProgressChangedEventArgs e)
			{
				_progress.Value = Math.Min(e.ProgressPercentage, _progress.Maximum);

				int index = e.ProgressPercentage - 1;
				RowData d = e.UserState as RowData;
				if (d != null && index >= 0 && index < _grid.Rows.Count)
					Render(_grid.Rows[index], d, code, limit, overload);
			};

			_worker.RunWorkerCompleted += delegate(object s, RunWorkerCompletedEventArgs e)
			{
				_progress.Visible = false;
				_go.Enabled = true;

				if (e.Error != null)
				{
					SetFeedback("Evaluation failed: " + e.Error.Message, true);
					ClearComputedCells();
				}
				else if (_engine.Method == RatingMethod.Acr)
				{
					// Rows were already filled in as they completed. Report what the
					// solver actually cost, on screen and in a log next to the exe -
					// guesswork about where the time goes has not served us well.
					AcrEngine.DumpTiming(System.IO.Path.Combine(
						Application.StartupPath, "acr_timing.txt"));
					_feedback.Text += "   [" + AcrEngine.TimingSummary() + "]";
				}

				if (_rerunWanted)
				{
					_rerunWanted = false;
					Evaluate();
				}
			};

			_worker.RunWorkerAsync();
		}

		/// <summary>Runs on the worker thread - must not touch any control.</summary>
		RowData Compute(AircraftSpec spec, PavementCode code, float limit)
		{
			RowData d = new RowData();
			d.Spec = spec;
			try
			{
				d.MaxWeight = _engine.MaxWeightLb(spec);
				d.MinWeight = _engine.MinWeightLb(spec);
				d.AtMax = _engine.Rating(spec, d.MaxWeight, code);
				d.AtMin = d.MinWeight > 0f ? _engine.Rating(spec, d.MinWeight, code) : 0f;
				d.Allowed = _engine.MaxAllowableWeightLb(spec, code, limit, out d.Limited);

				d.Psi = _engine.TyrePressurePsi(spec);
				d.TyreLimit = code.TyrePressureLimitPsi();
				d.TyreOk = d.Psi <= d.TyreLimit;
			}
			catch (Exception ex)
			{
				d.Error = ex.Message;
			}
			return d;
		}

		void Render(DataGridViewRow row, RowData d, PavementCode code, float limit, bool overload)
		{
			if (d.Error != null)
			{
				ClearRow(row);
				row.Cells[ColVerdict].Value = d.Error;
				row.Cells[ColVerdict].Style.ForeColor = Color.Firebrick;
				return;
			}

			// A dash rather than a zero where no minimum weight is published - the
			// ACR side computes at any weight and has no empty-weight point.
			bool hasMin = d.MinWeight > 0f;
			row.Cells[ColMinWeight].Value = hasMin ? FormatWeight(d.MinWeight) : "-";
			row.Cells[ColAtMin].Value = hasMin
				? d.AtMin.ToString("0.0", CultureInfo.InvariantCulture) : "-";

			row.Cells[ColMaxWeight].Value = FormatWeight(d.MaxWeight);
			row.Cells[ColAtMax].Value = d.AtMax.ToString("0.0", CultureInfo.InvariantCulture);

			// Tyre pressure is a separate limit: an aircraft can pass on the
			// rating and still be refused on tyre pressure.
			row.Cells[ColTyres].Value = d.Psi.ToString("0", CultureInfo.InvariantCulture) + " psi"
				+ (d.TyreOk ? " - ok" : " - EXCEEDS " + d.TyreLimit.ToString("0", CultureInfo.InvariantCulture));
			row.Cells[ColTyres].Style.ForeColor = d.TyreOk ? Color.FromArgb(30, 110, 40) : Color.Firebrick;

			string tag = overload ? " (overload " + code.OverloadText() + ")" : "";

			if (!d.Limited)
			{
				row.Cells[ColAllowed].Value = FormatWeight(d.MaxWeight);
				row.Cells[ColMargin].Value = (limit - d.AtMax).ToString("+0.0;-0.0", CultureInfo.InvariantCulture);
				SetVerdict(row, "No pavement limit at max weight" + tag,
					overload ? Color.FromArgb(170, 95, 0) : Color.FromArgb(30, 110, 40));
			}
			else if (d.Allowed > 0f)
			{
				row.Cells[ColAllowed].Value = FormatWeight(d.Allowed);
				row.Cells[ColMargin].Value = "-" + (d.AtMax - limit).ToString("0.0", CultureInfo.InvariantCulture) + " at max";
				SetVerdict(row, "Limited to " + FormatWeight(d.Allowed) + tag, Color.FromArgb(170, 95, 0));
			}
			else
			{
				row.Cells[ColAllowed].Value = "-";
				row.Cells[ColMargin].Value = "-";
				SetVerdict(row, "Not usable" + tag, Color.Firebrick);
			}

			if (!d.TyreOk)
				SetVerdict(row, row.Cells[ColVerdict].Value + " / tyre pressure too high", Color.Firebrick);

			RecomputeCheckWeight(row, d.Spec, code, limit);
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
			if (!_engine.Ready) return;

			PavementCode code;
			string error;
			if (!PavementCode.TryParse(_codeBox.Text, _engine.Method, out code, out error)) return;

			DataGridViewRow row = _grid.Rows[e.RowIndex];
			AircraftSpec spec = row.Tag as AircraftSpec;
			if (spec != null)
			{
				try { RecomputeCheckWeight(row, spec, code, code.EffectiveValue(_overloadBox.Checked)); }
				catch (Exception ex)
				{
					row.Cells[ColAtCheck].Value = ex.Message;
					row.Cells[ColAtCheck].Style.ForeColor = Color.Firebrick;
				}
			}
		}

		/// <summary>
		/// Reverse direction: type a weight, get the rating at it.
		///
		/// This one stays on the UI thread - it is a single evaluation, and on the
		/// ACR side the weights involved are usually already in the engine's cache.
		/// </summary>
		void RecomputeCheckWeight(DataGridViewRow row, AircraftSpec spec, PavementCode code, float limit)
		{
			object raw = row.Cells[ColCheckWeight].Value;
			string text = raw == null ? "" : raw.ToString().Trim();

			if (text.Length == 0)
			{
				row.Cells[ColAtCheck].Value = "";
				return;
			}

			float entered;
			if (!float.TryParse(text.Replace(" ", "").Replace(",", ""),
			                    NumberStyles.Float, CultureInfo.InvariantCulture, out entered) || entered <= 0f)
			{
				row.Cells[ColAtCheck].Value = "not a weight";
				row.Cells[ColAtCheck].Style.ForeColor = Color.Firebrick;
				return;
			}

			float lb = Metric ? AcrEngine.KgToLb(entered) : entered;
			float maxW = _engine.MaxWeightLb(spec);

			if (lb > maxW)
			{
				row.Cells[ColAtCheck].Value = "above max weight";
				row.Cells[ColAtCheck].Style.ForeColor = Color.Firebrick;
				return;
			}

			float v = _engine.Rating(spec, lb, code);
			bool fits = v <= limit;
			row.Cells[ColAtCheck].Value = v.ToString("0.0", CultureInfo.InvariantCulture) + (fits ? "  fits" : "  over");
			row.Cells[ColAtCheck].Style.ForeColor = fits ? Color.FromArgb(30, 110, 40) : Color.Firebrick;
		}

		void ClearRow(DataGridViewRow row)
		{
			foreach (string c in new[] { ColMinWeight, ColAtMin, ColMaxWeight, ColAtMax,
			                             ColAllowed, ColMargin, ColTyres, ColVerdict, ColAtCheck })
				row.Cells[c].Value = "";
		}

		void ClearComputedCells()
		{
			foreach (DataGridViewRow row in _grid.Rows) ClearRow(row);
		}
	}
}
