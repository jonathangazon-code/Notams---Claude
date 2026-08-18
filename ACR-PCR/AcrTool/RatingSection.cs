using System;
using System.Collections.Generic;
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
		readonly IRatingEngine _engine;
		readonly List<AircraftSpec> _fleet;

		TextBox _codeBox;
		Label _feedback;
		CheckBox _overloadBox;
		DataGridView _grid;

		const string ColAircraft = "aircraft";
		const string ColLibrary = "library";
		const string ColMtow = "mtow";
		const string ColAtMtow = "atmtow";
		const string ColMaxWeight = "maxweight";
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

			BuildTopBar(title, placeholder);
			BuildGrid();
		}

		void BuildTopBar(string title, string placeholder)
		{
			Panel bar = new Panel();
			bar.Dock = DockStyle.Top;
			bar.Height = 62;

			Label lbl = new Label();
			lbl.Text = title;
			lbl.AutoSize = true;
			lbl.Font = new Font(Font.FontFamily, 9.5f, FontStyle.Bold);
			lbl.Location = new Point(10, 10);
			bar.Controls.Add(lbl);

			_codeBox = new TextBox();
			_codeBox.Location = new Point(190, 6);
			_codeBox.Width = 170;
			_codeBox.Font = new Font("Consolas", 11f, FontStyle.Bold);
			_codeBox.Text = placeholder;
			_codeBox.KeyDown += CodeBoxKeyDown;
			bar.Controls.Add(_codeBox);

			Button go = new Button();
			go.Text = "Evaluate";
			go.Location = new Point(372, 5);
			go.Size = new Size(90, 27);
			go.Click += delegate { Evaluate(); };
			bar.Controls.Add(go);

			// Overload is never applied unless this is ticked: allowing movements
			// above the published rating is the aerodrome operator's call, and the
			// ICAO criteria carry conditions this tool cannot check.
			_overloadBox = new CheckBox();
			_overloadBox.Text = "Allow ICAO overload tolerance (+10% flexible / +5% rigid)";
			_overloadBox.AutoSize = true;
			_overloadBox.Location = new Point(478, 10);
			_overloadBox.ForeColor = Color.FromArgb(170, 95, 0);
			_overloadBox.CheckedChanged += delegate { Evaluate(); };
			bar.Controls.Add(_overloadBox);

			_feedback = new Label();
			_feedback.AutoSize = true;
			_feedback.Location = new Point(10, 40);
			_feedback.MaximumSize = new Size(1400, 0);
			bar.Controls.Add(_feedback);

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

			string r = _engine.RatingName;
			AddColumn(ColAircraft, "Aircraft", true, 90);
			AddColumn(ColLibrary, "Source entry", true, 100);
			AddColumn(ColMtow, "Max weight", true, 90);
			AddColumn(ColAtMtow, r + " at max", true, 85);
			AddColumn(ColMaxWeight, "Max weight on this code", true, 115);
			AddColumn(ColMargin, "Margin", true, 85);
			AddColumn(ColTyres, "Tyres", true, 105);
			AddColumn(ColVerdict, "Verdict", true, 145);
			AddColumn(ColCheckWeight, "Check a weight", false, 90);
			AddColumn(ColAtCheck, r + " at that weight", true, 105);

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

			UseWaitCursor = true;
			try
			{
				foreach (DataGridViewRow row in _grid.Rows)
				{
					AircraftSpec spec = row.Tag as AircraftSpec;
					if (spec == null) continue;

					try
					{
						EvaluateRow(row, spec, code, limit, overload);
					}
					catch (Exception ex)
					{
						ClearRow(row);
						row.Cells[ColVerdict].Value = ex.Message;
						row.Cells[ColVerdict].Style.ForeColor = Color.Firebrick;
					}
				}
			}
			finally
			{
				UseWaitCursor = false;
			}
		}

		void EvaluateRow(DataGridViewRow row, AircraftSpec spec, PavementCode code, float limit, bool overload)
		{
			float maxW = _engine.MaxWeightLb(spec);
			float atMax = _engine.Rating(spec, maxW, code);

			bool limited;
			float allowed = _engine.MaxAllowableWeightLb(spec, code, limit, out limited);

			row.Cells[ColMtow].Value = FormatWeight(maxW);
			row.Cells[ColAtMtow].Value = atMax.ToString("0.0", CultureInfo.InvariantCulture);

			// Tyre pressure is a separate limit: an aircraft can pass on the
			// rating and still be refused on tyre pressure.
			float psi = _engine.TyrePressurePsi(spec);
			float limit = code.TyrePressureLimitPsi();
			bool tyreOk = psi <= limit;
			row.Cells[ColTyres].Value = psi.ToString("0", CultureInfo.InvariantCulture) + " psi"
				+ (tyreOk ? " - ok" : " - EXCEEDS " + limit.ToString("0", CultureInfo.InvariantCulture));
			row.Cells[ColTyres].Style.ForeColor = tyreOk ? Color.FromArgb(30, 110, 40) : Color.Firebrick;

			if (!limited)
			{
				row.Cells[ColMaxWeight].Value = FormatWeight(maxW);
				row.Cells[ColMargin].Value = (limit - atMax).ToString("+0.0;-0.0", CultureInfo.InvariantCulture);
				SetVerdict(row, "No pavement limit at max weight" + (overload ? " (overload " + code.OverloadText() + ")" : ""),
					overload ? Color.FromArgb(170, 95, 0) : Color.FromArgb(30, 110, 40));
			}
			else if (allowed > 0f)
			{
				row.Cells[ColMaxWeight].Value = FormatWeight(allowed);
				row.Cells[ColMargin].Value = "-" + (atMax - limit).ToString("0.0", CultureInfo.InvariantCulture) + " at max";
				SetVerdict(row, "Limited to " + FormatWeight(allowed) + (overload ? " (overload " + code.OverloadText() + ")" : ""),
					Color.FromArgb(170, 95, 0));
			}
			else
			{
				row.Cells[ColMaxWeight].Value = "-";
				row.Cells[ColMargin].Value = "-";
				SetVerdict(row, "Not usable", Color.Firebrick);
			}

			if (!tyreOk)
				SetVerdict(row, row.Cells[ColVerdict].Value + " / tyre pressure too high", Color.Firebrick);

			RecomputeCheckWeight(row, spec, code, limit);
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

		/// <summary>Reverse direction: type a weight, get the rating at it.</summary>
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
			foreach (string c in new[] { ColMtow, ColAtMtow, ColMaxWeight, ColMargin, ColTyres, ColVerdict, ColAtCheck })
				row.Cells[c].Value = "";
		}

		void ClearComputedCells()
		{
			foreach (DataGridViewRow row in _grid.Rows) ClearRow(row);
		}
	}
}
