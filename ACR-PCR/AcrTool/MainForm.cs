using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace AcrTool
{
	/// <summary>
	/// Two halves, same presentation: the current ACR/PCR method on top, the
	/// legacy ACN/PCN method below. Both are RatingSection instances.
	///
	/// Built in code rather than through the WinForms designer, so the layout
	/// stays readable in a diff and there is no .resx to keep in step.
	/// </summary>
	public class MainForm : Form
	{
		readonly AcrEngine _acr = new AcrEngine();
		readonly AcnEngine _acn = new AcnEngine();
		readonly List<AircraftSpec> _fleet = AircraftSpec.Fleet();

		RatingSection _acrSection;
		RatingSection _acnSection;
		CheckBox _metricBox;
		Label _provenance;

		public MainForm()
		{
			Text = "ACR / PCR and ACN / PCN - maximum weight check";
			StartPosition = FormStartPosition.CenterScreen;
			ClientSize = new Size(1200, 780);
			MinimumSize = new Size(940, 560);
			Font = new Font("Segoe UI", 9f);

			BuildAppBar();
			BuildFooter();
			BuildSections();

			Shown += MainFormShown;
		}

		void BuildAppBar()
		{
			Panel bar = new Panel();
			bar.Dock = DockStyle.Top;
			bar.Height = 38;

			_metricBox = new CheckBox();
			_metricBox.Text = "Show weights in kg";
			_metricBox.Checked = true;
			_metricBox.AutoSize = true;
			_metricBox.Location = new Point(12, 10);
			_metricBox.CheckedChanged += delegate
			{
				_acrSection.Metric = _metricBox.Checked;
				_acnSection.Metric = _metricBox.Checked;
				_acrSection.Evaluate();
				_acnSection.Evaluate();
			};
			bar.Controls.Add(_metricBox);

			Button selfTest = new Button();
			selfTest.Text = "Self-test (ACR)";
			selfTest.Location = new Point(160, 6);
			selfTest.Size = new Size(120, 26);
			selfTest.Click += delegate { RunSelfTest(); };
			bar.Controls.Add(selfTest);

			Controls.Add(bar);
		}

		void BuildFooter()
		{
			Panel foot = new Panel();
			foot.Dock = DockStyle.Bottom;
			foot.Height = 78;
			foot.Padding = new Padding(12, 6, 12, 6);

			_provenance = new Label();
			_provenance.Dock = DockStyle.Fill;
			_provenance.ForeColor = Color.FromArgb(90, 90, 90);
			foot.Controls.Add(_provenance);

			Controls.Add(foot);
		}

		void BuildSections()
		{
			SplitContainer split = new SplitContainer();
			split.Dock = DockStyle.Fill;
			split.Orientation = Orientation.Horizontal;
			split.SplitterWidth = 6;

			_acrSection = new RatingSection(_acr, "Published PCR   (current method)", "690/R/B/W/T", _fleet);
			_acrSection.Dock = DockStyle.Fill;
			split.Panel1.Controls.Add(_acrSection);

			_acnSection = new RatingSection(_acn, "Published PCN   (legacy method)", "80/R/B/W/T", _fleet);
			_acnSection.Dock = DockStyle.Fill;
			split.Panel2.Controls.Add(_acnSection);

			Controls.Add(split);
			split.BringToFront();

			// SplitterDistance must be set once the control has a real height,
			// otherwise it is clamped against the design-time size.
			split.HandleCreated += delegate
			{
				try { split.SplitterDistance = Math.Max(160, split.Height / 2); }
				catch { /* ignored - only affects the initial split position */ }
			};
		}

		void MainFormShown(object sender, EventArgs e)
		{
			try { _acr.Load(Path.Combine(Application.StartupPath, "aircraft.xml")); }
			catch (Exception ex) { ShowLoadProblem("aircraft.xml", ex); }

			_acn.Load(Path.Combine(Application.StartupPath, "acn-data.xml"));

			_acrSection.Metric = _metricBox.Checked;
			_acnSection.Metric = _metricBox.Checked;
			_acrSection.Evaluate();
			_acnSection.Evaluate();

			UpdateProvenance();
		}

		void ShowLoadProblem(string what, Exception ex)
		{
			MessageBox.Show(this,
				"Could not load " + what + "." + Environment.NewLine + Environment.NewLine + ex.Message,
				"ACR / PCR", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
			_provenance.Text =
				_acr.Provenance + Environment.NewLine
				+ _acn.Provenance + Environment.NewLine
				+ "The 747 freighters are evaluated on the -400 / -400ER entries, which carry the same gear and the "
				+ "freighter weights. Planning aid only - check against the AIP. Operation above the published PCR or "
				+ "PCN is the aerodrome operator's decision, not a calculation, so no overload allowance is applied.";
		}
	}
}
